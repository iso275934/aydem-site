"""
Aydem sitesi - veri güncelleme programı
========================================

GitHub Actions bunu her sabah çalıştırır (elle de çalıştırabilirsin):

    py update_data.py            # sadece eksik/yeni günleri çeker (hızlı)
    py update_data.py --full     # 1 Ekim 2025'ten itibaren HER ŞEYİ baştan çeker (yavaş)

Ne yapar?
  1) config.json'u okur (santral listesi, YEKDEM fiyatları, kurlar)
  2) data/raw_daily.csv içindeki eski veriyi yükler; sadece eksik günleri
     (ve son 3 günü, EPİAŞ düzeltme yapmış olabilir diye) EPİAŞ'tan çeker
  3) Ciroyu hesaplar: YEKDEM'li santral -> üretim x fiyat x kur,
     diğerleri -> saatlik üretim x saatlik PTF
  4) docs/data.json (site bunu okur) ve docs/aydem_gunluk_karsilastirma.xlsx üretir

Not: Ham veride YEKDEM santrallerin de PTF'li cirosu saklanır; bu sayede
config.json'da kur veya YEKDEM fiyatı değiştirirsen veriyi tekrar
çekmeye gerek kalmaz, sadece hesap yenilenir.
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd

import aydem_core as core

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config.json"
RAW_CSV = ROOT / "data" / "raw_daily.csv"
DOCS = ROOT / "docs"
DATA_JSON = DOCS / "data.json"
XLSX_PATH = DOCS / "aydem_gunluk_karsilastirma.xlsx"

GAP_JSON = ROOT / "data" / "eksik.json"   # çekilemeyen en eski gün (bir sonraki çalışma buradan devam eder)
ABORT_AFTER = 5         # üst üste bu kadar bağlantı hatasında EPİAŞ'a erişilemiyor say, dur
OVERLAP_DAYS = 3        # her çalışmada son kaç gün yeniden çekilsin (EPİAŞ düzeltmeleri için)
RECENT_CHECK_DAYS = 45  # eksik santral kontrolü sadece son kaç gün için yapılsın
TZ = ZoneInfo("Europe/Istanbul")


# ----------------------------------------------------------------------
# Ayarlar
# ----------------------------------------------------------------------

def load_config() -> dict:
    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    core.AYDEM_PLANT_NAMES = list(cfg["santraller"])
    core.YEKDEM_PLANTS = {k: float(v) for k, v in cfg["yekdem"].items()}
    core.USDTRY_RATE_BY_YEAR = {int(k): float(v) for k, v in cfg["kur_yillik"].items()}
    core.USDTRY_RATE_BY_MONTH = {k: float(v) for k, v in cfg.get("kur_aylik", {}).items()}
    return cfg


# ----------------------------------------------------------------------
# Ham veri (EPİAŞ) - artımlı
# ----------------------------------------------------------------------

def load_raw() -> pd.DataFrame:
    if not RAW_CSV.exists():
        return pd.DataFrame(columns=["tarih", "santral", "uretim_mwh", "ciro_ptf_tl"])
    df = pd.read_csv(RAW_CSV, encoding="utf-8")
    df["tarih"] = pd.to_datetime(df["tarih"])
    return df


def incomplete_days(raw: pd.DataFrame, start: date, end: date, n_plants: int) -> list:
    """Veride tamamen eksik, PTF'si boş veya (son günlerde) bazı santralleri eksik günler."""
    expected = pd.date_range(start, end, freq="D")
    if raw.empty:
        return list(expected)

    g = raw.groupby("tarih")
    n_rows = g.size()
    has_nan = g["ciro_ptf_tl"].apply(lambda s: s.isna().any())

    bad = []
    recent_from = pd.Timestamp(end) - pd.Timedelta(days=RECENT_CHECK_DAYS)
    for d in expected:
        if d not in n_rows.index:
            bad.append(d)
        elif has_nan.get(d, False):
            bad.append(d)
        elif d >= recent_from and n_rows[d] < n_plants:
            bad.append(d)
    return bad


def fetch_raw(eptr, plants: pd.DataFrame, start: date, end: date):
    """start..end arası günlük ham veri: üretim (MWh) ve PTF'li ciro (TL). Aylık parçalarla çeker.

    Üst üste ABORT_AFTER bağlantı hatası olursa durur (EPİAŞ'a erişilemiyor demektir); o ana kadar
    çekilen veri geri döner, çekilemeyen en eski gün `gap_start` olarak bildirilir.
    Döner: (veri, uyarılar, gap_start, durdu_mu)
    """
    id_col = core.get_pp_id_column(plants)
    name_col = "name" if "name" in plants.columns else plants.columns[0]
    chunks = core.month_chunks(start, end)

    rows, failures = [], []
    state = {"gap": None, "consec": 0}

    def mark_gap(d):
        state["gap"] = d if state["gap"] is None else min(state["gap"], d)

    aborted = False
    for ci, (cs, ce) in enumerate(chunks, 1):
        s, e = cs.isoformat(), ce.isoformat()
        print(f"\n=== [{ci}/{len(chunks)}] {s} -> {e} ===", flush=True)
        try:
            ptf = core.fetch_ptf(eptr, s, e)
            state["consec"] = 0
        except Exception as exc:  # noqa: BLE001
            print(f"  HATA: PTF çekilemedi ({str(exc)[:160]})")
            failures.append(f"PTF {s} → {e}")
            mark_gap(cs)
            state["consec"] += 1
            ptf = None
            if state["consec"] >= ABORT_AFTER:
                aborted = True
                break

        for _, p in plants.iterrows():
            pname, pp_id = p[name_col], p[id_col]
            print(f"  -> {pname}", flush=True)
            try:
                prod = core.fetch_hourly_production(eptr, pp_id, s, e)
                state["consec"] = 0
            except Exception as exc:  # noqa: BLE001
                print(f"     HATA: {str(exc)[:160]}")
                failures.append(f"{pname} {s} → {e}")
                mark_gap(cs)
                state["consec"] += 1
                if state["consec"] >= ABORT_AFTER:
                    aborted = True
                    break
                continue
            if prod.empty:
                failures.append(f"{pname} {s} → {e} (veri boş)")
                mark_gap(cs)
                continue

            if ptf is not None:
                m = prod.merge(ptf, on="tarih_saat", how="left")
                m["ciro_ptf_tl"] = m["uretim_mwh"] * m["ptf_tl_mwh"].fillna(0)
            else:
                m = prod.copy()
                m["ciro_ptf_tl"] = np.nan

            m["tarih"] = m["tarih_saat"].dt.normalize()
            daily = m.groupby("tarih", as_index=False)[["uretim_mwh", "ciro_ptf_tl"]].sum(min_count=1)
            daily["santral"] = pname
            rows.append(daily)
        if aborted:
            break

    if aborted:
        print(
            f"\nÜst üste {ABORT_AFTER} bağlantı hatası: EPİAŞ'a şu an erişilemiyor. "
            "Şimdiye kadar çekilen veri kaydedilecek, bir sonraki çalışma kaldığı yerden devam eder."
        )

    if not rows:
        empty = pd.DataFrame(columns=["tarih", "santral", "uretim_mwh", "ciro_ptf_tl"])
        return empty, failures, state["gap"], aborted

    new = pd.concat(rows, ignore_index=True)
    new = new.groupby(["tarih", "santral"], as_index=False)[["uretim_mwh", "ciro_ptf_tl"]].sum(min_count=1)
    return new, failures, state["gap"], aborted


def load_gap():
    if GAP_JSON.exists():
        try:
            return date.fromisoformat(json.loads(GAP_JSON.read_text(encoding="utf-8"))["baslangic"])
        except Exception:  # noqa: BLE001
            return None
    return None


def save_gap(gap_start):
    if gap_start is None:
        if GAP_JSON.exists():
            GAP_JSON.unlink()
        return
    GAP_JSON.parent.mkdir(parents=True, exist_ok=True)
    GAP_JSON.write_text(json.dumps({"baslangic": gap_start.isoformat()}), encoding="utf-8")


# ----------------------------------------------------------------------
# Hesaplama
# ----------------------------------------------------------------------

def build_detail(raw: pd.DataFrame) -> pd.DataFrame:
    """Ham veriden santral bazında mekanizmaya göre ciroyu hesaplar."""
    d = raw.copy()
    d["tarih"] = pd.to_datetime(d["tarih"])

    price = d["santral"].map(core.yekdem_price_for).astype(float)  # YEKDEM değilse NaN
    is_yek = price.notna()
    d["mekanizma"] = np.where(is_yek, "YEKDEM", "PTF")

    kur = d["tarih"].map(lambda t: core.usdtry_for(t.date()))
    kur = pd.to_numeric(kur, errors="coerce")
    if (is_yek & kur.isna()).any():
        eksik = sorted(d.loc[is_yek & kur.isna(), "tarih"].dt.strftime("%Y-%m").unique())
        sys.exit(f"HATA: config.json'da şu aylar için USD/TRY kuru yok: {eksik}")

    d["ciro_tl"] = d["ciro_ptf_tl"]
    d.loc[is_yek, "ciro_tl"] = d.loc[is_yek, "uretim_mwh"] * (price[is_yek] / 100) * 1000 * kur[is_yek]
    return d[["tarih", "santral", "mekanizma", "uretim_mwh", "ciro_tl"]]


def build_plant_table(window: pd.DataFrame, this_year: int) -> list:
    ly = this_year - 1
    g = window.groupby(["santral", "mekanizma", "yil"])[["uretim_mwh", "ciro_tl"]].sum().unstack("yil")
    out = []
    for (santral, mek), r in g.iterrows():
        def val(metric, y):
            v = r.get((metric, y), np.nan)
            return None if pd.isna(v) else float(v)

        def chg(a, b):
            return None if (a is None or b is None or b == 0) else a / b - 1

        u_p, u_n = val("uretim_mwh", ly), val("uretim_mwh", this_year)
        c_p, c_n = val("ciro_tl", ly), val("ciro_tl", this_year)
        out.append(
            {
                "santral": santral,
                "mekanizma": mek,
                "uretim_onceki": u_p, "uretim_simdiki": u_n, "uretim_degisim": chg(u_n, u_p),
                "ciro_onceki": c_p, "ciro_simdiki": c_n, "ciro_degisim": chg(c_n, c_p),
            }
        )
    out.sort(key=lambda x: -(x["ciro_simdiki"] or 0))
    return out


def sanitize(obj):
    """NaN/inf -> None, numpy -> python (JSON için)."""
    if isinstance(obj, dict):
        return {str(k): sanitize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (float, np.floating)):
        f = float(obj)
        return None if (math.isnan(f) or math.isinf(f)) else f
    if isinstance(obj, (pd.Timestamp, datetime, date)):
        return obj.isoformat()
    return obj


def build_json(cfg, comp, summary, avg, plant_tbl, this_year, end_day, start_day, failures) -> dict:
    ly = this_year - 1

    gunluk = []
    for _, r in comp.iterrows():
        gunluk.append(
            {
                "gun": int(r["gun"]),
                "etiket": r["tarih_etiketi"],
                "uretim_onceki": r[f"uretim_mwh_{ly}"],
                "uretim_simdiki": r[f"uretim_mwh_{this_year}"],
                "uretim_degisim": r["uretim_degisim"],
                "ciro_onceki": r[f"ciro_tl_{ly}"],
                "ciro_simdiki": r[f"ciro_tl_{this_year}"],
                "ciro_degisim": r["ciro_degisim"],
            }
        )

    def ozet_row(i):
        r = summary.iloc[i]
        return {"onceki": r[str(ly)], "simdiki": r[str(this_year)], "oran": r.iloc[3]}

    ortalamalar = []
    for i, r in avg.iterrows():
        ortalamalar.append(
            {
                "donem": r["donem"],
                "gun": int(r["gun"]),
                "ort_uretim_mwh": r["ort_uretim_mwh"],
                "ort_ciro_tl": r["ort_ciro_tl"],
                "oran_uretim": r["oran_uretim"],
                "oran_ciro": r["oran_ciro"],
                "referans": bool(i == len(avg) - 1),
            }
        )

    return sanitize(
        {
            "olusturma": datetime.now(TZ).isoformat(timespec="seconds"),
            "yil": this_year,
            "onceki_yil": ly,
            "baslangic": f"{this_year}-{core.START_MONTH:02d}-{core.START_DAY:02d}",
            "bitis": end_day.isoformat(),
            "veri_baslangici": start_day.isoformat(),
            "kur_onaylandi": bool(cfg.get("kur_onaylandi", False)),
            "gunluk": gunluk,
            "ozet": {"uretim": ozet_row(0), "ciro": ozet_row(1)},
            "ortalamalar": ortalamalar,
            "santraller": plant_tbl,
            "uyarilar": failures[:25],
            "uyari_sayisi": len(failures),
            "varsayimlar": {
                "yekdem": cfg["yekdem"],
                "kur_yillik": cfg["kur_yillik"],
                "kur_aylik": cfg.get("kur_aylik", {}),
            },
        }
    )


# ----------------------------------------------------------------------
# Ana akış
# ----------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true", help="tüm veriyi baştan çek")
    ap.add_argument("--end", help="bitiş günü (YYYY-AA-GG); verilmezse dün")
    args = ap.parse_args()

    cfg = load_config()
    if not cfg.get("kur_onaylandi", False):
        print("UYARI: config.json'da kur_onaylandi=false. Sitede 'örnek kur' uyarısı görünecek.")

    end_day = date.fromisoformat(args.end) if args.end else (datetime.now(TZ).date() - timedelta(days=1))
    this_year = end_day.year
    if end_day < date(this_year, core.START_MONTH, core.START_DAY):
        sys.exit("HATA: Bitiş günü 1 Ekim'den önce; bu site Ekim karşılaştırması için tasarlandı.")
    start_day = date(this_year - 1, core.START_MONTH, core.START_DAY)
    print(f"Veri dönemi: {start_day} -> {end_day}")

    eptr = core.connect()
    print("\nAydem santralleri aranıyor...")
    plants = core.find_aydem_plants(eptr)

    raw = load_raw()
    raw = raw[(raw["tarih"] >= pd.Timestamp(start_day)) & (raw["tarih"] <= pd.Timestamp(end_day))]

    if args.full or raw.empty:
        fetch_from = start_day
        print("\nTüm veri baştan çekilecek.")
    else:
        bad = incomplete_days(raw, start_day, end_day, len(plants))
        recent = end_day - timedelta(days=OVERLAP_DAYS - 1)
        prev_gap = load_gap()
        candidates = [recent] + [d.date() for d in bad] + ([prev_gap] if prev_gap else [])
        fetch_from = max(min(candidates), start_day)
        print(f"\nArtımlı güncelleme: {fetch_from} -> {end_day} ({len(bad)} eksik gün bulundu)")

    new, failures, gap_start, aborted = fetch_raw(eptr, plants, fetch_from, end_day)
    if new.empty:
        sys.exit("HATA: EPİAŞ'tan hiç veri alınamadı. Dosyalara dokunulmadı.")

    keep = raw[raw["tarih"] < pd.Timestamp(fetch_from)]
    parts = [df for df in (keep, new) if not df.empty]
    raw = pd.concat(parts, ignore_index=True)
    raw["tarih"] = pd.to_datetime(raw["tarih"])
    raw = raw.sort_values(["tarih", "santral"])
    RAW_CSV.parent.mkdir(parents=True, exist_ok=True)
    out = raw.copy()
    out["tarih"] = out["tarih"].dt.strftime("%Y-%m-%d")
    out.to_csv(RAW_CSV, index=False, encoding="utf-8")
    print(f"\nHam veri kaydedildi: {RAW_CSV} ({len(raw)} satır)")
    save_gap(gap_start)
    if aborted:
        sys.exit(
            "EPİAŞ'a erişim sorunu nedeniyle durdum. Çekilen veri kaydedildi; "
            "programı tekrar çalıştırınca kaldığı yerden devam eder."
        )

    detail = build_detail(raw)
    window = core.comparison_window(detail, end_day)
    if window[window["yil"] == this_year].empty:
        sys.exit("HATA: Bu yılın karşılaştırma döneminde veri yok.")

    comp = core.build_comparison(window, this_year)
    summary = core.build_summary(comp, this_year)
    avg = core.build_averages(detail, end_day)
    plant_tbl = build_plant_table(window, this_year)

    DOCS.mkdir(parents=True, exist_ok=True)
    core.OUTPUT_XLSX = str(XLSX_PATH)
    core.write_excel(comp, detail, summary, avg, this_year)

    payload = build_json(cfg, comp, summary, avg, plant_tbl, this_year, end_day, start_day, failures)
    DATA_JSON.write_text(
        json.dumps(payload, ensure_ascii=False, allow_nan=False, indent=1), encoding="utf-8"
    )
    print(f"Site verisi yazıldı: {DATA_JSON}")

    if failures:
        print(f"\nUYARI: {len(failures)} parça çekilemedi/boş geldi (sitede uyarı olarak görünür):")
        for f in failures[:25]:
            print("  -", f)


if __name__ == "__main__":
    main()
