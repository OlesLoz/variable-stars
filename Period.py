import os, sys
from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
try:
    import lightkurve as lk
except ImportError:
    print("pip install lightkurve"); sys.exit(1)
try:
    import pyperclip
    _HAS_CLIPBOARD = True
except Exception:
    _HAS_CLIPBOARD = False

EXCEL_PATH = "Excel_labeled.xlsx"
DOWNLOAD_DIR = "download"
SECTOR = 95
LABEL_NEEDLE = "Періодична"
FREQUENCY_XLIM = (0.0, 7.0)
OVERSAMPLE = 10
MANUAL_TOLERANCE = 0.05
RESULTS_XLSX = "results_periods.xlsx"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def _read_tic_list_from_excel(xlsx_path: str) -> list:
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    if 'F' in df.columns:
        label_series = df['F']
    else:
        if len(df.columns) < 6:
            raise ValueError("У файлі менше 6 колонок.")
        label_series = df.iloc[:, 5]
    mask_periodic = label_series.astype(str).str.strip().str.lower() == LABEL_NEEDLE.lower()
    dfp = df[mask_periodic].copy()
    if 'TIC' in dfp.columns:
        tic_series = dfp['TIC']
    else:
        tic_series = dfp.iloc[:, 0]
    tics = []
    for val in tic_series.astype(str):
        s = val.strip()
        if not s:
            continue
        s = s.upper()
        if not s.startswith("TIC"):
            s = "TIC " + s
        tics.append(s)
    return tics

def _load_lightcurve_spoc_sector(tic_id: str, sector: int):
    sr = lk.search_lightcurve(tic_id, author='SPOC', sector=sector)
    if len(sr) == 0:
        return None
    try:
        lc = sr.download()
        if hasattr(lc, "stitch"):
            lc = lc.stitch().remove_nans().normalize()
        else:
            lc = lc.remove_nans().normalize()
        return lc
    except Exception as e:
        print(f"[{tic_id}] Помилка LC сектор {sector}: {e}")
        return None

def _compute_periodogram(lc):
    pg = lc.to_periodogram(oversample_factor=OVERSAMPLE)
    auto_period = pg.period_at_max_power.value
    auto_freq = 1.0 / auto_period
    return pg, auto_period, auto_freq

def _fold_lc_manual(lc, period_days: float):
    t = lc.time.value if hasattr(lc.time, 'value') else np.array(lc.time)
    fl = lc.flux.value if hasattr(lc.flux, 'value') else np.array(lc.flux)
    t0 = t[0]
    phase = ((t - t0) / period_days) % 1.0
    phase[phase > 0.5] -= 1.0
    return phase, fl

def _export_phase_png(phase, flux, tic_id: str, sector: int, p_days: float, outpath: str):
    tic_num = tic_id.replace("TIC ", "").strip()
    fig = plt.figure(figsize=(15, 8))
    ax = fig.add_subplot(111)
    ax.plot(phase, flux, ".k", markersize=2)
    ax.set_xlim(-0.5, 0.5)
    ax.set_xlabel("Phase")
    ax.set_ylabel("Magnitude, mmag")
    ax.set_title(f"Phase curve for TIC {tic_num} ( P = {p_days:.5f} днів )")
    fig.savefig(outpath, dpi=200, bbox_inches='tight')
    plt.close(fig)

def _append_result_xlsx(tic_id: str, sector: int, chosen_p: float, method: str):
    row = {
        "TIC": tic_id,
        "Sector": sector,
        "P_days": round(float(chosen_p), 5),
        "Method": method,
        "Timestamp": datetime.utcnow().isoformat() + "Z",
    }
    if os.path.exists(RESULTS_XLSX):
        try:
            df = pd.read_excel(RESULTS_XLSX, engine="openpyxl")
        except Exception:
            df = pd.DataFrame(columns=row.keys())
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    else:
        df = pd.DataFrame([row])
    with pd.ExcelWriter(RESULTS_XLSX, engine="openpyxl", mode="w") as w:
        df.to_excel(w, index=False, sheet_name="results")

def interactive_period_picker(lc, tic_id: str, sector: int, auto_period: float, auto_frequency: float, periodogram):
    chosen_period = float(auto_period)
    chosen_method = "auto"
    phase, flux = _fold_lc_manual(lc, chosen_period)
    fig = plt.figure(figsize=(15, 7), constrained_layout=True)
    gs = fig.add_gridspec(1, 2, width_ratios=[1.1, 1.0])
    ax_pg = fig.add_subplot(gs[0, 0])
    ax_ph = fig.add_subplot(gs[0, 1])
    freqs = periodogram.frequency.value
    power = periodogram.power
    ax_pg.plot(freqs, power, lw=1)
    auto_vline = ax_pg.axvline(x=auto_frequency, linestyle='--', color='gray', label=f"Auto: {auto_period:.5f} d")
    manual_vline = ax_pg.axvline(x=auto_frequency, linestyle=':', color='red', label=f"Manual: {auto_period:.5f} d")
    ax_pg.set_title(f"Періодограма для {tic_id} (Sector {sector})")
    ax_pg.set_xlabel("Частота (1/д)")
    ax_pg.set_ylabel("Потужність")
    ax_pg.set_xlim(*FREQUENCY_XLIM)
    ax_pg.legend(loc="upper right")
    ph_scatter = ax_ph.plot(phase, flux, ".k", markersize=2)[0]
    ax_ph.set_title(f"Фазова крива для {tic_id}\nP = {chosen_period:.5f} днів")
    ax_ph.set_xlabel("Фаза")
    ax_ph.set_ylabel("Magnitude, mmag")
    ax_ph.set_xlim(-0.5, 0.5)
    btn_ax_auto = fig.add_axes([0.10, 0.02, 0.12, 0.06])
    btn_ax_save = fig.add_axes([0.27, 0.02, 0.12, 0.06])
    btn_ax_next = fig.add_axes([0.44, 0.02, 0.12, 0.06])
    btn_auto = Button(btn_ax_auto, "Use Auto")
    btn_save = Button(btn_ax_save, "Save PNG")
    btn_next = Button(btn_ax_next, "Next ▶")

    def _update_phase(period_days: float, method: str):
        nonlocal chosen_period, chosen_method, phase, flux
        chosen_period = float(period_days)
        chosen_method = method
        manual_vline.set_xdata([1.0/period_days, 1.0/period_days])
        manual_vline.set_label(f"Manual: {period_days:.5f} d")
        ax_pg.legend(loc="upper right")
        phase, flux = _fold_lc_manual(lc, chosen_period)
        ph_scatter.set_xdata(phase)
        ph_scatter.set_ydata(flux)
        ax_ph.set_title(f"Фазова крива для {tic_id}\nP = {chosen_period:.5f} днів")
        ax_ph.set_xlim(-0.5, 0.5)
        fig.canvas.draw_idle()
        print(f"[{tic_id}] Обраний період: {chosen_period:.5f} днів ({chosen_method})")

    def on_click(event):
        if event.inaxes != ax_pg or event.xdata is None: return
        fx = event.xdata
        idx_win = np.where(np.abs(freqs - fx) < MANUAL_TOLERANCE)[0]
        if len(idx_win) > 0:
            local_idx = idx_win[np.argmax(power[idx_win])]
        else:
            local_idx = int(np.argmin(np.abs(freqs - fx)))
        f_sel = float(freqs[local_idx])
        p_sel = 1.0 / f_sel
        _update_phase(p_sel, "manual")

    def on_key(event):
        if event.key is None: return
        key = event.key.lower()
        if key == 'a': _update_phase(auto_period, "auto")
        elif key == 's': do_save_png()
        elif key in ('enter', 'n'): plt.close(fig)
        elif key == 'c' and _HAS_CLIPBOARD:
            try:
                pyperclip.copy(f"{chosen_period:.5f}")
                print(f"[{tic_id}] Період {chosen_period:.5f} скопійовано.")
            except Exception: pass

    def do_auto(event=None): _update_phase(auto_period, "auto")

    def do_save_png(event=None):
        out_png = os.path.join(DOWNLOAD_DIR, f"{tic_id.replace(' ', '_')}_sector{sector}_phase.png")
        _export_phase_png(phase, flux, tic_id, sector, chosen_period, out_png)
        print(f"[{tic_id}] Збережено PNG: {out_png}")

    def do_next(event=None): plt.close(fig)

    fig.canvas.mpl_connect('button_press_event', on_click)
    fig.canvas.mpl_connect('key_press_event', on_key)
    btn_auto.on_clicked(do_auto)
    btn_save.on_clicked(do_save_png)
    btn_next.on_clicked(do_next)
    plt.show()
    out_png = os.path.join(DOWNLOAD_DIR, f"{tic_id.replace(' ', '_')}_sector{sector}_phase.png")
    if not os.path.exists(out_png):
        _export_phase_png(phase, flux, tic_id, sector, chosen_period, out_png)
        print(f"[{tic_id}] Автозбережено PNG: {out_png}")
    _append_result_xlsx(tic_id, sector, chosen_period, chosen_method)
    return chosen_period, chosen_method

def main():
    try:
        tic_list = _read_tic_list_from_excel(EXCEL_PATH)
    except Exception as e:
        print(f"Помилка читання '{EXCEL_PATH}': {e}")
        sys.exit(1)
    if not tic_list:
        print("Немає зір із міткою 'Періодична' у колонці F."); sys.exit(0)
    print(f"Знайдено {len(tic_list)} зір. Sector {SECTOR}.")
    for i, tic_id in enumerate(tic_list, start=1):
        print(f"[{i}/{len(tic_list)}] {tic_id}...")
        lc = _load_lightcurve_spoc_sector(tic_id, SECTOR)
        if lc is None:
            print(f"[{tic_id}] Немає даних SPOC у секторі {SECTOR}. Пропуск.")
            continue
        try:
            periodogram, auto_p, auto_f = _compute_periodogram(lc)
        except Exception as e:
            print(f"[{tic_id}] Не вдалося порахувати періодограму: {e}")
            continue
        print(f"[{tic_id}] Auto період: {auto_p:.5f} дн (f={auto_f:.5f} 1/д)")
        _ = interactive_period_picker(lc, tic_id, SECTOR, auto_p, auto_f, periodogram)
    print("Готово. Дивись 'results_periods.xlsx' і PNG у 'download/'.")

if __name__ == "__main__":
    main()
