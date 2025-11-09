import os
import glob
import pandas as pd
import numpy as np
from tkinter import Tk, Frame, Button, Label, StringVar, LEFT, RIGHT, TOP, BOTTOM, X
from tkinter.messagebox import showinfo
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.widgets import RectangleSelector

xlsx_in = "Excel.xlsx"
xlsx_out = "Excel_labeled.xlsx"
download_dir = os.path.join(os.getcwd(), "download")

df = pd.read_excel(xlsx_in, engine="openpyxl", header=None)
n = df.shape[1]
for i in range(n + 1, 6):
    df[f"col_{i}"] = np.nan
cols = [f"col_{i+1}" for i in range(max(n,5))]
df.columns = cols
col_A, col_D, col_E = "col_1", "col_4", "col_5"
label_col = "col_6"
if label_col not in df.columns:
    df[label_col] = np.nan

def not_empty(x):
    if pd.isna(x):
        return False
    if isinstance(x, str):
        return x.strip() != ""
    try:
        return str(x).strip() != ""
    except:
        return False

mask = (df[col_D].astype(str).str.strip() == "Ні") & df[col_E].apply(not_empty)
candidates = df[mask].copy()
indices = candidates.index.tolist()
tic_values = candidates[col_A].astype(str).str.extract(r"(\d+)")[0].tolist()

class App:
    def __init__(self, master):
        self.master = master
        master.title("TESS Light Curve Classifier")
        self.idx_pos = 0
        self.current_tic = None
        self.current_index = None
        self.data = None
        self.x_all = None
        self.status_text = StringVar()
        top_frame = Frame(master)
        top_frame.pack(side=TOP, fill=X)
        self.lbl = Label(top_frame, textvariable=self.status_text, anchor="w")
        self.lbl.pack(side=LEFT, fill=X, expand=True)
        Button(top_frame, text="⬅ Попередня", command=self.prev_star, width=14).pack(side=RIGHT)
        Button(top_frame, text="Наступна ➡", command=self.next_star, width=14).pack(side=RIGHT)
        self.fig = Figure(figsize=(8,5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=master)
        self.canvas.get_tk_widget().pack(side=TOP, fill="both", expand=True)
        bottom = Frame(master)
        bottom.pack(side=BOTTOM, fill=X)
        Button(bottom, text="Скинути зум", command=self.reset_zoom, width=14).pack(side=RIGHT)
        Button(bottom, text="Квазі-періодична", command=lambda: self.classify_and_save("Квазі-періодична"), width=18).pack(side=LEFT)
        Button(bottom, text="Неперіодична", command=lambda: self.classify_and_save("Неперіодична"), width=18).pack(side=LEFT)
        Button(bottom, text="Періодична", command=lambda: self.classify_and_save("Періодична"), width=18).pack(side=LEFT)
        rectprops = dict(edgecolor="black", facecolor=(1,1,1,0.35), linestyle=":", linewidth=1)
        try:
            self.selector = RectangleSelector(self.ax, self.on_select, useblit=True, button=[1], minspanx=0, minspany=0, spancoords='data', interactive=False, props=rectprops)
        except TypeError:
            self.selector = RectangleSelector(self.ax, self.on_select, useblit=True, button=[1], minspanx=0, minspany=0, spancoords='data', interactive=False, rectprops=rectprops)
        master.bind("1", lambda e: self.classify_and_save("Квазі-періодична"))
        master.bind("2", lambda e: self.classify_and_save("Неперіодична"))
        master.bind("3", lambda e: self.classify_and_save("Періодична"))
        master.bind("<Left>", lambda e: self.prev_star())
        master.bind("<Right>", lambda e: self.next_star())
        master.bind("<Escape>", lambda e: self.reset_zoom())
        self.load_current()

    def tess_path_for_tic(self, tic_str):
        bases = [download_dir, os.getcwd()]
        pats = [
            f"TIC_{tic_str}__mag.tess",
            f"TIC_{tic_str}_mag.tess",
            f"TIC_{tic_str}*__mag.tess",
            f"TIC_{tic_str}*_mag.tess",
            f"**/TIC_{tic_str}__mag.tess",
            f"**/TIC_{tic_str}_mag.tess",
            f"**/TIC_{tic_str}*mag.tess",
            f"**/*TIC_{tic_str}*mag.tess",
        ]
        files = []
        for b in bases:
            for p in pats:
                files.extend(glob.glob(os.path.join(b, p), recursive=True))
        files = sorted(set(files), key=lambda x: (0 if x.startswith(download_dir) else 1, len(x), x))
        return files[0] if files else None

    def read_tess(self, path):
        try:
            dat = pd.read_csv(path, delim_whitespace=True, comment="#", header=None, engine="python")
            numeric = dat.apply(pd.to_numeric, errors="coerce")
            x = numeric.iloc[:,0].to_numpy()
            y = numeric.iloc[:,1].to_numpy()
            m = np.isfinite(x) & np.isfinite(y)
            return x[m], y[m]
        except:
            dat = pd.read_csv(path, sep=r"[,\s;]+", comment="#", header=None, engine="python")
            numeric = dat.apply(pd.to_numeric, errors="coerce")
            x = numeric.iloc[:,0].to_numpy()
            y = numeric.iloc[:,1].to_numpy()
            m = np.isfinite(x) & np.isfinite(y)
            return x[m], y[m]

    def load_current(self):
        if not indices:
            self.status_text.set("Немає рядків для класифікації")
            self.canvas.draw()
            return
        self.current_index = indices[self.idx_pos]
        self.current_tic = tic_values[self.idx_pos]
        path = self.tess_path_for_tic(self.current_tic)
        self.ax.clear()
        if path is None or not os.path.exists(path):
            self.data = None
            self.x_all = None
            self.status_text.set(f"TIC {self.current_tic} | Файл не знайдено | {self.idx_pos+1}/{len(indices)} | Рядок {self.current_index+1}")
            self.canvas.draw()
            return
        x, y = self.read_tess(path)
        if x.size == 0 or y.size == 0:
            self.data = None
            self.x_all = None
            self.status_text.set(f"TIC {self.current_tic} | Порожні дані | {self.idx_pos+1}/{len(indices)} | Рядок {self.current_index+1}")
            self.canvas.draw()
            return
        self.data = (x, y, path)
        self.x_all = (np.nanmin(x), np.nanmax(x))
        self.ax.scatter(x, y, s=14, c="k", marker=".", linewidths=0)  # ← зробив точки трохи більшими
        try:
            self.ax.invert_yaxis()
        except:
            pass
        self.ax.set_xlabel("Час")
        self.ax.set_ylabel("Яскравість (mag)")
        self.ax.set_title(f"Light curve for TIC {self.current_tic}")
        self.status_text.set(f"TIC {self.current_tic} | {os.path.basename(path)} | {self.idx_pos+1}/{len(indices)} | Рядок {self.current_index+1}")
        self.canvas.draw()

    def on_select(self, eclick, erelease):
        if self.data is None:
            return
        x1, x2 = eclick.xdata, erelease.xdata
        if x1 is None or x2 is None:
            return
        a, b = sorted([x1, x2])
        self.ax.set_xlim(a, b)
        self.canvas.draw()

    def reset_zoom(self):
        if self.data is None or self.x_all is None:
            return
        self.ax.set_xlim(self.x_all[0], self.x_all[1])
        self.canvas.draw()

    def classify_and_save(self, label):
        df.at[self.current_index, label_col] = label
        try:
            df.to_excel(xlsx_out, index=False, engine="openpyxl", header=False)
        except:
            tmp = xlsx_out.replace(".xlsx", "_tmp.xlsx")
            df.to_excel(tmp, index=False, engine="openpyxl", header=False)
        if self.idx_pos < len(indices) - 1:
            self.idx_pos += 1
            self.load_current()
        else:
            showinfo("Готово", "Це була остання зоря.")

    def next_star(self):
        if self.idx_pos < len(indices) - 1:
            self.idx_pos += 1
            self.load_current()
        else:
            showinfo("Готово", "Це була остання зоря.")

    def prev_star(self):
        if self.idx_pos > 0:
            self.idx_pos -= 1
            self.load_current()

root = Tk()
app = App(root)
root.mainloop()
