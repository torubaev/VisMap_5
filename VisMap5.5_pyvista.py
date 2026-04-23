#!/usr/bin/env python3
"""
VisMap GUI + safer launcher for VisMap4.2 workflow

What this version does:
- starts with a GUI
- lets you browse for the input .wfn / .wfx / .fchk file
- lets you set nproc, mode, visualization, and optional CP isovalue
- fixes Windows path handling for Multiwfn.exe
- safely overwrites existing cube/output files instead of crashing
- runs Multiwfn from the folder containing the selected wavefunction file
- keeps original VisMap processing logic as much as possible

Run:
    py VisMapGUI_fixed.py

Optional CLI mode is also supported:
    py VisMapGUI_fixed.py yourfile.wfn -nproc=8 -mode=old -vis=y
"""

import os
import sys
import shutil
import subprocess
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser

# ----------------------------
# Config
# ----------------------------

DEFAULT_MULTIWFN_PATHS = [
    r"C:/Multiwfn_2026.2.2_bin_Win64/Multiwfn.exe",
    r"C:\Multiwfn_3.8_dev_bin_Win64\Multiwfn.exe",
    r"C:\Multiwfn\Multiwfn.exe",
]

Multiwfnpath = DEFAULT_MULTIWFN_PATHS[0]

# dictionary of (nuclear_charge : Nucleus, vdW radii in Angstroms)
dnc2all = {1: ['H', 1.09, 0.23, 1.00794, 0.99609375, 0.9765625, 0.80078125],
           2: ['He', 1.40, 1.50, 4.002602, 0.72265625, 0.82421875, 0.9296875],
           3: ['Li', 1.82, 1.28, 6.941, 0.7421875, 0.7421875, 0.7421875],
           4: ['Be', 2.00, 0.96, 9.012182, 0.7421875, 0.7421875, 0.7421875],
           5: ['B', 2.00, 0.83, 10.811, 0.625, 0.234375, 0.234375],
           6: ['C', 1.70, 0.68, 12.0107, 0.328125, 0.328125, 0.328125],
           7: ['N', 1.55, 0.68, 14.0067, 0.1171875, 0.5625, 0.99609375],
           8: ['O', 1.52, 0.68, 15.9994, 0.99609375, 0.0, 0.0],
           9: ['F', 1.47, 0.64, 18.998403, 0.99609375, 0.99609375, 0.0],
           10: ['Ne', 1.54, 1.50, 20.1797, 0.72265625, 0.82421875, 0.9296875],
           11: ['Na', 2.27, 1.66, 22.98977, 0.7421875, 0.7421875, 0.7421875],
           12: ['Mg', 1.73, 1.41, 24.305, 0.7421875, 0.7421875, 0.7421875],
           13: ['Al', 2.00, 1.21, 26.981538, 0.7421875, 0.7421875, 0.7421875],
           14: ['Si', 2.10, 1.20, 28.0855, 0.82421875, 0.82421875, 0.82421875],
           15: ['P', 1.80, 1.05, 30.973761, 0.99609375, 0.546875, 0.0],
           16: ['S', 1.80, 1.02, 32.065, 0.99609375, 0.9609375, 0.55859375],
           17: ['Cl', 1.75, 0.99, 35.453, 0.0, 0.99609375, 0.0],
           18: ['Ar', 1.88, 1.51, 39.948, 0.72265625, 0.82421875, 0.9296875],
           19: ['K', 2.75, 2.03, 39.0983, 0.7421875, 0.7421875, 0.7421875],
           20: ['Ca', 2.00, 1.76, 40.078, 0.7421875, 0.7421875, 0.7421875],
           21: ['Sc', 2.00, 1.70, 44.95591, 0.7421875, 0.7421875, 0.7421875],
           22: ['Ti', 2.00, 1.60, 47.867, 0.7421875, 0.7421875, 0.7421875],
           23: ['V', 2.00, 1.53, 50.9415, 0.7421875, 0.7421875, 0.7421875],
           24: ['Cr', 2.00, 1.39, 51.9961, 0.7421875, 0.7421875, 0.7421875],
           25: ['Mn', 2.50, 1.61, 54.938049, 0.7421875, 0.7421875, 0.7421875],
           26: ['Fe', 2.00, 1.52, 55.845, 0.7421875, 0.7421875, 0.7421875],
           27: ['Co', 2.00, 1.26, 58.9332, 0.7421875, 0.7421875, 0.7421875],
           28: ['Ni', 1.63, 1.24, 58.6934, 0.7421875, 0.7421875, 0.7421875],
           29: ['Cu', 1.40, 1.32, 63.546, 0.99609375, 0.5078125, 0.27734375],
           30: ['Zn', 1.39, 1.22, 65.409, 0.7421875, 0.7421875, 0.7421875],
           31: ['Ga', 1.87, 1.22, 69.723, 0.7421875, 0.7421875, 0.7421875],
           32: ['Ge', 2.00, 1.17, 72.64, 0.7421875, 0.7421875, 0.7421875],
           33: ['As', 1.85, 1.21, 74.9216, 0.7421875, 0.7421875, 0.7421875],
           34: ['Se', 1.90, 1.22, 78.96, 0.7421875, 0.7421875, 0.7421875],
           35: ['Br', 1.85, 1.21, 79.904, 0.7421875, 0.5078125, 0.234375],
           36: ['Kr', 2.02, 1.50, 83.798, 0.72265625, 0.82421875, 0.9296875],
           37: ['Rb', 2.00, 2.20, 85.4678, 0.7421875, 0.7421875, 0.7421875],
           38: ['Sr', 2.00, 1.95, 87.62, 0.7421875, 0.7421875, 0.7421875],
           39: ['Y', 2.00, 1.90, 88.90585, 0.7421875, 0.7421875, 0.7421875],
           40: ['Zr', 2.00, 1.75, 91.224, 0.7421875, 0.7421875, 0.7421875],
           41: ['Nb', 2.00, 1.64, 92.90638, 0.7421875, 0.7421875, 0.7421875],
           42: ['Mo', 2.00, 1.54, 95.94, 0.7421875, 0.7421875, 0.7421875],
           43: ['Tc', 2.00, 1.47, 98.0, 0.7421875, 0.7421875, 0.7421875],
           44: ['Ru', 2.00, 1.46, 101.07, 0.7421875, 0.7421875, 0.7421875],
           45: ['Rh', 2.00, 1.45, 102.9055, 0.7421875, 0.7421875, 0.7421875],
           46: ['Pd', 1.63, 1.39, 106.42, 0.7421875, 0.7421875, 0.7421875],
           47: ['Ag', 1.72, 1.45, 107.8682, 0.99609375, 0.99609375, 0.99609375],
           48: ['Cd', 1.58, 1.44, 112.411, 0.7421875, 0.7421875, 0.7421875],
           49: ['In', 1.93, 1.42, 114.818, 0.7421875, 0.7421875, 0.7421875],
           50: ['Sn', 2.17, 1.39, 118.71, 0.7421875, 0.7421875, 0.7421875],
           51: ['Sb', 2.00, 1.39, 121.76, 0.7421875, 0.7421875, 0.7421875],
           52: ['Te', 2.06, 1.47, 127.6, 0.7421875, 0.7421875, 0.7421875],
           53: ['I', 2.58, 1.40, 126.90447, 0.625, 0.125, 0.9375],
           54: ['Xe', 2.16, 1.50, 131.293, 0.72265625, 0.82421875, 0.9296875],
           55: ['Cs', 2.00, 2.44, 132.90545, 0.7421875, 0.7421875, 0.7421875],
           56: ['Ba', 2.00, 2.15, 137.327, 0.7421875, 0.7421875, 0.7421875],
           57: ['La', 2.00, 2.07, 138.9055, 0.7421875, 0.7421875, 0.7421875],
           58: ['Ce', 2.00, 2.04, 140.116, 0.7421875, 0.7421875, 0.7421875],
           59: ['Pr', 2.00, 2.03, 140.90765, 0.7421875, 0.7421875, 0.7421875],
           60: ['Nd', 2.00, 2.01, 144.24, 0.7421875, 0.7421875, 0.7421875],
           61: ['Pm', 2.00, 1.99, 145.0, 0.7421875, 0.7421875, 0.7421875],
           62: ['Sm', 2.00, 1.98, 150.36, 0.7421875, 0.7421875, 0.7421875],
           63: ['Eu', 2.00, 1.98, 151.964, 0.7421875, 0.7421875, 0.7421875],
           64: ['Gd', 2.00, 1.96, 157.25, 0.7421875, 0.7421875, 0.7421875],
           65: ['Tb', 2.00, 1.94, 158.92534, 0.7421875, 0.7421875, 0.7421875],
           66: ['Dy', 2.00, 1.92, 162.5, 0.7421875, 0.7421875, 0.7421875],
           67: ['Ho', 2.00, 1.92, 164.93032, 0.7421875, 0.7421875, 0.7421875],
           68: ['Er', 2.00, 1.89, 167.259, 0.7421875, 0.7421875, 0.7421875],
           69: ['Tm', 2.00, 1.90, 168.93421, 0.7421875, 0.7421875, 0.7421875],
           70: ['Yb', 2.00, 1.87, 173.04, 0.7421875, 0.7421875, 0.7421875],
           71: ['Lu', 2.00, 1.87, 174.967, 0.7421875, 0.7421875, 0.7421875],
           72: ['Hf', 2.00, 1.75, 178.49, 0.7421875, 0.7421875, 0.7421875],
           73: ['Ta', 2.00, 1.70, 180.9479, 0.7421875, 0.7421875, 0.7421875],
           74: ['W', 2.00, 1.62, 183.84, 0.7421875, 0.7421875, 0.7421875],
           75: ['Re', 2.00, 1.51, 186.207, 0.7421875, 0.7421875, 0.7421875],
           76: ['Os', 2.00, 1.44, 190.23, 0.7421875, 0.7421875, 0.7421875],
           77: ['Ir', 2.00, 1.41, 192.217, 0.7421875, 0.7421875, 0.7421875],
           78: ['Pt', 1.72, 1.36, 195.078, 0.7421875, 0.7421875, 0.7421875],
           79: ['Au', 1.66, 1.50, 196.96655, 0.99609375, 0.83984375, 0.0],
           80: ['Hg', 1.55, 1.32, 200.59, 0.7421875, 0.7421875, 0.7421875],
           81: ['Tl', 1.96, 1.45, 204.3833, 0.7421875, 0.7421875, 0.7421875],
           82: ['Pb', 2.02, 1.46, 207.2, 0.7421875, 0.7421875, 0.7421875],
           83: ['Bi', 2.00, 1.48, 208.98038, 0.7421875, 0.7421875, 0.7421875],
           84: ['Po', 2.00, 1.40, 290.0, 0.7421875, 0.7421875, 0.7421875],
           85: ['At', 2.00, 1.21, 210.0, 0.7421875, 0.7421875, 0.7421875],
           86: ['Rn', 2.00, 1.50, 222.0, 0.72265625, 0.82421875, 0.9296875],
           87: ['Fr', 2.00, 2.60, 223.0, 0.7421875, 0.7421875, 0.7421875],
           88: ['Ra', 2.00, 2.21, 226.0, 0.7421875, 0.7421875, 0.7421875],
           89: ['Ac', 2.00, 2.15, 227.0, 0.7421875, 0.7421875, 0.7421875],
           90: ['Th', 2.00, 2.06, 232.0381, 0.7421875, 0.7421875, 0.7421875],
           91: ['Pa', 2.00, 2.00, 231.03588, 0.7421875, 0.7421875, 0.7421875],
           92: ['U', 1.86, 1.96, 238.02891, 0.7421875, 0.7421875, 0.7421875],
           93: ['Np', 2.00, 1.90, 237.0, 0.7421875, 0.7421875, 0.7421875],
           94: ['Pu', 2.00, 1.87, 244.0, 0.7421875, 0.7421875, 0.7421875],
           95: ['Am', 2.00, 1.80, 243.0, 0.7421875, 0.7421875, 0.7421875],
           96: ['Cm', 2.00, 1.69, 247.0, 0.7421875, 0.7421875, 0.7421875],
           97: ['Bk', 2.00, 1.54, 247.0, 0.7421875, 0.7421875, 0.7421875],
           98: ['Cf', 2.00, 1.83, 251.0, 0.7421875, 0.7421875, 0.7421875],
           99: ['Es', 2.00, 1.50, 252.0, 0.7421875, 0.7421875, 0.7421875],
           100: ['Fm', 2.00, 1.50, 257.0, 0.7421875, 0.7421875, 0.7421875],
           101: ['Md', 2.00, 1.50, 258.0, 0.7421875, 0.7421875, 0.7421875],
           102: ['No', 2.00, 1.50, 259.0, 0.7421875, 0.7421875, 0.7421875],
           103: ['Lr', 2.00, 1.50, 262.0, 0.7421875, 0.7421875, 0.7421875]
           }


# ----------------------------
# Helpers
# ----------------------------

def find_multiwfn_path():
    env_path = os.environ.get("Multiwfnpath")
    if env_path and os.path.exists(env_path):
        return env_path

    for p in DEFAULT_MULTIWFN_PATHS:
        if os.path.exists(p):
            return p

    return Multiwfnpath


def safe_remove(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def safe_move(src, dst):
    if not os.path.exists(src):
        raise FileNotFoundError(f"Expected output file was not created: {src}")
    if os.path.exists(dst):
        os.remove(dst)
    shutil.move(src, dst)


def base_name_no_ext(path):
    return os.path.splitext(os.path.abspath(path))[0]


def run_command_capture(cmd, cwd=None):
    proc = subprocess.run(
        cmd,
        cwd=cwd,
        text=True,
        capture_output=True,
        shell=False
    )
    return proc.returncode, proc.stdout, proc.stderr


# ----------------------------
# Global runtime variables
# ----------------------------

inputfile = ""
fname = ""
nproc = "4"
mode = "old"
vis = "y"
PreGenCP = False
CPisov = None
workdir = ""
mwfn_exe = ""

APP_STATE = {
    "root": None,
    "status_label": None,
    "extrema_text": None,
    "viewer_controls": [],
    "extrema_scrollbar": None,
}

VIEWER_STATE = None


def _get_screen_size():
    root = APP_STATE.get("root")
    try:
        if root is not None and root.winfo_exists():
            return int(root.winfo_screenwidth()), int(root.winfo_screenheight())
    except Exception:
        pass
    try:
        probe = tk.Tk()
        probe.withdraw()
        width = int(probe.winfo_screenwidth())
        height = int(probe.winfo_screenheight())
        probe.destroy()
        return width, height
    except Exception:
        return 1920, 1080


def _main_window_size():
    screen_w, screen_h = _get_screen_size()
    return int(screen_w * 0.45), int(screen_h * 0.88)


def _viewer_window_size():
    screen_w, screen_h = _get_screen_size()
    return int(screen_w * 0.45), int(screen_h * 0.88)


# ----------------------------
# Multiwfn execution
# ----------------------------

def Run_MWFN(mytext, needout=False):
    inp_path = os.path.join(workdir, "myprog.inp")
    out_path = os.path.join(workdir, "myprog.out")

    with open(inp_path, "w", newline="\n") as inp:
        inp.write("\n".join(mytext) + "\n")

    with open(inp_path, "r") as fin:
        if needout:
            with open(out_path, "w", newline="\n") as fout:
                proc = subprocess.run(
                    [mwfn_exe, inputfile],
                    stdin=fin,
                    stdout=fout,
                    stderr=subprocess.STDOUT,
                    cwd=workdir,
                    text=True,
                    shell=False
                )
        else:
            proc = subprocess.run(
                [mwfn_exe, inputfile],
                stdin=fin,
                cwd=workdir,
                text=True,
                shell=False
            )

    safe_remove(inp_path)

    # Multiwfn may exit non-zero after finishing useful work if scripted input ended.
    # We do not hard-fail here; downstream file existence is the real criterion.
    return proc.returncode


def ReadCUB(inputcube):
    CENTERS = []
    Scalars = []
    with open(inputcube, "r") as cube:
        lines = cube.readlines()[2:]
        nat = int(lines[0].split()[0])
        pointsv1 = int(lines[1].split()[0])
        pointsv2 = int(lines[2].split()[0])
        pointsv3 = int(lines[3].split()[0])

        origin = np.array([float(x) for x in lines[0].split()[-3:]])
        v1 = np.array([float(x) for x in lines[1].split()[1:]])
        v2 = np.array([float(x) for x in lines[2].split()[1:]])
        v3 = np.array([float(x) for x in lines[3].split()[1:]])

        for i in range(nat):
            line = lines[4 + i].split()
            CENTERS.append([int(line[0])] + [float(x) for x in line[2:]] + [dnc2all[int(line[0])][0] + str(i + 1)])
        for line in lines[4 + nat:]:
            Scalars += [float(x) for x in line.split()]

        print("Found", len(Scalars), "Scalars in", inputcube)

    XYZS_Data = [origin, v1, v2, v3, pointsv1, pointsv2, pointsv3, Scalars]
    return XYZS_Data, CENTERS


def CalcCub(IsCalced, fname_base):
    dens_target = fname_base + "_Dens.cub"
    esp_target = fname_base + "_ESP.cub"

    if IsCalced[0] is False:
        print("Calculating Density cube")
        safe_remove(os.path.join(workdir, "density.cub"))

        if IsCalced[1]:
            text = ["1000", "10", nproc, "5", "1", "8", esp_target, "2"]
            Run_MWFN(text, False)
        else:
            text = ["1000", "10", nproc, "5", "1", "2", "2"]
            Run_MWFN(text, False)

        safe_move(os.path.join(workdir, "density.cub"), dens_target)
        IsCalced[0] = True

    if IsCalced[1] is False:
        print("Calculating ESP cube")
        safe_remove(os.path.join(workdir, "totesp.cub"))

        if IsCalced[0]:
            text = ["1000", "10", nproc, "5", "12", "8", dens_target, "2"]
            Run_MWFN(text, False)
        else:
            text = ["1000", "10", nproc, "5", "12", "2", "2"]
            Run_MWFN(text, False)

        safe_move(os.path.join(workdir, "totesp.cub"), esp_target)
        IsCalced[1] = True


def CalcPoints(isoval):
    print("Searching surfanalysis.txt file with points")
    out_name = fname + "_sa_" + str(isoval) + ".txt"

    if not os.path.exists(out_name):
        print("File not found. Calling Multiwfn for min/max locating")
        safe_remove(os.path.join(workdir, "surfanalysis.txt"))
        text = ["1000", "10", nproc, "12", "1", "1", str(isoval), "0", "1"]
        Run_MWFN(text, False)
        safe_move(os.path.join(workdir, "surfanalysis.txt"), out_name)

    MAXMIN = []
    with open(out_name, "r") as out:
        for line in out:
            line = line.split()
            if len(line) > 5:
                if "*" not in line and "eV" not in line:
                    MAXMIN.append([float(line[3])] + [float(x) / 0.529 for x in line[4:]])
                elif "*" in line and "eV" not in line:
                    MAXMIN.append([float(line[4])] + [float(x) / 0.529 for x in line[5:]])

    print("Located", len(MAXMIN), "extremum points on the surface")
    return MAXMIN


def BuildBondPairs(CENTERS):
    bond_pairs = []
    for i, atom1 in enumerate(CENTERS):
        pos1 = np.array(atom1[1:4], dtype=float)
        for j in range(i + 1, len(CENTERS)):
            atom2 = CENTERS[j]
            pos2 = np.array(atom2[1:4], dtype=float)
            dist = np.linalg.norm(pos1 - pos2)
            if 1.0e-8 < dist < (dnc2all[atom1[0]][1] + dnc2all[atom2[0]][1]):
                bond_pairs.append((i, j))
    return bond_pairs


def BuildPyVistaGrid(CUBdat, CUBdatESP, xx, yy, zz):
    import pyvista as pv

    shape = CUBdat.shape
    if shape != CUBdatESP.shape:
        raise ValueError("Density and ESP grids have different shapes.")

    if xx.shape != shape or yy.shape != shape or zz.shape != shape:
        raise ValueError("Coordinate grids and scalar grids have inconsistent shapes.")

    image = pv.ImageData()
    image.dimensions = np.array(shape, dtype=int)
    image.origin = (float(xx[0, 0, 0]), float(yy[0, 0, 0]), float(zz[0, 0, 0]))

    spacing_x = float(np.linalg.norm([xx[1, 0, 0] - xx[0, 0, 0], yy[1, 0, 0] - yy[0, 0, 0], zz[1, 0, 0] - zz[0, 0, 0]])) if shape[0] > 1 else 1.0
    spacing_y = float(np.linalg.norm([xx[0, 1, 0] - xx[0, 0, 0], yy[0, 1, 0] - yy[0, 0, 0], zz[0, 1, 0] - zz[0, 0, 0]])) if shape[1] > 1 else 1.0
    spacing_z = float(np.linalg.norm([xx[0, 0, 1] - xx[0, 0, 0], yy[0, 0, 1] - yy[0, 0, 0], zz[0, 0, 1] - zz[0, 0, 0]])) if shape[2] > 1 else 1.0
    image.spacing = (spacing_x, spacing_y, spacing_z)

    image.point_data["Density"] = np.ascontiguousarray(CUBdat).ravel(order="F")
    image.point_data["ESP"] = np.ascontiguousarray(CUBdatESP).ravel(order="F")
    image.set_active_scalars("Density")
    return image


def update_status(message, color="blue"):
    label = APP_STATE.get("status_label")
    if label is not None:
        label.config(text=message, fg=color)
        try:
            label.update_idletasks()
        except Exception:
            pass


def set_viewer_controls_state(enabled):
    state = "normal" if enabled else "disabled"
    for widget in APP_STATE.get("viewer_controls", []):
        try:
            widget.config(state=state)
        except Exception:
            pass


def refresh_extrema_panel(lines=None):
    text_widget = APP_STATE.get("extrema_text")
    if text_widget is None:
        return
    try:
        text_widget.delete("1.0", tk.END)
        if lines:
            text_widget.insert(tk.END, "\n".join(lines))
        else:
            text_widget.insert(tk.END, "No extrema loaded.\n")
    except Exception:
        pass


def _cleanup_viewer_state():
    global VIEWER_STATE
    VIEWER_STATE = None
    set_viewer_controls_state(False)


def _viewer_is_alive():
    if VIEWER_STATE is None:
        return False
    plotter = VIEWER_STATE.get("plotter")
    return plotter is not None and getattr(plotter, "ren_win", None) is not None


def close_viewer():
    global VIEWER_STATE
    if VIEWER_STATE is None:
        return
    plotter = VIEWER_STATE.get("plotter")
    try:
        if plotter is not None and getattr(plotter, "ren_win", None) is not None:
            plotter.close()
    except Exception:
        pass
    _cleanup_viewer_state()


def pump_viewer():
    root = APP_STATE.get("root")
    if root is None:
        return
    if VIEWER_STATE is not None:
        plotter = VIEWER_STATE.get("plotter")
        try:
            if plotter is None or getattr(plotter, "ren_win", None) is None:
                _cleanup_viewer_state()
            else:
                plotter.update()
        except Exception:
            _cleanup_viewer_state()
    root.after(30, pump_viewer)


def viewer_apply_isovalue(*_args):
    if VIEWER_STATE is None:
        return
    value_text = APP_STATE["isovalue_var"].get().strip()
    try:
        value = float(value_text)
    except ValueError:
        messagebox.showerror("Invalid isovalue", "Density isovalue must be a numeric value.", parent=APP_STATE.get("root"))
        APP_STATE["isovalue_var"].set(f"{VIEWER_STATE['state']['isovalue']:.6g}")
        return
    VIEWER_STATE["state"]["isovalue"] = value
    VIEWER_STATE["rebuild_surface"]()


def viewer_apply_esp_range(*_args):
    if VIEWER_STATE is None:
        return
    min_text = APP_STATE["esp_min_var"].get().strip()
    max_text = APP_STATE["esp_max_var"].get().strip()

    if not min_text and not max_text:
        VIEWER_STATE["state"]["esp_use_custom_range"] = False
        VIEWER_STATE["rebuild_surface"]()
        update_status("ESP scale bar range reset to the original Multiwfn data range.", "darkgreen")
        return

    try:
        esp_min = float(min_text)
        esp_max = float(max_text)
    except ValueError:
        messagebox.showerror("Invalid ESP range", "Scale bar min and max must be numeric values.", parent=APP_STATE.get("root"))
        if VIEWER_STATE["state"].get("esp_use_custom_range"):
            APP_STATE["esp_min_var"].set(f"{VIEWER_STATE['state']['esp_min']:.6g}")
            APP_STATE["esp_max_var"].set(f"{VIEWER_STATE['state']['esp_max']:.6g}")
        else:
            APP_STATE["esp_min_var"].set("")
            APP_STATE["esp_max_var"].set("")
        return

    if esp_min >= esp_max:
        messagebox.showerror("Invalid ESP range", "Scale bar min must be smaller than max.", parent=APP_STATE.get("root"))
        if VIEWER_STATE["state"].get("esp_use_custom_range"):
            APP_STATE["esp_min_var"].set(f"{VIEWER_STATE['state']['esp_min']:.6g}")
            APP_STATE["esp_max_var"].set(f"{VIEWER_STATE['state']['esp_max']:.6g}")
        else:
            APP_STATE["esp_min_var"].set("")
            APP_STATE["esp_max_var"].set("")
        return

    VIEWER_STATE["state"]["esp_min"] = esp_min
    VIEWER_STATE["state"]["esp_max"] = esp_max
    VIEWER_STATE["state"]["esp_use_custom_range"] = True
    VIEWER_STATE["rebuild_surface"]()
    update_status(f"ESP scale bar range override set to {esp_min:.3f} ... {esp_max:.3f}", "darkgreen")


def viewer_update_opacity(value):
    if VIEWER_STATE is None:
        return
    VIEWER_STATE["state"]["opacity"] = float(value) / 100.0
    VIEWER_STATE["rebuild_surface"]()


def viewer_update_cmap(value):
    if VIEWER_STATE is None:
        return
    VIEWER_STATE["state"]["cmap_index"] = VIEWER_STATE["cmap_list"].index(value)
    VIEWER_STATE["rebuild_surface"]()


def viewer_toggle_atoms():
    if VIEWER_STATE is None:
        return
    VIEWER_STATE["state"]["show_atoms"] = bool(APP_STATE["show_atoms_var"].get())
    VIEWER_STATE["build_atoms"]()


def viewer_toggle_bonds():
    if VIEWER_STATE is None:
        return
    VIEWER_STATE["state"]["show_bonds"] = bool(APP_STATE["show_bonds_var"].get())
    VIEWER_STATE["build_bonds"]()


def viewer_save_as():
    if not _viewer_is_alive():
        return
    file_path = filedialog.asksaveasfilename(
        title="Save current view",
        defaultextension=".png",
        filetypes=[("PNG image", "*.png"), ("JPEG image", "*.jpg *.jpeg")],
        parent=APP_STATE.get("root"),
    )
    if not file_path:
        return
    VIEWER_STATE["plotter"].screenshot(file_path, window_size=VIEWER_STATE["plotter"].window_size, scale=1)
    update_status(f"Saved image: {file_path}", "darkgreen")


def _copy_image_to_clipboard_windows(rgb_image):
    try:
        import io
        from PIL import Image
        import win32clipboard
    except Exception as exc:
        raise RuntimeError("Clipboard image copy requires Pillow and pywin32 on Windows.") from exc

    image = Image.fromarray(rgb_image)
    output = io.BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()


def viewer_copy_to_clipboard():
    if not _viewer_is_alive():
        return
    if os.name != "nt":
        update_status("Ctrl+C image copy is currently implemented for Windows only.", "red")
        return
    try:
        rgb_image = VIEWER_STATE["plotter"].screenshot(return_img=True, window_size=VIEWER_STATE["plotter"].window_size, scale=1)
        _copy_image_to_clipboard_windows(rgb_image)
        update_status("Current PyVista view copied to clipboard.", "darkgreen")
    except Exception as exc:
        update_status(f"Clipboard copy failed: {exc}", "red")


def _set_scalar_bar_style(plotter):
    try:
        scalar_bar = plotter.scalar_bars.get("ESP, kcal/mol\n\n\n") or plotter.scalar_bars.get("ESP, kcal/mol")
        if scalar_bar is None:
            return
        title_prop = scalar_bar.GetTitleTextProperty()
        label_prop = scalar_bar.GetLabelTextProperty()
        title_prop.BoldOn()
        label_prop.BoldOn()
        title_prop.SetFontSize(24)
        label_prop.SetFontSize(20)
        title_prop.SetVerticalJustificationToBottom()
    except Exception:
        pass


def _extrema_to_lines(points):
    lines = []
    for idx, point in enumerate(points, start=1):
        lines.append(f"{idx:>3d}  {point[0]:>10.2f}  x={point[1]:>8.3f}  y={point[2]:>8.3f}  z={point[3]:>8.3f}")
    return lines


def _get_viewer_colors():
    bg = APP_STATE.get("bg_color_var")
    label = APP_STATE.get("label_color_var")
    return (bg.get() if bg else "black", label.get() if label else "white")


def _parse_extrema_lines_from_widget():
    text_widget = APP_STATE.get("extrema_text")
    if text_widget is None:
        return None
    raw = text_widget.get("1.0", tk.END).strip()
    if not raw or raw == "No extrema loaded.":
        return []
    parsed = []
    import re
    number_re = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")
    for line in raw.splitlines():
        nums = number_re.findall(line)
        if len(nums) < 4:
            continue
        if len(nums) >= 5:
            nums = nums[-4:]
        try:
            parsed.append([float(nums[0]), float(nums[1]), float(nums[2]), float(nums[3])])
        except Exception:
            continue
    return parsed


def _remove_extrema_actors():
    if VIEWER_STATE is None:
        return
    plotter = VIEWER_STATE["plotter"]
    for key in ["extrema_points_actor", "extrema_labels_actor"]:
        actor = VIEWER_STATE.get(key)
        if actor is not None:
            try:
                plotter.remove_actor(actor, render=False)
            except Exception:
                pass
            VIEWER_STATE[key] = None


def _render_extrema(points):
    if VIEWER_STATE is None:
        return
    plotter = VIEWER_STATE["plotter"]
    _remove_extrema_actors()
    VIEWER_STATE["extrema_points"] = list(points)
    refresh_extrema_panel(_extrema_to_lines(points))
    if not points:
        plotter.render()
        return
    coords = np.array([p[1:4] for p in points], dtype=float)
    poly = VIEWER_STATE["pv"].PolyData(coords)
    labels = [f"{p[0]:.2f}" for p in points]
    _bg_color, label_color = _get_viewer_colors()
    VIEWER_STATE["extrema_points_actor"] = plotter.add_points(
        poly,
        color=label_color,
        point_size=12,
        render_points_as_spheres=True,
        name="extrema_points",
        render=False,
    )
    VIEWER_STATE["extrema_labels_actor"] = plotter.add_point_labels(
        poly,
        labels,
        font_size=16,
        point_size=0,
        text_color=label_color,
        fill_shape=False,
        shape=None,
        margin=0,
        always_visible=True,
        name="extrema_labels",
        render=False,
    )
    plotter.render()


def viewer_generate_extrema(scan_near_ua=False):
    if VIEWER_STATE is None:
        return
    isoval = VIEWER_STATE["state"]["isovalue"]
    points = CalcPoints(isoval)
    if scan_near_ua:
        filtered = []
        for atom in VIEWER_STATE["centers"]:
            if atom[0] in [1, 8, 9, 16, 17, 34, 35, 52, 53, 84, 85]:
                for cp in points:
                    if np.linalg.norm(np.array(cp[1:]) - np.array(atom[1:-1])) < 5.0 and cp not in filtered:
                        filtered.append(cp)
        points = filtered
    _render_extrema(points)
    update_status(f"Loaded {len(points)} extrema point(s).", "darkgreen")


def viewer_clear_extrema():
    if VIEWER_STATE is None:
        return
    VIEWER_STATE["extrema_points"] = []
    _remove_extrema_actors()
    refresh_extrema_panel([])
    try:
        VIEWER_STATE["plotter"].render()
    except Exception:
        pass


def viewer_delete_selected_extrema():
    if VIEWER_STATE is None:
        return
    text_widget = APP_STATE.get("extrema_text")
    if text_widget is None:
        return
    ranges = list(text_widget.tag_ranges(tk.SEL))
    if not ranges:
        messagebox.showinfo("Delete extrema", "Select one or more lines in the extrema list first.", parent=APP_STATE.get("root"))
        return
    start = str(ranges[0])
    end = str(ranges[-1])
    first_line = int(start.split(".")[0])
    last_line = int(end.split(".")[0])
    current = VIEWER_STATE.get("extrema_points", [])
    kept = [p for idx, p in enumerate(current, start=1) if not (first_line <= idx <= last_line)]
    _render_extrema(kept)
    update_status(f"Removed {len(current) - len(kept)} extrema point(s) from list and screen.", "darkgreen")


def viewer_apply_edited_extrema():
    if VIEWER_STATE is None:
        return
    parsed = _parse_extrema_lines_from_widget()
    if parsed is None:
        return
    _render_extrema(parsed)
    update_status(f"Applied edited extrema list: {len(parsed)} point(s) kept on screen.", "darkgreen")


def viewer_kill_range():
    if VIEWER_STATE is None:
        return
    points = VIEWER_STATE.get("extrema_points", [])
    if not points:
        return
    try:
        killval = float(APP_STATE["kill_value_var"].get().strip())
        killpm = float(APP_STATE["kill_pm_var"].get().strip())
    except ValueError:
        messagebox.showerror("Invalid kill range", "Kill value and ± range must be numeric.", parent=APP_STATE.get("root"))
        return
    kept = [p for p in points if not (killval - killpm < p[0] < killval + killpm)]
    removed = len(points) - len(kept)
    _render_extrema(kept)
    update_status(f"Removed {removed} extrema point(s).", "darkgreen")


def viewer_choose_background_color():
    if VIEWER_STATE is None:
        return
    chosen = colorchooser.askcolor(title="Choose background color", parent=APP_STATE.get("root"))
    if not chosen or not chosen[1]:
        return
    APP_STATE["bg_color_var"].set(chosen[1])
    VIEWER_STATE["apply_colors"]()


def viewer_choose_label_color():
    if VIEWER_STATE is None:
        return
    chosen = colorchooser.askcolor(title="Choose label color", parent=APP_STATE.get("root"))
    if not chosen or not chosen[1]:
        return
    APP_STATE["label_color_var"].set(chosen[1])
    VIEWER_STATE["apply_colors"]()


def viewer_toggle_molecule_offset():
    if VIEWER_STATE is None:
        return
    VIEWER_STATE["state"]["molecule_offset_mode"] = bool(APP_STATE["molecule_offset_var"].get())
    VIEWER_STATE["build_atoms"]()
    VIEWER_STATE["build_bonds"]()


def viewer_reset():

    if not _viewer_is_alive():
        return
    VIEWER_STATE["plotter"].reset_camera()
    VIEWER_STATE["plotter"].render()


def sync_main_controls_from_viewer():
    if VIEWER_STATE is None:
        set_viewer_controls_state(False)
        return
    state = VIEWER_STATE["state"]
    APP_STATE["isovalue_var"].set(f"{state['isovalue']:.6g}")
    APP_STATE["opacity_scale"].set(int(round(state["opacity"] * 100.0)))
    APP_STATE["cmap_var"].set(VIEWER_STATE["cmap_list"][state["cmap_index"]])
    APP_STATE["show_atoms_var"].set(state["show_atoms"])
    APP_STATE["show_bonds_var"].set(state["show_bonds"])
    APP_STATE["molecule_offset_var"].set(state.get("molecule_offset_mode", False))
    APP_STATE["suggested_range_var"].set(
        f"Suggested density range: {VIEWER_STATE['dens_min']:.6g} to {VIEWER_STATE['dens_max']:.6g}"
    )
    APP_STATE["esp_range_hint_var"].set(
        f"ESP data range: {VIEWER_STATE['esp_min_data']:.6g} to {VIEWER_STATE['esp_max_data']:.6g}"
    )
    if state.get("esp_use_custom_range"):
        APP_STATE["esp_min_var"].set(f"{state['esp_min']:.6g}")
        APP_STATE["esp_max_var"].set(f"{state['esp_max']:.6g}")
    else:
        APP_STATE["esp_min_var"].set("")
        APP_STATE["esp_max_var"].set("")
    set_viewer_controls_state(True)


def VisualizeData(CENTERS, CUBdat, CUBdatESP, xx, yy, zz):
    import pyvista as pv

    global VIEWER_STATE

    if VIEWER_STATE is not None:
        close_viewer()

    grid = BuildPyVistaGrid(CUBdat, CUBdatESP, xx, yy, zz)
    dens_min = float(np.min(CUBdat))
    dens_max = float(np.max(CUBdat))
    esp_min_data = float(np.min(CUBdatESP))
    esp_max_data = float(np.max(CUBdatESP))
    bond_pairs = BuildBondPairs(CENTERS)

    cmap_list = [
        "gist_rainbow",
        "turbo",
        "coolwarm",
        "RdBu_r",
        "plasma",
        "viridis",
    ]

    state = {
        "isovalue": 0.001,
        "opacity": 0.5,
        "cmap_index": 0,
        "show_atoms": True,
        "show_bonds": True,
        "molecule_offset_mode": False,
        "esp_min": esp_min_data,
        "esp_max": esp_max_data,
        "esp_use_custom_range": False,
        "surface_actor": None,
        "atom_actors": [],
        "bond_actors": [],
    }

    viewer_width, viewer_height = _viewer_window_size()
    plotter = pv.Plotter(window_size=(viewer_width, viewer_height))
    try:
        plotter.enable_anti_aliasing("msaa")
    except Exception:
        pass

    def _scaled_point(point, centroid, factor):
        return centroid + (point - centroid) * factor

    def _molecule_scale_factor():
        return 1.06 if state.get("molecule_offset_mode") else 1.0

    def apply_colors():
        bg_color, label_color = _get_viewer_colors()
        try:
            plotter.set_background(bg_color)
        except Exception:
            pass
        if VIEWER_STATE is not None and VIEWER_STATE.get("extrema_points"):
            _render_extrema(VIEWER_STATE.get("extrema_points", []))
        else:
            try:
                plotter.render()
            except Exception:
                pass

    def rebuild_surface():
        contour = grid.contour(isosurfaces=[float(state["isovalue"])], scalars="Density")
        if contour.n_points == 0:
            if state["surface_actor"] is not None:
                plotter.remove_actor(state["surface_actor"], render=False)
                state["surface_actor"] = None
            plotter.add_text("No surface at current isovalue", position="lower_right", font_size=10, name="surface_status")
            plotter.render()
            return
        if state["surface_actor"] is not None:
            plotter.remove_actor(state["surface_actor"], render=False)
        plotter.add_text(" " * 32, position="lower_right", font_size=10, name="surface_status")
        mesh_kwargs = dict(
            scalars="ESP",
            cmap=cmap_list[state["cmap_index"]],
            opacity=float(state["opacity"]),
            smooth_shading=True,
            scalar_bar_args={
                "title": "ESP, kcal/mol\n\n\n",
                "vertical": True,
                "position_x": 0.02,
                "position_y": 0.07,
                "width": 0.08,
                "height": 0.86,
                "title_font_size": 26,
                "label_font_size": 22,
                "n_labels": 5,
                "fmt": "%.1f",
                "color": APP_STATE.get("label_color_var", tk.StringVar(value="white")).get(),
            },
            name="esp_surface",
            render=False,
        )
        if state.get("esp_use_custom_range"):
            mesh_kwargs["clim"] = [float(state["esp_min"]), float(state["esp_max"])]

        state["surface_actor"] = plotter.add_mesh(
            contour,
            **mesh_kwargs
        )
        _set_scalar_bar_style(plotter)
        plotter.render()

    def build_atoms():
        for actor in state["atom_actors"]:
            plotter.remove_actor(actor, render=False)
        state["atom_actors"] = []
        if not state["show_atoms"]:
            plotter.render()
            return
        coords = np.array([atom[1:4] for atom in CENTERS], dtype=float)
        centroid = np.mean(coords, axis=0) if len(coords) else np.zeros(3)
        scale_factor = _molecule_scale_factor()
        for atom in CENTERS:
            center = np.array(atom[1:4], dtype=float)
            if scale_factor != 1.0:
                center = _scaled_point(center, centroid, scale_factor)
            radius = max(0.18, 0.28 * float(dnc2all[atom[0]][1]))
            color = tuple(float(c) for c in dnc2all[atom[0]][-3:])
            sphere = pv.Sphere(radius=radius, center=center, theta_resolution=28, phi_resolution=28)
            actor = plotter.add_mesh(sphere, color=color, smooth_shading=True, render=False)
            state["atom_actors"].append(actor)
        plotter.render()

    def build_bonds():
        for actor in state["bond_actors"]:
            plotter.remove_actor(actor, render=False)
        state["bond_actors"] = []
        if not state["show_bonds"]:
            plotter.render()
            return
        coords = np.array([atom[1:4] for atom in CENTERS], dtype=float)
        centroid = np.mean(coords, axis=0) if len(coords) else np.zeros(3)
        scale_factor = _molecule_scale_factor()
        for i, j in bond_pairs:
            p1 = np.array(CENTERS[i][1:4], dtype=float)
            p2 = np.array(CENTERS[j][1:4], dtype=float)
            if scale_factor != 1.0:
                p1 = _scaled_point(p1, centroid, scale_factor)
                p2 = _scaled_point(p2, centroid, scale_factor)
            line = pv.Line(p1, p2, resolution=1)
            tube = line.tube(radius=0.10)
            actor = plotter.add_mesh(tube, color="lightgray", smooth_shading=True, render=False)
            state["bond_actors"].append(actor)
        plotter.render()

    VIEWER_STATE = {
        "pv": pv,
        "plotter": plotter,
        "state": state,
        "rebuild_surface": rebuild_surface,
        "build_atoms": build_atoms,
        "build_bonds": build_bonds,
        "apply_colors": apply_colors,
        "cmap_list": cmap_list,
        "dens_min": dens_min,
        "dens_max": dens_max,
        "esp_min_data": esp_min_data,
        "esp_max_data": esp_max_data,
        "centers": CENTERS,
        "extrema_points": [],
        "extrema_points_actor": None,
        "extrema_labels_actor": None,
    }

    try:
        interactor = getattr(plotter, "iren", None)
        vtk_interactor = getattr(interactor, "interactor", None)
        if vtk_interactor is not None:
            def _ctrl_c_observer(_obj, _event):
                try:
                    key = vtk_interactor.GetKeySym()
                    ctrl = vtk_interactor.GetControlKey()
                except Exception:
                    key = None
                    ctrl = 0
                if ctrl and str(key).lower() == "c":
                    viewer_copy_to_clipboard()
            vtk_interactor.AddObserver("KeyPressEvent", _ctrl_c_observer)
        else:
            plotter.add_key_event("c", viewer_copy_to_clipboard)
    except Exception:
        pass

    apply_colors()
    rebuild_surface()
    build_atoms()
    build_bonds()
    plotter.show(title="VisMap PyVista Viewer", auto_close=False, interactive_update=True)
    sync_main_controls_from_viewer()
    refresh_extrema_panel([])


# ----------------------------
# Main processing

# ----------------------------
# Main processing
# ----------------------------

def process_selected_file(selected_inputfile, selected_nproc="4", selected_mode="old", selected_vis="y",
                          selected_pregen=False, selected_cpisov=None, selected_multiwfn=None):
    global inputfile, fname, nproc, mode, vis, PreGenCP, CPisov, workdir, mwfn_exe

    inputfile = os.path.abspath(selected_inputfile)
    fname = base_name_no_ext(inputfile)
    workdir = os.path.dirname(inputfile)

    mwfn_exe = os.path.abspath(selected_multiwfn) if selected_multiwfn else find_multiwfn_path()

    if not os.path.exists(inputfile):
        raise FileNotFoundError(f"Input file not found:\n{inputfile}")

    if not os.path.exists(mwfn_exe):
        raise FileNotFoundError(
            "Multiwfn.exe not found.\n\n"
            "Set the correct path in the GUI, or edit DEFAULT_MULTIWFN_PATHS in the script."
        )

    mode = selected_mode if selected_mode in ["new", "old"] else "old"
    vis = selected_vis if selected_vis in ["y", "n"] else "y"

    try:
        nproc = str(int(selected_nproc))
    except ValueError:
        print("Could not convert given nproc to integer. Using 4.")
        nproc = "4"

    PreGenCP = bool(selected_pregen)
    CPisov = selected_cpisov

    if PreGenCP:
        try:
            CPisov = str(float(CPisov))
        except Exception:
            print("Error reading CPisov. Canceling ECP CPs pre-generation.")
            PreGenCP = False
            CPisov = None

    if mode == "old":
        IsCalced = [True, True]
        for i, x in enumerate([fname + "_Dens.cub", fname + "_ESP.cub"]):
            if os.path.exists(x):
                print("Located", x)
            else:
                IsCalced[i] = False
                print("Could not locate", x, "- It will be (re)calculated")
    else:
        IsCalced = [False, False]

    CalcCub(IsCalced, fname)

    if PreGenCP:
        CalcPoints(CPisov)

    CUB, CENTERS = ReadCUB(fname + "_Dens.cub")
    origin, v1, v2, v3, pointsv1, pointsv2, pointsv3, Scalars = CUB

    CUBdat = np.empty((pointsv1, pointsv2, pointsv3))
    for x in range(pointsv1):
        for y in range(pointsv2):
            for z in range(pointsv3):
                CUBdat[x][y][z] = Scalars[x * pointsv2 * pointsv3 + y * pointsv3 + z]

    x = np.array([origin[0] + v1[0] * i for i in range(pointsv1)])
    y = np.array([origin[1] + v2[1] * i for i in range(pointsv2)])
    z = np.array([origin[2] + v3[2] * i for i in range(pointsv3)])
    xx, yy, zz = np.meshgrid(x, y, z, indexing="ij")

    CUB, CENTERS = ReadCUB(fname + "_ESP.cub")
    TotESP = CUB[-1]
    CUBdatESP = np.empty((pointsv1, pointsv2, pointsv3))
    for x in range(pointsv1):
        for y in range(pointsv2):
            for z in range(pointsv3):
                CUBdatESP[x][y][z] = TotESP[x * pointsv2 * pointsv3 + y * pointsv3 + z] * 627.0

    if vis == "y":
        VisualizeData(CENTERS, CUBdat, CUBdatESP, xx, yy, zz)
        update_status("Viewer loaded.", "darkgreen")


# ----------------------------
# GUI
# ----------------------------

def launch_gui():
    root = tk.Tk()
    APP_STATE["root"] = root
    root.title("VisMap GUI")
    gui_width, gui_height = _main_window_size()
    root.geometry(f"{gui_width}x{gui_height}")
    root.resizable(True, True)

    frm = tk.Frame(root, padx=10, pady=10)
    frm.pack(fill="both", expand=True)
    frm.grid_columnconfigure(0, weight=3)
    frm.grid_columnconfigure(1, weight=2)
    frm.grid_rowconfigure(3, weight=1)

    input_box = tk.LabelFrame(frm, text="Input and execution", padx=10, pady=10)
    input_box.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
    input_box.grid_columnconfigure(0, weight=1)

    tk.Label(input_box, text="Input wavefunction file (.wfn / .wfx / .fchk)").grid(row=0, column=0, sticky="w")
    entry_file = tk.Entry(input_box, width=110)
    entry_file.grid(row=1, column=0, padx=(0, 8), pady=(2, 10), sticky="we")

    def browse_file():
        f = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[("Wavefunction files", "*.wfn *.wfx *.fchk"), ("All files", "*.*")],
        )
        if f:
            entry_file.delete(0, tk.END)
            entry_file.insert(0, f)

    tk.Button(input_box, text="Browse", command=browse_file, width=12).grid(row=1, column=1, sticky="w")

    tk.Label(input_box, text="Multiwfn.exe").grid(row=2, column=0, sticky="w")
    entry_mwfn = tk.Entry(input_box, width=110)
    entry_mwfn.grid(row=3, column=0, padx=(0, 8), pady=(2, 0), sticky="we")
    entry_mwfn.insert(0, r"C:/Multiwfn_2026.2.2_bin_Win64/Multiwfn.exe")

    def browse_mwfn():
        f = filedialog.askopenfilename(title="Select Multiwfn.exe", filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if f:
            entry_mwfn.delete(0, tk.END)
            entry_mwfn.insert(0, f)

    tk.Button(input_box, text="Browse", command=browse_mwfn, width=12).grid(row=3, column=1, sticky="w")

    options = tk.LabelFrame(frm, text="Calculation options", padx=10, pady=10)
    options.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))

    tk.Label(options, text="nproc").grid(row=0, column=0, sticky="w")
    entry_nproc = tk.Entry(options, width=8)
    entry_nproc.grid(row=1, column=0, padx=(0, 14), sticky="w")
    entry_nproc.insert(0, "8")

    tk.Label(options, text="mode").grid(row=0, column=1, sticky="w")
    mode_var = tk.StringVar(value="old")
    tk.OptionMenu(options, mode_var, "old", "new").grid(row=1, column=1, padx=(0, 14), sticky="w")

    tk.Label(options, text="visualization").grid(row=0, column=2, sticky="w")
    vis_var = tk.StringVar(value="y")
    tk.OptionMenu(options, vis_var, "y", "n").grid(row=1, column=2, padx=(0, 14), sticky="w")

    preg_var = tk.BooleanVar(value=False)
    tk.Checkbutton(options, text="Pre-generate CP points", variable=preg_var).grid(row=0, column=3, columnspan=2, sticky="w")

    tk.Label(options, text="CP isovalue").grid(row=1, column=3, sticky="e", padx=(10, 4))
    entry_cp = tk.Entry(options, width=10)
    entry_cp.grid(row=1, column=4, sticky="w")
    entry_cp.insert(0, "0.001")

    action_box = tk.LabelFrame(frm, text="Viewer controls (main UI)", padx=10, pady=10)
    action_box.grid(row=1, column=1, sticky="nsew", pady=(0, 8))
    action_box.grid_columnconfigure(1, weight=1)

    APP_STATE["suggested_range_var"] = tk.StringVar(value="Suggested density range: n/a")
    APP_STATE["esp_range_hint_var"] = tk.StringVar(value="ESP data range: n/a")
    APP_STATE["isovalue_var"] = tk.StringVar(value="0.001")
    APP_STATE["esp_min_var"] = tk.StringVar(value="-50")
    APP_STATE["esp_max_var"] = tk.StringVar(value="50")
    APP_STATE["cmap_var"] = tk.StringVar(value="gist_rainbow")
    APP_STATE["show_atoms_var"] = tk.BooleanVar(value=True)
    APP_STATE["show_bonds_var"] = tk.BooleanVar(value=True)
    APP_STATE["molecule_offset_var"] = tk.BooleanVar(value=False)
    APP_STATE["kill_value_var"] = tk.StringVar(value="0.0")
    APP_STATE["kill_pm_var"] = tk.StringVar(value="1.0")
    APP_STATE["bg_color_var"] = tk.StringVar(value="black")
    APP_STATE["label_color_var"] = tk.StringVar(value="white")

    tk.Label(action_box, text="Density isovalue").grid(row=0, column=0, sticky="w", pady=(0, 6))
    isovalue_entry = tk.Entry(action_box, textvariable=APP_STATE["isovalue_var"], width=14)
    isovalue_entry.grid(row=0, column=1, sticky="w", pady=(0, 6))
    isovalue_entry.bind("<Return>", viewer_apply_isovalue)
    apply_btn = tk.Button(action_box, text="Apply", command=viewer_apply_isovalue, width=10)
    apply_btn.grid(row=0, column=2, sticky="w", padx=(6, 0), pady=(0, 6))

    tk.Label(action_box, textvariable=APP_STATE["suggested_range_var"]).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 2))
    tk.Label(action_box, textvariable=APP_STATE["esp_range_hint_var"]).grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))

    tk.Label(action_box, text="Scale bar min").grid(row=3, column=0, sticky="w")
    esp_min_entry = tk.Entry(action_box, textvariable=APP_STATE["esp_min_var"], width=14)
    esp_min_entry.grid(row=3, column=1, sticky="w")
    esp_min_entry.bind("<Return>", viewer_apply_esp_range)
    tk.Label(action_box, text="Scale bar max").grid(row=4, column=0, sticky="w")
    esp_max_entry = tk.Entry(action_box, textvariable=APP_STATE["esp_max_var"], width=14)
    esp_max_entry.grid(row=4, column=1, sticky="w")
    esp_max_entry.bind("<Return>", viewer_apply_esp_range)
    esp_apply_btn = tk.Button(action_box, text="Apply range", command=viewer_apply_esp_range, width=10)
    esp_apply_btn.grid(row=3, column=2, rowspan=2, sticky="w")

    tk.Label(action_box, text="Opacity").grid(row=5, column=0, sticky="w")
    opacity_scale = tk.Scale(action_box, from_=0, to=100, orient="horizontal", resolution=1, command=viewer_update_opacity, showvalue=True, length=260)
    opacity_scale.set(50)
    opacity_scale.grid(row=5, column=1, columnspan=2, sticky="we", pady=(0, 8))
    APP_STATE["opacity_scale"] = opacity_scale

    tk.Label(action_box, text="Colormap").grid(row=6, column=0, sticky="w")
    cmap_menu = tk.OptionMenu(action_box, APP_STATE["cmap_var"], "gist_rainbow", "turbo", "coolwarm", "RdBu_r", "plasma", "viridis", command=viewer_update_cmap)
    cmap_menu.grid(row=6, column=1, sticky="w", pady=(0, 8))

    atoms_cb = tk.Checkbutton(action_box, text="Show atoms", variable=APP_STATE["show_atoms_var"], command=viewer_toggle_atoms)
    atoms_cb.grid(row=7, column=0, sticky="w")
    bonds_cb = tk.Checkbutton(action_box, text="Show bonds", variable=APP_STATE["show_bonds_var"], command=viewer_toggle_bonds)
    bonds_cb.grid(row=7, column=1, sticky="w")
    offset_cb = tk.Checkbutton(action_box, text="Molecule level-up mode", variable=APP_STATE["molecule_offset_var"], command=viewer_toggle_molecule_offset)
    offset_cb.grid(row=7, column=2, sticky="w")

    bg_btn = tk.Button(action_box, text="Background color", command=viewer_choose_background_color, width=16)
    bg_btn.grid(row=8, column=0, sticky="w", pady=(8, 6))
    label_btn = tk.Button(action_box, text="Labels color", command=viewer_choose_label_color, width=12)
    label_btn.grid(row=8, column=1, sticky="w", pady=(8, 6))

    save_btn = tk.Button(action_box, text="Save as...", command=viewer_save_as, width=12)
    save_btn.grid(row=9, column=0, sticky="w", pady=(4, 0))
    copy_btn = tk.Button(action_box, text="Copy image", command=viewer_copy_to_clipboard, width=12)
    copy_btn.grid(row=9, column=1, sticky="w", pady=(4, 0))
    reset_btn = tk.Button(action_box, text="Reset view", command=viewer_reset, width=12)
    reset_btn.grid(row=9, column=2, sticky="w", pady=(4, 0))
    close_btn = tk.Button(action_box, text="Close viewer", command=close_viewer, width=12)
    close_btn.grid(row=10, column=0, sticky="w", pady=(4, 0))

    main_area = tk.LabelFrame(frm, text="Extrema tools and values", padx=10, pady=10)
    main_area.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
    main_area.grid_columnconfigure(0, weight=0)
    main_area.grid_columnconfigure(1, weight=1)
    main_area.grid_rowconfigure(0, weight=1)

    left_tools = tk.Frame(main_area)
    left_tools.grid(row=0, column=0, sticky="nsw", padx=(0, 12))

    gen_btn = tk.Button(left_tools, text="Generate extrema", command=lambda: viewer_generate_extrema(False), width=18)
    gen_btn.grid(row=0, column=0, sticky="we", pady=(0, 6))
    scan_btn = tk.Button(left_tools, text="Scan near UA", command=lambda: viewer_generate_extrema(True), width=18)
    scan_btn.grid(row=1, column=0, sticky="we", pady=(0, 6))
    clear_btn = tk.Button(left_tools, text="Clear extrema", command=viewer_clear_extrema, width=18)
    clear_btn.grid(row=2, column=0, sticky="we", pady=(0, 12))

    tk.Label(left_tools, text="Kill value").grid(row=3, column=0, sticky="w")
    kill_value_entry = tk.Entry(left_tools, textvariable=APP_STATE["kill_value_var"], width=12)
    kill_value_entry.grid(row=4, column=0, sticky="w", pady=(0, 6))
    tk.Label(left_tools, text="± range").grid(row=5, column=0, sticky="w")
    kill_pm_entry = tk.Entry(left_tools, textvariable=APP_STATE["kill_pm_var"], width=12)
    kill_pm_entry.grid(row=6, column=0, sticky="w", pady=(0, 6))
    kill_btn = tk.Button(left_tools, text="Kill range", command=viewer_kill_range, width=18)
    kill_btn.grid(row=7, column=0, sticky="we", pady=(0, 12))

    del_btn = tk.Button(left_tools, text="Delete selected lines", command=viewer_delete_selected_extrema, width=18)
    del_btn.grid(row=8, column=0, sticky="we", pady=(0, 6))
    sync_btn = tk.Button(left_tools, text="Apply edited list", command=viewer_apply_edited_extrema, width=18)
    sync_btn.grid(row=9, column=0, sticky="we", pady=(0, 6))

    right_panel = tk.Frame(main_area)
    right_panel.grid(row=0, column=1, sticky="nsew")
    right_panel.grid_columnconfigure(0, weight=1)
    right_panel.grid_rowconfigure(1, weight=1)

    tk.Label(right_panel, text="Extrema values (editable; delete lines and click 'Apply edited list')").grid(row=0, column=0, sticky="w", pady=(0, 6))
    extrema_text = tk.Text(right_panel, width=90, height=20, wrap="none")
    extrema_text.grid(row=1, column=0, sticky="nsew")
    APP_STATE["extrema_text"] = extrema_text

    extrema_scroll_y = tk.Scrollbar(right_panel, orient="vertical", command=extrema_text.yview)
    extrema_scroll_y.grid(row=1, column=1, sticky="ns")
    extrema_text.configure(yscrollcommand=extrema_scroll_y.set)
    APP_STATE["extrema_scrollbar"] = extrema_scroll_y

    APP_STATE["viewer_controls"] = [
        isovalue_entry, apply_btn, esp_min_entry, esp_max_entry, esp_apply_btn, opacity_scale, cmap_menu,
        atoms_cb, bonds_cb, offset_cb, bg_btn, label_btn, save_btn, copy_btn, reset_btn, close_btn,
        gen_btn, scan_btn, clear_btn, kill_value_entry, kill_pm_entry, kill_btn, del_btn, sync_btn
    ]
    set_viewer_controls_state(False)

    status = tk.Label(frm, text="", anchor="w", justify="left", fg="blue")
    status.grid(row=4, column=0, columnspan=2, sticky="we", pady=(4, 0))
    APP_STATE["status_label"] = status

    button_bar = tk.Frame(frm)
    button_bar.grid(row=5, column=0, columnspan=2, sticky="we", pady=(10, 0))

    def run_clicked():
        selected_file = entry_file.get().strip()
        selected_mwfn = entry_mwfn.get().strip()
        selected_nproc = entry_nproc.get().strip() or "4"
        selected_mode = mode_var.get().strip()
        selected_vis = vis_var.get().strip()
        selected_pregen = preg_var.get()
        selected_cp = entry_cp.get().strip()

        if not selected_file:
            messagebox.showerror("Error", "Please select an input .wfn/.wfx/.fchk file.")
            return

        update_status("Running...", "blue")
        root.update_idletasks()

        try:
            process_selected_file(
                selected_inputfile=selected_file,
                selected_nproc=selected_nproc,
                selected_mode=selected_mode,
                selected_vis=selected_vis,
                selected_pregen=selected_pregen,
                selected_cpisov=selected_cp,
                selected_multiwfn=selected_mwfn,
            )
            if selected_vis != "y":
                update_status("Done.", "darkgreen")
                messagebox.showinfo("Finished", "Processing completed successfully.")
        except Exception as e:
            update_status("Failed.", "red")
            messagebox.showerror("Execution error", str(e))

    tk.Button(button_bar, text="Run VisMap", command=run_clicked, width=18, height=2).pack(side="left")
    tk.Button(button_bar, text="Quit", command=root.destroy, width=12).pack(side="right")

    root.bind_all("<Control-c>", lambda _event: viewer_copy_to_clipboard())
    root.after(30, pump_viewer)
    refresh_extrema_panel([])
    root.mainloop()


# ----------------------------
# CLI support
# ----------------------------

def run_from_cli(argv):
    selected_inputfile = argv[1]
    selected_nproc = "4"
    selected_mode = "old"
    selected_vis = "y"
    selected_pregen = False
    selected_cpisov = None
    selected_multiwfn = None

    for item in argv[2:]:
        if item.startswith("-nproc="):
            selected_nproc = item[7:]
        elif item.startswith("-mode="):
            selected_mode = item[6:]
        elif item.startswith("-vis="):
            selected_vis = item[5:]
        elif item.startswith("-CPisov="):
            selected_cpisov = item[8:]
            selected_pregen = True
        elif item.startswith("-mwfn="):
            selected_multiwfn = item[6:]

    process_selected_file(
        selected_inputfile=selected_inputfile,
        selected_nproc=selected_nproc,
        selected_mode=selected_mode,
        selected_vis=selected_vis,
        selected_pregen=selected_pregen,
        selected_cpisov=selected_cpisov,
        selected_multiwfn=selected_multiwfn
    )


if __name__ == "__main__":
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        run_from_cli(sys.argv)
    else:
        launch_gui()