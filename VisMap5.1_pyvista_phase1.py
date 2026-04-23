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
import time
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox

# ----------------------------
# Config
# ----------------------------

DEFAULT_MULTIWFN_PATHS = [
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


def VisualizeData(CENTERS, CUBdat, CUBdatESP, xx, yy, zz):
    import pyvista as pv

    pv.set_plot_theme("dark")

    grid = BuildPyVistaGrid(CUBdat, CUBdatESP, xx, yy, zz)
    bond_pairs = BuildBondPairs(CENTERS)

    cmap_list = [
        "gist_rainbow",
        "turbo",
        "coolwarm",
        "RdBu_r",
        "plasma",
        "viridis",
    ]

    dens_min = float(np.min(CUBdat))
    dens_max = float(np.max(CUBdat))
    default_isovalue = min(max(0.001, dens_min), dens_max)

    state = {
        "isovalue": default_isovalue,
        "opacity": 0.5,
        "cmap_index": 0,
        "show_atoms": True,
        "show_bonds": True,
        "surface_actor": None,
        "atom_actors": [],
        "bond_actors": [],
        "running": True,
    }

    plotter = pv.Plotter(window_size=(1400, 900))
    plotter.add_axes()
    try:
        plotter.enable_anti_aliasing("msaa")
    except Exception:
        pass
    plotter.add_text(
        "Left drag: rotate | Right drag: zoom | Middle drag/Shift+Left: pan",
        position="upper_right",
        font_size=10,
        name="nav_help",
    )

    def rebuild_surface():
        contour = grid.contour(isosurfaces=[float(state["isovalue"])], scalars="Density")
        if contour.n_points == 0:
            if state["surface_actor"] is not None:
                plotter.remove_actor(state["surface_actor"], render=False)
                state["surface_actor"] = None
            plotter.add_text(
                "No surface at current isovalue",
                position="lower_right",
                font_size=10,
                name="surface_status",
            )
            plotter.render()
            return

        if state["surface_actor"] is not None:
            plotter.remove_actor(state["surface_actor"], render=False)

        plotter.add_text(" " * 32, position="lower_right", font_size=10, name="surface_status")
        state["surface_actor"] = plotter.add_mesh(
            contour,
            scalars="ESP",
            cmap=cmap_list[state["cmap_index"]],
            opacity=float(state["opacity"]),
            smooth_shading=True,
            scalar_bar_args={
                "title": "ESP, kcal/mol",
                "vertical": True,
                "position_x": 0.02,
                "position_y": 0.18,
                "width": 0.08,
                "height": 0.64,
                "title_font_size": 12,
                "label_font_size": 10,
            },
            name="esp_surface",
            render=False,
        )
        plotter.render()

    def build_atoms():
        for actor in state["atom_actors"]:
            plotter.remove_actor(actor, render=False)
        state["atom_actors"] = []
        if not state["show_atoms"]:
            plotter.render()
            return

        for atom in CENTERS:
            center = np.array(atom[1:4], dtype=float)
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

        for i, j in bond_pairs:
            p1 = np.array(CENTERS[i][1:4], dtype=float)
            p2 = np.array(CENTERS[j][1:4], dtype=float)
            line = pv.Line(p1, p2, resolution=1)
            tube = line.tube(radius=0.10)
            actor = plotter.add_mesh(tube, color="lightgray", smooth_shading=True, render=False)
            state["bond_actors"].append(actor)
        plotter.render()

    def update_isovalue_from_entry(*_args):
        value_text = isovalue_var.get().strip()
        try:
            value = float(value_text)
        except ValueError:
            messagebox.showerror("Invalid isovalue", "Density isovalue must be a numeric value.", parent=control_window)
            isovalue_var.set(f"{state['isovalue']:.6g}")
            return

        state["isovalue"] = value
        rebuild_surface()

    def update_opacity(value):
        state["opacity"] = float(value) / 100.0
        rebuild_surface()

    def update_cmap(value):
        idx = cmap_list.index(value)
        state["cmap_index"] = idx
        rebuild_surface()

    def toggle_atoms():
        state["show_atoms"] = bool(show_atoms_var.get())
        build_atoms()

    def toggle_bonds():
        state["show_bonds"] = bool(show_bonds_var.get())
        build_bonds()

    def close_both():
        state["running"] = False
        try:
            if getattr(plotter, "ren_win", None) is not None:
                plotter.close()
        except Exception:
            pass
        try:
            if control_window.winfo_exists():
                control_window.destroy()
        except Exception:
            pass

    rebuild_surface()
    build_atoms()
    build_bonds()

    root = tk._default_root
    if root is None:
        root = tk.Tk()
        root.withdraw()

    control_window = tk.Toplevel(root)
    control_window.title("VisMap controls")
    control_window.geometry("380x220+40+40")
    control_window.resizable(False, False)
    control_window.protocol("WM_DELETE_WINDOW", close_both)

    panel = tk.Frame(control_window, padx=12, pady=12)
    panel.pack(fill="both", expand=True)
    panel.grid_columnconfigure(1, weight=1)

    tk.Label(panel, text="Density isovalue").grid(row=0, column=0, sticky="w", pady=(0, 8))
    isovalue_var = tk.StringVar(value=f"{state['isovalue']:.6g}")
    isovalue_entry = tk.Entry(panel, textvariable=isovalue_var, width=16)
    isovalue_entry.grid(row=0, column=1, sticky="we", pady=(0, 8))
    isovalue_entry.bind("<Return>", update_isovalue_from_entry)
    tk.Button(panel, text="Apply", command=update_isovalue_from_entry, width=10).grid(row=0, column=2, padx=(8, 0), pady=(0, 8))

    tk.Label(panel, text=f"Suggested range: {dens_min:.6g} to {dens_max:.6g}").grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(0, 12)
    )

    tk.Label(panel, text="Opacity").grid(row=2, column=0, sticky="w", pady=(0, 8))
    opacity_scale = tk.Scale(
        panel,
        from_=0,
        to=100,
        orient="horizontal",
        resolution=1,
        command=update_opacity,
        showvalue=True,
        length=220,
    )
    opacity_scale.set(int(round(state["opacity"] * 100.0)))
    opacity_scale.grid(row=2, column=1, columnspan=2, sticky="we", pady=(0, 8))

    tk.Label(panel, text="Colormap").grid(row=3, column=0, sticky="w", pady=(0, 8))
    cmap_var = tk.StringVar(value=cmap_list[state["cmap_index"]])
    cmap_menu = tk.OptionMenu(panel, cmap_var, *cmap_list, command=update_cmap)
    cmap_menu.grid(row=3, column=1, columnspan=2, sticky="we", pady=(0, 8))

    show_atoms_var = tk.BooleanVar(value=state["show_atoms"])
    tk.Checkbutton(panel, text="Show atoms", variable=show_atoms_var, command=toggle_atoms).grid(
        row=4, column=0, columnspan=3, sticky="w", pady=(0, 6)
    )

    show_bonds_var = tk.BooleanVar(value=state["show_bonds"])
    tk.Checkbutton(panel, text="Show bonds", variable=show_bonds_var, command=toggle_bonds).grid(
        row=5, column=0, columnspan=3, sticky="w", pady=(0, 6)
    )

    tk.Label(panel, text="Press Enter or Apply after editing the isovalue.").grid(
        row=6, column=0, columnspan=3, sticky="w", pady=(10, 0)
    )

    plotter.show(title="VisMap PyVista Viewer", auto_close=False, interactive_update=True)

    try:
        while state["running"]:
            if getattr(plotter, "ren_win", None) is None:
                break
            plotter.update()
            try:
                control_window.update_idletasks()
                control_window.update()
            except tk.TclError:
                break
            time.sleep(0.02)
    finally:
        state["running"] = False
        try:
            if getattr(plotter, "ren_win", None) is not None:
                plotter.close()
        except Exception:
            pass
        try:
            if control_window.winfo_exists():
                control_window.destroy()
        except Exception:
            pass


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


# ----------------------------
# GUI
# ----------------------------

def launch_gui():
    root = tk.Tk()
    root.title("VisMap GUI")
    root.geometry("760x300")
    root.resizable(False, False)

    frm = tk.Frame(root, padx=12, pady=12)
    frm.pack(fill="both", expand=True)

    tk.Label(frm, text="Input wavefunction file (.wfn / .wfx / .fchk)").grid(row=0, column=0, sticky="w")
    entry_file = tk.Entry(frm, width=72)
    entry_file.grid(row=1, column=0, padx=(0, 8), pady=(2, 10), sticky="we")

    def browse_file():
        f = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[
                ("Wavefunction files", "*.wfn *.wfx *.fchk"),
                ("All files", "*.*"),
            ]
        )
        if f:
            entry_file.delete(0, tk.END)
            entry_file.insert(0, f)

    tk.Button(frm, text="Browse", command=browse_file).grid(row=1, column=1, sticky="w")

    tk.Label(frm, text="Multiwfn.exe").grid(row=2, column=0, sticky="w")
    entry_mwfn = tk.Entry(frm, width=72)
    entry_mwfn.grid(row=3, column=0, padx=(0, 8), pady=(2, 10), sticky="we")
    entry_mwfn.insert(0, find_multiwfn_path())

    def browse_mwfn():
        f = filedialog.askopenfilename(
            title="Select Multiwfn.exe",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")]
        )
        if f:
            entry_mwfn.delete(0, tk.END)
            entry_mwfn.insert(0, f)

    tk.Button(frm, text="Browse", command=browse_mwfn).grid(row=3, column=1, sticky="w")

    options = tk.Frame(frm)
    options.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 8))

    tk.Label(options, text="nproc").grid(row=0, column=0, sticky="w")
    entry_nproc = tk.Entry(options, width=8)
    entry_nproc.grid(row=1, column=0, padx=(0, 14))
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

    status = tk.Label(frm, text="", anchor="w", justify="left", fg="blue")
    status.grid(row=5, column=0, columnspan=2, sticky="we", pady=(8, 0))

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

        status.config(text="Running...")
        root.update_idletasks()

        try:
            process_selected_file(
                selected_inputfile=selected_file,
                selected_nproc=selected_nproc,
                selected_mode=selected_mode,
                selected_vis=selected_vis,
                selected_pregen=selected_pregen,
                selected_cpisov=selected_cp,
                selected_multiwfn=selected_mwfn
            )
            status.config(text="Done.")
            if selected_vis != "y":
                messagebox.showinfo("Finished", "Processing completed successfully.")
        except Exception as e:
            status.config(text="Failed.")
            messagebox.showerror("Execution error", str(e))

    tk.Button(frm, text="Run VisMap", command=run_clicked, width=18, height=2).grid(row=6, column=0, sticky="w", pady=(14, 0))
    tk.Button(frm, text="Quit", command=root.destroy, width=12).grid(row=6, column=1, sticky="e", pady=(14, 0))

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