#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generuje .xlsx s českými vzorcami:
- typ 'minimal' (A..G) alebo 'full' (rozšírené + súhrn),
- vždy list 'Kružnice' (cos/sin v stupňoch),
- voliteľne list 'Leibniz'.

Použitie:
  python generate_pi_excel.py --out Pi_aproximace_CZ_minimal.xlsx --rows 3000 --type minimal
  python generate_pi_excel.py --out Pi_aproximace_CZ_vzorce.xlsx  --rows 3000 --type full --with-leibniz
"""

import argparse
import xlsxwriter

def build_circle_sheet(wb):
    ws = wb.add_worksheet("Kružnice")
    fmt_h = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    ws.write(0,0,"Stupeň", fmt_h)
    ws.write(0,1,"x=cos", fmt_h)
    ws.write(0,2,"y=sin", fmt_h)
    for deg in range(0, 91):  # 0..90°
        r = deg + 1
        ws.write(r, 0, deg)
        ws.write_formula(r, 1, f"=COS(RADIÁNY(A{r+1}))")
        ws.write_formula(r, 2, f"=SIN(RADIÁNY(A{r+1}))")
    ws.set_column("A:C", 14)

def build_mc_minimal(wb, rows: int):
    fmt_h = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    ws = wb.add_worksheet("MonteCarlo")
    headers = ["x = NÁHČÍSLO()", "y = NÁHČÍSLO()", "Uvnitř (0/1)",
               "X_uvnitř", "Y_uvnitř", "X_vně", "Y_vně"]
    for c,h in enumerate(headers):
        ws.write(0, c, h, fmt_h)

    for i in range(1, rows+1):
        r = i
        excel_row = r + 1
        ws.write_formula(r, 0, "=NÁHČÍSLO()")
        ws.write_formula(r, 1, "=NÁHČÍSLO()")
        ws.write_formula(r, 2, f"=KDYŽ(A{excel_row+1}^2+B{excel_row+1}^2<=1;1;0)")
        ws.write_formula(r, 3, f"=KDYŽ(C{excel_row+1}=1;A{excel_row+1};CHYBA.NA())")
        ws.write_formula(r, 4, f"=KDYŽ(C{excel_row+1}=1;B{excel_row+1};CHYBA.NA())")
        ws.write_formula(r, 5, f"=KDYŽ(C{excel_row+1}=0;A{excel_row+1};CHYBA.NA())")
        ws.write_formula(r, 6, f"=KDYŽ(C{excel_row+1}=0;B{excel_row+1};CHYBA.NA())")

    ws.set_column("A:G", 16)

    chart = wb.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
    chart.add_series({
        'name': 'Uvnitř',
        'categories': ['MonteCarlo', 1, 3, rows, 3],  # D
        'values':     ['MonteCarlo', 1, 4, rows, 4],  # E
        'marker': {'type': 'circle', 'size': 3},
        'line': {'none': True},
    })
    chart.add_series({
        'name': 'Vně',
        'categories': ['MonteCarlo', 1, 5, rows, 5],  # F
        'values':     ['MonteCarlo', 1, 6, rows, 6],  # G
        'marker': {'type': 'circle', 'size': 3},
        'line': {'none': True},
    })
    chart.add_series({
        'name': 'Čtvrtkružnice',
        'categories': ['Kružnice', 1, 1, 91, 1],
        'values':     ['Kružnice', 1, 2, 91, 2],
        'marker': {'type': 'none'},
        'line': {'width': 1.5},
    })
    chart.set_title({'name':'Monte Carlo – aproximace π'})
    chart.set_x_axis({'name':'x','min':-0.05,'max':1.05,'major_gridlines': {'visible': False}})
    chart.set_y_axis({'name':'y','min':-0.05,'max':1.05,'major_gridlines': {'visible': False}})
    chart.set_legend({'position':'bottom'})
    ws.insert_chart("I2", chart, {'x_scale':1.2, 'y_scale':1.2})

def build_mc_full(wb, rows: int):
    fmt_h = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    fmt_pct = wb.add_format({"num_format": "0.00%"})
    ws = wb.add_worksheet("MonteCarlo")
    headers = [
        "x = NÁHČÍSLO()", "y = NÁHČÍSLO()", "r² = x^2 + y^2",
        "Uvnitř (≤1)", "Kumulativně uvnitř", "n (počet bodů)",
        "π̂ = 4 * uvnitř / n", "X_uvnitř", "Y_uvnitř", "X_vně", "Y_vně"
    ]
    for c,h in enumerate(headers):
        ws.write(0, c, h, fmt_h)

    for i in range(1, rows+1):
        r = i
        excel_row = r + 1
        ws.write_formula(r, 0, "=NÁHČÍSLO()")
        ws.write_formula(r, 1, "=NÁHČÍSLO()")
        ws.write_formula(r, 2, f"=A{excel_row+1}^2+B{excel_row+1}^2")
        ws.write_formula(r, 3, f"=KDYŽ(C{excel_row+1}<=1;1;0)")
        ws.write_formula(r, 4, f"=SUMA($D$2:D{excel_row+1})")
        ws.write_formula(r, 5, f"=ŘÁDEK()-1")
        ws.write_formula(r, 6, f"=4*E{excel_row+1}/F{excel_row+1}")
        ws.write_formula(r, 7, f"=KDYŽ(D{excel_row+1}=1;A{excel_row+1};CHYBA.NA())")
        ws.write_formula(r, 8, f"=KDYŽ(D{excel_row+1}=1;B{excel_row+1};CHYBA.NA())")
        ws.write_formula(r, 9, f"=KDYŽ(D{excel_row+1}=0;A{excel_row+1};CHYBA.NA())")
        ws.write_formula(r,10, f"=KDYŽ(D{excel_row+1}=0;B{excel_row+1};CHYBA.NA())")

    ws.set_column("A:K", 16)
    ws.set_column("M:N", 24)

    ws.write("M2", "Aktuální odhad π:", fmt_h)
    ws.write_formula("N2", "=INDEX(G:G;POČET(G:G))")
    ws.write("M3", "Bodů celkem:", fmt_h)
    ws.write_formula("N3", "=POČET(A:A)")
    ws.write("M4", "Podíl uvnitř kruhu:", fmt_h)
    ws.write_formula("N4", "=INDEX(E:E;POČET(E:E))/INDEX(F:F;POČET(F:F))", fmt_pct)

    chart = wb.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
    chart.add_series({
        'name': 'Uvnitř kruhu',
        'categories': ['MonteCarlo', 1, 7, rows, 7],
        'values':     ['MonteCarlo', 1, 8, rows, 8],
        'marker': {'type':'circle','size':3},
        'line': {'none': True},
    })
    chart.add_series({
        'name': 'Vně kruhu',
        'categories': ['MonteCarlo', 1, 9, rows, 9],
        'values':     ['MonteCarlo', 1, 10, rows, 10],
        'marker': {'type':'circle','size':3},
        'line': {'none': True},
    })
    chart.add_series({
        'name': 'Jednotková čtvrtkružnice',
        'categories': ['Kružnice', 1, 1, 91, 1],
        'values':     ['Kružnice', 1, 2, 91, 2],
        'marker': {'type':'none'},
        'line': {'width': 1.5},
    })
    chart.set_title({'name':'Monte Carlo – aproximace π'})
    chart.set_x_axis({'name':'x','min':-0.05,'max':1.05,'major_gridlines': {'visible': False}})
    chart.set_y_axis({'name':'y','min':-0.05,'max':1.05,'major_gridlines': {'visible': False}})
    chart.set_legend({'position':'bottom'})
    ws.insert_chart("M6", chart, {'x_scale':1.25, 'y_scale':1.25})

def build_leibniz(wb, terms: int):
    fmt_h = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    fmt_num = wb.add_format({"num_format": "0.000000"})
    wl = wb.add_worksheet("Leibniz")
    wl.write(0,0,"k", fmt_h)
    wl.write(0,1,"Člen a_k = (-1)^k/(2k+1)", fmt_h)
    wl.write(0,2,"Suma S_n = Σ a_k", fmt_h)
    wl.write(0,3,"π̂_n = 4*S_n", fmt_h)

    for k in range(terms):
        r = k + 1
        wl.write(r, 0, k)
        wl.write_formula(r, 1, f"=(-1)^A{r+1}/(2*A{r+1}+1)")
        if r == 1:
            wl.write_formula(r, 2, f"=B{r+1}")
        else:
            wl.write_formula(r, 2, f"=C{r}+B{r+1}")
        wl.write_formula(r, 3, f"=4*C{r+1}")

    chart = wb.add_chart({'type': 'line'})
    chart.add_series({
        'name': 'π̂_n (Leibniz)',
        'categories': ['Leibniz', 1, 0, terms, 0],
        'values':     ['Leibniz', 1, 3, terms, 3],
    })
    chart.set_title({'name': 'Leibniz – konvergence k π'})
    chart.set_x_axis({'name':'k'})
    chart.set_y_axis({'name':'Odhad π'})
    chart.set_legend({'none': True})
    wl.insert_chart("F2", chart, {'x_scale':1.45,'y_scale':1.2})
    wl.write("F20", "Aktuální odhad π (poslední řádek):", fmt_h)
    wl.write_formula("G20", "=INDEX(D:D;POČET(D:D))", fmt_num)
    wl.set_column("A:D", 24)
    wl.set_column("F:G", 28)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="Výstupní .xlsx soubor")
    ap.add_argument("--rows", type=int, default=3000, help="Počet náhodných bodů Monte Carlo (default 3000)")
    ap.add_argument("--type", choices=["minimal","full"], default="minimal",
                    help="Typ Excelu: minimal (A..G) nebo full (s běžícím odhadem π)")
    ap.add_argument("--with-leibniz", action="store_true", help="Přidat list Leibniz")
    args = ap.parse_args()

    wb = xlsxwriter.Workbook(args.out)
    build_circle_sheet(wb)
    if args.type == "minimal":
        build_mc_minimal(wb, args.rows)
    else:
        build_mc_full(wb, args.rows)
    if args.with_leibniz:
        build_leibniz(wb, terms=max(2000, args.rows))
    wb.close()
    print(f"Hotovo → {args.out}")

if __name__ == "__main__":
    main()
