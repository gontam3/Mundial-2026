#!/usr/bin/env python3
"""
update_data.py — Porra Mundial 2026 (SOLO FASE ELIMINATORIA)
Uso: python update_data.py <Porra_KO_2026.xlsx>
"""
import sys, json, openpyxl

PLAYERS = ["Tamayo","Vispo","Juanpa","Macaco","Aranguren","Arocena","Alonso","Quemada"]
GROUP_TOTALS = {"Tamayo":349,"Vispo":308,"Juanpa":311,"Macaco":332,
                "Aranguren":332,"Arocena":357,"Alonso":311,"Quemada":346}

# Official sheet (col D = winner)
OFF_R16 = list(range(5,21))    # rows 5-20
OFF_R8  = list(range(24,32))   # rows 24-31
OFF_R4  = list(range(35,39))   # rows 35-38
OFF_R2  = list(range(42,44))   # rows 42-43
OFF_R3  = 47
OFF_FIN = 51

# Player sheets (col F = prediction)
PLR_R16 = list(range(6,22))    # rows 6-21
PLR_R8  = list(range(25,33))   # rows 25-32
PLR_R4  = list(range(35,39))   # rows 35-38
PLR_R2  = list(range(41,43))   # rows 41-42
PLR_R3  = 45
PLR_FIN = 49

def nv(v): return v if v not in (None, "") else None

def ko_pts(preds, offs, each):
    result = []
    for pred, off in zip(preds or [], offs or []):
        pts = each if (pred and off and pred == off) else 0
        result.append({"pred": pred, "official": off, "pts": pts})
    return result

def run(ko_path, out_path="data.json"):
    try:
        with open(out_path, encoding="utf-8") as f:
            data = json.load(f)
        print("  Datos de grupos cargados del data.json existente")
    except FileNotFoundError:
        print("  ERROR: data.json no encontrado"); sys.exit(1)

    wb = openpyxl.load_workbook(ko_path, data_only=True)
    ws_off = wb["🏆 Resultados Oficiales"]

    r16_off = [nv(ws_off.cell(row=r, column=4).value) for r in OFF_R16]
    r8_off  = [nv(ws_off.cell(row=r, column=4).value) for r in OFF_R8]
    r4_off  = [nv(ws_off.cell(row=r, column=4).value) for r in OFF_R4]
    r2_off  = [nv(ws_off.cell(row=r, column=4).value) for r in OFF_R2]
    r3_off  = nv(ws_off.cell(row=OFF_R3,  column=4).value)
    fin_off = nv(ws_off.cell(row=OFF_FIN, column=4).value)

    official_ko = {"r16":r16_off,"r8":r8_off,"r4":r4_off,
                   "r2":r2_off,"r3":r3_off,"final":fin_off}

    print(f"  R16: {[x for x in r16_off if x]}")
    print(f"  R8:  {[x for x in r8_off  if x]}")
    print(f"  R4:  {[x for x in r4_off  if x]}")
    print(f"  R2:  {[x for x in r2_off  if x]}")
    if r3_off:  print(f"  3°:  {r3_off}")
    if fin_off: print(f"  Final: {fin_off}")

    for p in PLAYERS:
        ws = wb[p] if p in wb.sheetnames else None
        if not ws:
            print(f"  ⚠️  '{p}' no encontrada"); continue

        r16_p = [nv(ws.cell(row=r, column=6).value) for r in PLR_R16]
        r8_p  = [nv(ws.cell(row=r, column=6).value) for r in PLR_R8]
        r4_p  = [nv(ws.cell(row=r, column=6).value) for r in PLR_R4]
        r2_p  = [nv(ws.cell(row=r, column=6).value) for r in PLR_R2]
        r3_p  = nv(ws.cell(row=PLR_R3,  column=6).value)
        fin_p = nv(ws.cell(row=PLR_FIN, column=6).value)

        r16d = ko_pts(r16_p, r16_off, 4)
        r8d  = ko_pts(r8_p,  r8_off,  8)
        r4d  = ko_pts(r4_p,  r4_off,  12)
        r2d  = ko_pts(r2_p,  r2_off,  16)
        r3pts  = 8  if (r3_p  and r3_off  and r3_p  == r3_off)  else 0
        finpts = 30 if (fin_p and fin_off and fin_p == fin_off)  else 0

        ko_total = (sum(x["pts"] for x in r16d) + sum(x["pts"] for x in r8d) +
                    sum(x["pts"] for x in r4d)  + sum(x["pts"] for x in r2d) +
                    r3pts + finpts)
        group_total = GROUP_TOTALS.get(p, data["players"].get(p, {}).get("group_total", 0))

        if p not in data["players"]: data["players"][p] = {}
        data["players"][p].update({
            "group_total": group_total,
            "ko_total":    ko_total,
            "total":       group_total + ko_total,
            "ko_preds":  {"r16":r16_p,"r8":r8_p,"r4":r4_p,"r2":r2_p,"r3":r3_p,"final":fin_p},
            "ko_detail": {"r16":r16d,"r8":r8d,"r4":r4d,"r2":r2d,
                          "r3":    {"pred":r3_p,  "official":r3_off,  "pts":r3pts},
                          "final": {"pred":fin_p, "official":fin_off, "pts":finpts}},
        })
        print(f"  {p}: grp={group_total} + KO={ko_total} = {group_total+ko_total}")

    data["official_ko"] = official_ko
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",",":"))
    print(f"\n✅ data.json actualizado")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python update_data.py <Porra_KO_2026.xlsx>"); sys.exit(1)
    run(sys.argv[1], "data.json")
