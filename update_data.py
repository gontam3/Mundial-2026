#!/usr/bin/env python3
"""
update_data.py
Uso: python update_data.py Porra_Mundial_V3.xlsx
Extrae los datos del Excel y actualiza data.json en la misma carpeta.
"""
import sys, json, openpyxl

GROUPS = {
    "A":["México","Corea del Sur","Sudáfrica","República Checa"],
    "B":["Canadá","Suiza","Catar","Bosnia-Herzegovina"],
    "C":["Brasil","Marruecos","Escocia","Haití"],
    "D":["USA","Paraguay","Australia","Turquía"],
    "E":["Alemania","Curaçao","Costa de Marfil","Ecuador"],
    "F":["Países Bajos","Japón","Suecia","Túnez"],
    "G":["Bélgica","Egipto","Irán","Nueva Zelanda"],
    "H":["España","Cabo Verde","Arabia Saudí","Uruguay"],
    "I":["Francia","Senegal","Irak","Noruega"],
    "J":["Argentina","Argelia","Austria","Jordania"],
    "K":["Portugal","R.D. Congo","Uzbekistán","Colombia"],
    "L":["Inglaterra","Croacia","Ghana","Panamá"],
}
PLAYERS = ["Tamayo","Vispo","Juanpa","Macaco","Aranguren","Arocena","Alonso","Quemada"]

def run(excel_path, out_path="data.json"):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws_res = wb["🏟️ Resultados Oficiales"]

    official = {}
    for r in range(5, 102):
        a = ws_res.cell(row=r, column=1).value
        if isinstance(a, int) and 1 <= a <= 72:
            t1 = ws_res.cell(row=r, column=2).value
            g1 = ws_res.cell(row=r, column=3).value
            g2 = ws_res.cell(row=r, column=5).value
            t2 = ws_res.cell(row=r, column=6).value
            official[str(a)] = [t1, None if g1 is None else int(g1),
                                     None if g2 is None else int(g2), t2]

    players = {}
    for p in PLAYERS:
        ws = wb[p] if p in wb.sheetnames else None
        if not ws:
            print(f"  ⚠️  Hoja '{p}' no encontrada, omitiendo")
            continue
        preds = {}
        for r in range(5, 102):
            a = ws.cell(row=r, column=1).value
            if isinstance(a, int) and 1 <= a <= 72:
                g1 = ws.cell(row=r, column=3).value
                g2 = ws.cell(row=r, column=5).value
                preds[str(a)] = [None if g1 is None else int(g1),
                                  None if g2 is None else int(g2)]
        total_cell = ws["E459"]
        total = int(total_cell.value or 0) if total_cell else 0
        players[p] = {"total": total, "predictions": preds}
        print(f"  {p}: {total} pts, {sum(1 for v in preds.values() if v[0] is not None)}/72 partidos")

    data = {
        "groups": GROUPS,
        "group_order": list("ABCDEFGHIJKL"),
        "official": official,
        "players": players,
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

    scored = sum(1 for v in official.values() if v[1] is not None)
    print(f"\n✅ data.json actualizado — {scored}/72 resultados oficiales")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python update_data.py <archivo.xlsx>")
        sys.exit(1)
    run(sys.argv[1], "data.json")
