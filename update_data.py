#!/usr/bin/env python3
"""
update_data.py  —  Porra Mundial 2026
Uso: python update_data.py <archivo.xlsx>
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
GROUP_ORDER = list("ABCDEFGHIJKL")
GM = [(0,1),(0,2),(1,2),(0,3),(1,3),(2,3)]
PLAYERS = ["Tamayo","Vispo","Juanpa","Macaco","Aranguren","Arocena","Alonso","Quemada"]

# Chronological order of the 72 group stage matches (real FIFA calendar order)
CHRONO_ORDER = [
    2,5,10,19,9,13,14,20,25,31,37,43,44,49,54,55,60,67,72,
    1,6,8,11,15,16,21,22,26,27,28,29,32,33,34,38,40,41,45,47,
    50,53,56,59,61,66,68,69,
    3,4,7,12,17,18,23,24,30,35,36,39,42,46,48,
    51,52,57,58,62,63,64,65,70,71
]

# Match labels for tooltip
def match_label(mid, official):
    o = official.get(str(mid), [None,None,None,None])
    t1 = o[0] or "?"
    t2 = o[3] or "?"
    g1 = o[1]
    g2 = o[2]
    if g1 is not None:
        return f"P{mid}: {t1} {g1}-{g2} {t2}"
    return f"P{mid}: {t1} vs {t2}"

def safe_int(v):
    if v is None: return None
    try: return int(float(str(v)))
    except: return None

def compute_standings(group, preds_g):
    teams = GROUPS[group]
    stats = {t:{"pts":0,"gf":0,"gc":0,"pj":0} for t in teams}
    g_idx = GROUP_ORDER.index(group)
    for mi,(i,j) in enumerate(GM):
        mid = str(g_idx*6+mi+1)
        t1,t2 = teams[i],teams[j]
        pr = preds_g.get(mid,[None,None])
        g1,g2 = pr[0],pr[1]
        if g1 is None or g2 is None: continue
        stats[t1]["gf"]+=g1;stats[t1]["gc"]+=g2;stats[t1]["pj"]+=1
        stats[t2]["gf"]+=g2;stats[t2]["gc"]+=g1;stats[t2]["pj"]+=1
        if g1>g2: stats[t1]["pts"]+=3
        elif g1<g2: stats[t2]["pts"]+=3
        else: stats[t1]["pts"]+=1;stats[t2]["pts"]+=1
    ranked=sorted(teams,key=lambda t:(stats[t]["pts"],stats[t]["gf"]-stats[t]["gc"],stats[t]["gf"]),reverse=True)
    return ranked, stats

def calc_match_pts(pg1,pg2,og1,og2):
    if any(v is None for v in [pg1,pg2,og1,og2]): return 0
    pts=0
    if pg1==og1 and pg2==og2: pts=6
    elif (pg1>pg2)==(og1>og2) and (pg1==pg2)==(og1==og2): pts=3
    if pg1==og1 and pts<6: pts+=1
    if pg2==og2 and pts<6: pts+=1
    return pts

def run(excel_path, out_path="data.json"):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws_res = wb["🏟️ Resultados Oficiales"]

    official = {}
    for r in range(5,102):
        a = ws_res.cell(row=r,column=1).value
        if isinstance(a,int) and 1<=a<=72:
            t1=ws_res.cell(row=r,column=2).value
            g1=safe_int(ws_res.cell(row=r,column=3).value)
            g2=safe_int(ws_res.cell(row=r,column=5).value)
            t2=ws_res.cell(row=r,column=6).value
            official[str(a)]=[t1,g1,g2,t2]

    # Scored matches in CHRONOLOGICAL order
    scored_chrono = [m for m in CHRONO_ORDER if official.get(str(m),[None,None,None,None])[1] is not None]
    print(f"  Resultados oficiales: {len(scored_chrono)}/72")

    # Match labels for chart tooltips
    match_labels = {str(m): match_label(m, official) for m in CHRONO_ORDER}

    # Official team pts
    official_team_pts = {}
    for teams in GROUPS.values():
        for team in teams:
            pts=0
            for val in official.values():
                t1,g1,g2,t2=val
                if g1 is None: continue
                if t1==team: pts+=3 if g1>g2 else(1 if g1==g2 else 0)
                elif t2==team: pts+=3 if g2>g1 else(1 if g1==g2 else 0)
            official_team_pts[team]=pts

    # Official group rankings
    official_rankings = {}
    for g in GROUP_ORDER:
        g_idx=GROUP_ORDER.index(g)
        off_preds={}
        for mi in range(6):
            mid=str(g_idx*6+mi+1)
            if mid in official and official[mid][1] is not None:
                off_preds[mid]=[official[mid][1],official[mid][2]]
        ranked,_=compute_standings(g,off_preds)
        official_rankings[g]=ranked

    players={}
    for p in PLAYERS:
        ws=wb[p] if p in wb.sheetnames else None
        if not ws: print(f"  ⚠️  '{p}' no encontrada"); continue
        preds={}
        for r in range(5,102):
            a=ws.cell(row=r,column=1).value
            if isinstance(a,int) and 1<=a<=72:
                g1=safe_int(ws.cell(row=r,column=3).value)
                g2=safe_int(ws.cell(row=r,column=5).value)
                preds[str(a)]=[g1,g2]
        total=int(ws["E459"].value or 0) if ws["E459"] else 0

        # Match history in CHRONOLOGICAL order
        match_history=[]
        cumulative=0
        for mid_int in scored_chrono:
            mid=str(mid_int)
            off=official.get(mid,[None,None,None,None])
            pred=preds.get(mid,[None,None])
            pts=calc_match_pts(pred[0],pred[1],off[1],off[2])
            cumulative+=pts
            match_history.append({"mid":mid_int,"pts":pts,"cum":cumulative,
                                   "lbl":match_label(mid_int,official)})
        # Add group classification points as final step so graph reaches Excel total
        # This includes position pts + bonus + team exact pts
        excel_total=int(ws["E459"].value or 0) if ws["E459"] else 0
        if match_history:
            group_extra=excel_total-cumulative
            if group_extra>0:
                match_history.append({"mid":0,"pts":group_extra,"cum":excel_total,
                                       "lbl":f"Clasificación de grupos (+{group_extra}pts)"})

        # Group breakdown
        groups_breakdown={}
        for g in GROUP_ORDER:
            g_idx=GROUP_ORDER.index(g)
            teams=GROUPS[g]
            p_preds_g={str(g_idx*6+mi+1):preds.get(str(g_idx*6+mi+1),[None,None]) for mi in range(6)}
            p_ranked,p_stats=compute_standings(g,p_preds_g)
            o_ranked=official_rankings[g]
            pos_pts=0
            positions_detail=[]
            for pos_i in range(4):
                pt=p_ranked[pos_i] if pos_i<len(p_ranked) else None
                ot=o_ranked[pos_i] if pos_i<len(o_ranked) else None
                correct=(pt==ot and pt is not None and ot is not None)
                if correct: pos_pts+=3
                positions_detail.append({"pos":pos_i+1,"predicted":pt,"official":ot,"correct":correct})
            bonus=3 if pos_pts==12 else 0
            team_pts=0
            team_pts_detail=[]
            for team in teams:
                ptp=p_stats.get(team,{}).get("pts",0)
                otp=official_team_pts.get(team,0)
                match=(ptp==otp)
                if match: team_pts+=2
                team_pts_detail.append({"team":team,"predicted_pts":ptp,"official_pts":otp,"correct":match,"earned":2 if match else 0})
            groups_breakdown[g]={
                "pos_pts":pos_pts,"bonus":bonus,"team_pts":team_pts,
                "total_group_extra":pos_pts+bonus+team_pts,
                "positions":positions_detail,"teams":team_pts_detail
            }

        players[p]={"total":total,"predictions":preds,"match_history":match_history,"groups_breakdown":groups_breakdown}
        extra=sum(v["total_group_extra"] for v in groups_breakdown.values())
        print(f"  {p}: {total} pts")

    data={"groups":GROUPS,"group_order":GROUP_ORDER,"official":official,
          "official_team_pts":official_team_pts,"official_rankings":official_rankings,
          "scored_chrono":scored_chrono,"match_labels":match_labels,"players":players}
    with open(out_path,"w",encoding="utf-8") as f:
        json.dump(data,f,ensure_ascii=False,separators=(",",":"))
    print(f"\n✅ data.json — {len(scored_chrono)}/72 partidos (orden cronológico)")

if __name__=="__main__":
    if len(sys.argv)<2: print("Uso: python update_data.py <archivo.xlsx>"); sys.exit(1)
    run(sys.argv[1],"data.json")

# NOTE: patch applied - add group classification bonus to match_history
# (This block is handled inline in the run() function below, 
#  see the match_history.append at the end)
