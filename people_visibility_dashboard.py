#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
People Visibility Dashboard generator
=====================================
Reads "People Visiblity new format.xlsx" (sheets: Data, Master) and writes a
self-contained, interactive HTML dashboard (Vianaar theme) next to it.

USAGE
-----
    python people_visibility_dashboard.py                # uses INPUT_PATH below
    python people_visibility_dashboard.py "C:\\path\\to\\file.xlsx"
    python people_visibility_dashboard.py in.xlsx out.html

Requires: pandas, openpyxl   ->  pip install pandas openpyxl

NOTE: If the workbook is open in Excel you may hit "PermissionError [Errno 13]".
The script automatically falls back to reading a temporary copy, but closing the
file in Excel first is the cleanest fix.

METRIC / TABLE DEFINITIONS (locked with stakeholder)
----------------------------------------------------
* Project Count          = number of projects in the Data sheet.
* Team Count             = ALL people rows in the Master sheet.
* Unallocated Resources  = Master people whose name is NOT found among the
                           deployed names (People1..N + Project Incharge) in Data.
* Resource Utilization % = allocated Master people / Team Count * 100.
* Allocation %  (table)  = 1 / (number of projects the person is staffed on),
                           e.g. 4 projects -> 25%.  One project -> 100%.
                           A person's project list includes projects they lead
                           as Project Incharge.
* Project Incharge       = "Project Incharge" column (Data sheet); counted as a
                           team member of that project and available as a left
                           filter (NOT shown as a table column).
* Baseline Finish        = "Baseline Finished" column (Data sheet).
* Forecasted End         = "Forecasted End Date" column (Data sheet).
* Over Due Days          = Forecasted End - Baseline Finish, in days.
                           Positive = late (red v). Negative = ahead (green ^).
  (NOTE: "Baseline Start" and "Actual Start" are empty in the source file,
   so they are not shown. Fill them in Excel and they can be added.)
"""

import os
import re
import sys
import json
import html
import shutil
import tempfile
from datetime import datetime

import pandas as pd

# --------------------------------------------------------------------------- #
# CONFIG - edit INPUT_PATH to point at the file on your OneDrive if needed
# --------------------------------------------------------------------------- #
INPUT_PATH = r"C:\Users\jitender.chaurasia\OneDrive - Vianaar Homes Pvt Ltd\Power BI Dashboard - Documents\10. People Visibility\People Visiblity new format.xlsx"
OUTPUT_NAME = "People_Visibility_Dashboard.html"

PEOPLE_COLS = [f"People{i}" for i in range(1, 11)]
INCHARGE_COL = "Project Incharge"
AS_OF = datetime.today()


# --------------------------------------------------------------------------- #
# Robust workbook loader (handles file open in Excel / OneDrive lock)
# --------------------------------------------------------------------------- #
def load_excel_file(path):
    try:
        return pd.ExcelFile(path, engine="openpyxl")
    except (PermissionError, OSError):
        tmp = os.path.join(tempfile.gettempdir(), f"_pv_dashboard_{os.getpid()}.xlsx")
        try:
            shutil.copy2(path, tmp)
        except (PermissionError, OSError) as e:
            sys.exit(
                "ERROR: Could not read the workbook.\n"
                f"  {path}\n"
                "It looks locked. Please CLOSE the file in Excel (and make sure "
                "OneDrive has it available offline), then run again.\n"
                f"Details: {e}")
        return pd.ExcelFile(tmp, engine="openpyxl")


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def norm(name):
    if name is None:
        return ""
    s = str(name).strip().upper()
    if s in ("", "NAN", "NONE"):
        return ""
    s = re.sub(r"[.\-_]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tenure_bucket(joining):
    if joining is None or pd.isna(joining):
        return "Unknown"
    try:
        jd = pd.to_datetime(joining)
    except Exception:
        return "Unknown"
    yrs = (AS_OF - jd).days / 365.25
    if yrs < 1:
        return "Below 1 yr"
    if yrs <= 3:
        return "1 - 3 yrs"
    if yrs <= 7:
        return "3.1 - 7 yrs"
    if yrs <= 10:
        return "7.1 - 10 yrs"
    if yrs <= 15:
        return "10.1 - 15 yrs"
    return "15+ yrs"


def to_dt(v):
    if v is None or pd.isna(v):
        return None
    try:
        return pd.to_datetime(v)
    except Exception:
        return None


def fmt_date(dt):
    return "" if dt is None else dt.strftime("%d-%b-%Y")


def cell_str(v):
    return "" if (v is None or pd.isna(v) or not str(v).strip()) else str(v).strip()


# --------------------------------------------------------------------------- #
# Load + transform
# --------------------------------------------------------------------------- #
def build_model(xlsx_path):
    xl = load_excel_file(xlsx_path)
    data = pd.read_excel(xl, sheet_name="Data", header=0)
    master = pd.read_excel(xl, sheet_name="Master", header=0)

    data = data[data["Project Name"].notna()].copy()
    has_incharge = INCHARGE_COL in data.columns

    # ---- Projects, dates, deployed names -------------------------------- #
    projects = []
    project_meta = {}                # project name -> {bf, fe, overdue, incharge}
    deployed_set = set()             # normalised deployed names (People + Incharge)
    person_to_projects = {}          # normalised deployed name -> [project names]

    def attach(n, pname):
        if not n:
            return
        deployed_set.add(n)
        person_to_projects.setdefault(n, [])
        if pname not in person_to_projects[n]:
            person_to_projects[n].append(pname)

    for _, r in data.iterrows():
        pname = str(r["Project Name"]).strip()
        bf = to_dt(r.get("Baseline Finished"))
        fe = to_dt(r.get("Forecasted End Date"))
        overdue = (fe - bf).days if (bf is not None and fe is not None) else None
        incharge = cell_str(r.get(INCHARGE_COL)) if has_incharge else ""
        project_meta[pname] = {
            "baseline_finish": fmt_date(bf),
            "forecast_end": fmt_date(fe),
            "overdue": overdue,
            "incharge": incharge,
        }

        cells = []
        for c in PEOPLE_COLS:
            v = cell_str(r.get(c)) if c in data.columns else ""
            if v:
                cells.append(v)
                attach(norm(v), pname)
        # Project Incharge counts as a team member of the project too
        if incharge:
            attach(norm(incharge), pname)

        projects.append({
            "name": pname,
            "units": None if pd.isna(r.get("Units")) else int(r.get("Units")),
            "total_people": 0 if pd.isna(r.get("Total People")) else int(r.get("Total People")),
            "deployed_cells": len(cells),
            "incharge": incharge,
            "baseline_finish": project_meta[pname]["baseline_finish"],
            "forecast_end": project_meta[pname]["forecast_end"],
            "overdue": overdue,
        })

    # ---- People (Master) + per-project assignments ---------------------- #
    people = []
    for _, r in master.iterrows():
        nm = cell_str(r.get("EMP NAME"))
        if not nm:
            continue
        nnm = norm(nm)
        proj_list = person_to_projects.get(nnm, [])
        allocated = len(proj_list) > 0
        n_proj = len(proj_list)
        alloc_pct = round(100 / n_proj) if n_proj else None
        assignments = []
        for p in proj_list:
            meta = project_meta.get(p, {})
            assignments.append({
                "project": p,
                "incharge": meta.get("incharge", ""),
                "is_incharge": norm(meta.get("incharge", "")) == nnm,
                "allocation": alloc_pct,
                "baseline_finish": meta.get("baseline_finish", ""),
                "forecast_end": meta.get("forecast_end", ""),
                "overdue": meta.get("overdue", None),
            })

        people.append({
            "emp_code": cell_str(r.get("EMP CODE")),
            "name": nm,
            "location": cell_str(r.get("Location Name")),
            "designation": re.sub(r"\s+", " ", cell_str(r.get("DESIGNATION"))),
            "sub_dept": cell_str(r.get("Sub - Department")),
            "status": cell_str(r.get("In System / Left")),
            "tenure": tenure_bucket(r.get("Joining Dates")),
            "allocated": allocated,
            "allocation": alloc_pct,
            "projects": proj_list,
            "assignments": assignments,
        })

    team_count = len(people)
    allocated_count = sum(1 for p in people if p["allocated"])
    unallocated_count = team_count - allocated_count
    utilization = round(allocated_count / team_count * 100) if team_count else 0

    summary = {
        "project_count": len(projects),
        "team_count": team_count,
        "deployed_cells": sum(p["deployed_cells"] for p in projects),
        "allocated_count": allocated_count,
        "unallocated_count": unallocated_count,
        "utilization": utilization,
        "as_of": AS_OF.strftime("%d-%b-%Y"),
    }
    return {"summary": summary, "projects": projects, "people": people}


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #
def render_html(model, source_name):
    payload = json.dumps(model, ensure_ascii=False)
    src = html.escape(source_name)
    return HTML_TEMPLATE.replace("/*__DATA__*/", payload).replace("__SOURCE__", src)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>People Visibility - Project Wise</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"></script>
<style>
  :root{
    --green:#2e7d4f; --green-dark:#1f5d3a; --green-light:#e7f1ea; --green-mid:#cfe3d6;
    --bar:#8e8ec9; --ink:#1f2d27; --muted:#6b7d73; --line:#d7e4dc;
    --red:#c0392b; --amber:#e0a800;
  }
  *{box-sizing:border-box}
  body{margin:0;background:#f3f7f4;color:var(--ink);
       font-family:"Segoe UI",Roboto,Helvetica,Arial,sans-serif;font-size:13px}
  .wrap{max-width:1500px;margin:0 auto;padding:14px}
  .topbar{display:flex;align-items:center;justify-content:space-between;
          background:var(--green-light);border:2px solid var(--green);
          border-radius:6px;padding:10px 18px;margin-bottom:12px}
  .topbar h1{font-size:20px;color:var(--green-dark);margin:0;font-weight:700}
  .topbar .period{color:var(--green-dark);font-weight:600}
  .brand{font-weight:800;color:var(--green-dark);letter-spacing:1px;font-size:18px}
  .layout{display:grid;grid-template-columns:230px 1fr;gap:12px}
  .panel{background:var(--green-light);border:1px solid var(--green-mid);
         border-radius:6px;padding:12px}
  .panel h3{margin:2px 0 6px;color:var(--green-dark);font-size:12px;text-transform:uppercase;letter-spacing:.5px}
  select,.fbtn{width:100%;padding:7px 9px;border:1px solid var(--line);border-radius:4px;
          background:#fff;font-size:13px;margin-bottom:12px;color:var(--ink)}
  .loc-btns{display:flex;flex-direction:column;gap:8px;margin-bottom:6px}
  .fbtn{cursor:pointer;text-align:center;font-weight:600;color:var(--green-dark);background:#fff;
        border:1px solid var(--green-mid);transition:.15s}
  .fbtn.active{background:var(--green);color:#fff;border-color:var(--green)}
  .reset{cursor:pointer;background:#fff;border:1px solid var(--line);color:var(--muted)}
  .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:12px}
  .kpi{background:var(--green-light);border:1px solid var(--green-mid);border-radius:6px;
       padding:14px;text-align:center}
  .kpi .label{color:var(--green-dark);font-weight:600;font-size:12px}
  .kpi .val{font-size:34px;font-weight:800;color:var(--ink);margin-top:4px}
  .kpi .sub{font-size:11px;color:var(--muted);margin-top:2px}
  .grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:12px}
  .grid2{display:grid;grid-template-columns:360px 1fr;gap:12px}
  .chartbox{background:#fff;border:1px solid var(--green-mid);border-radius:6px;padding:10px}
  .chartbox h3{margin:0 0 8px;color:var(--green-dark);font-size:13px;text-align:center}
  .chartwrap{position:relative;height:230px}
  .tablewrap{background:#fff;border:1px solid var(--green-mid);border-radius:6px;overflow:hidden}
  .tablewrap h3{margin:0;background:var(--green);color:#fff;padding:9px 12px;font-size:14px}
  table{width:100%;border-collapse:collapse}
  th,td{padding:6px 9px;border-bottom:1px solid var(--line);text-align:left;font-size:12px;vertical-align:top}
  th{background:var(--green-dark);color:#fff;position:sticky;top:0;white-space:nowrap}
  td.num,th.num{text-align:right}
  tbody tr.grp-start td{border-top:2px solid var(--green-mid)}
  .name-cell{font-weight:700;color:var(--green-dark)}
  .scroll{max-height:600px;overflow:auto}
  .pill{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}
  .pill.alloc{background:#e3f3e9;color:var(--green-dark)}
  .pill.un{background:#fbe6e3;color:var(--red)}
  .od-late{color:var(--red);font-weight:700}
  .od-ahead{color:var(--green);font-weight:700}
  .od-on{color:var(--amber);font-weight:700}
  .foot{color:var(--muted);font-size:11px;text-align:right;margin-top:8px}
  .seg{display:flex;gap:6px;margin-bottom:10px}
  .seg .fbtn{margin-bottom:0}
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <span class="period" id="asof"></span>
    <h1>People Visibility &mdash; Project Wise</h1>
    <span class="brand">VIANAAR</span>
  </div>

  <div class="layout">
    <!-- Filters -->
    <div class="panel">
      <h3>Project Name</h3>
      <select id="fProject"><option value="">All</option></select>
      <h3>Project Incharge</h3>
      <select id="fIncharge"><option value="">All</option></select>
      <h3>Team Member</h3>
      <select id="fMember"><option value="">All</option></select>
      <h3>Designation</h3>
      <select id="fDesig"><option value="">All</option></select>
      <h3>Allocation</h3>
      <div class="seg">
        <div class="fbtn active" data-alloc="">All</div>
        <div class="fbtn" data-alloc="alloc">Allocated</div>
        <div class="fbtn" data-alloc="un">Unallocated</div>
      </div>
      <h3>Location</h3>
      <div class="loc-btns" id="locBtns"></div>
      <div class="fbtn reset" id="resetBtn">Reset filters</div>
    </div>

    <!-- Main -->
    <div>
      <div class="kpis">
        <div class="kpi"><div class="label">Project Count</div><div class="val" id="kProjects">0</div></div>
        <div class="kpi"><div class="label">Team Count</div><div class="val" id="kTeam">0</div><div class="sub">Master headcount</div></div>
        <div class="kpi"><div class="label">Unallocated Resources</div><div class="val" id="kUnalloc">0</div><div class="sub">not matched in Data</div></div>
        <div class="kpi"><div class="label">Resource Utilization %</div><div class="val" id="kUtil">0%</div><div class="sub">allocated / team</div></div>
      </div>

      <div class="grid3">
        <div class="chartbox"><h3>Team Members by Location</h3><div class="chartwrap"><canvas id="cLoc"></canvas></div></div>
        <div class="chartbox"><h3>Designation &rarr; Roles</h3><div class="chartwrap"><canvas id="cDesig"></canvas></div></div>
        <div class="chartbox"><h3>Experience in Company</h3><div class="chartwrap"><canvas id="cTen"></canvas></div></div>
      </div>

      <div class="grid2">
        <div class="chartbox"><h3>Projects : Team Count</h3><div class="chartwrap" style="height:620px"><canvas id="cProj"></canvas></div></div>
        <div class="tablewrap">
          <h3>Team Members <span id="tblCount" style="font-weight:400;font-size:12px"></span></h3>
          <div class="scroll">
            <table id="tbl">
              <thead><tr>
                <th>Team Member</th>
                <th>Project</th>
                <th class="num">Allocation %</th>
                <th>Baseline Finish</th>
                <th>Forecasted End</th>
                <th class="num">Over Due Days</th>
              </tr></thead>
              <tbody></tbody>
            </table>
          </div>
        </div>
      </div>
      <div class="foot" id="foot"></div>
    </div>
  </div>
</div>

<script>
const MODEL = /*__DATA__*/;
const SRC = "__SOURCE__";
if (window.ChartDataLabels) Chart.register(window.ChartDataLabels);
const BAR="#8e8ec9", PURPLE="#9b8ec9";
const DONUT=["#2b3a8c","#e07b39","#6a1b8a","#c0392b","#1f9bd6","#27ae60","#999"];
let charts={};
let state={project:"",incharge:"",member:"",desig:"",alloc:"",location:""};

const people=MODEL.people, projects=MODEL.projects, S=MODEL.summary;
document.getElementById("asof").textContent="Data as of "+S.as_of;

// ---- populate filters ----
[...new Set(projects.map(p=>p.name))].sort().forEach(n=>add(fProject,n));
[...new Set(projects.map(p=>p.incharge).filter(Boolean))].sort().forEach(n=>add(fIncharge,n));
[...new Set(people.map(p=>p.name))].sort().forEach(n=>add(fMember,n));
[...new Set(people.map(p=>p.designation).filter(Boolean))].sort().forEach(n=>add(fDesig,n));
const locs=[...new Set(people.map(p=>p.location).filter(Boolean))].sort();
locs.forEach(l=>{const b=document.createElement("div");b.className="fbtn";b.dataset.loc=l;b.textContent=l;locBtns.appendChild(b);});
function add(sel,v){const o=document.createElement("option");o.value=v;o.textContent=v;sel.appendChild(o);}

function inchargeProjects(){
  return state.incharge ? new Set(projects.filter(p=>p.incharge===state.incharge).map(p=>p.name)) : null;
}

// ---- filtering ----
function filtPeople(){
  const ip=inchargeProjects();
  return people.filter(p=>{
    if(state.member && p.name!==state.member) return false;
    if(state.location && p.location!==state.location) return false;
    if(state.desig && p.designation!==state.desig) return false;
    if(state.alloc==="alloc" && !p.allocated) return false;
    if(state.alloc==="un" && p.allocated) return false;
    if(state.project && !(p.projects||[]).includes(state.project)) return false;
    if(ip && !(p.projects||[]).some(x=>ip.has(x))) return false;
    return true;
  });
}
function countBy(arr,key){const m={};arr.forEach(x=>{const v=x[key]||"(blank)";m[v]=(m[v]||0)+1;});return m;}

function render(){
  const fp=filtPeople();
  const ip=inchargeProjects();
  const projShown = projects.filter(p=>(!state.project||p.name===state.project)&&(!ip||ip.has(p.name)));
  kProjects.textContent = projShown.length;
  kTeam.textContent = fp.length;
  const un = fp.filter(p=>!p.allocated).length;
  kUnalloc.textContent = un;
  kUtil.textContent = fp.length ? Math.round((fp.length-un)/fp.length*100)+"%" : "0%";

  drawBar("cLoc", countBy(fp,"location"), BAR, true);
  drawBar("cDesig", topN(countBy(fp,"designation"),10), BAR, true);
  drawDonut("cTen", countBy(fp,"tenure"));
  const pj={};
  projShown.forEach(p=>{ pj[p.name]=p.total_people; });
  drawBar("cProj", pj, PURPLE, true);

  buildTable(fp,ip);
  foot.textContent="Source: "+SRC+"  |  Allocation % = 1 / projects per person  |  Incharge counts as team member  |  Over Due = Forecasted End - Baseline Finish";
}

function buildTable(fp,ip){
  const tb=document.querySelector("#tbl tbody");tb.innerHTML="";
  fp.forEach(p=>{
    let asg=(p.assignments||[]).filter(a=>(!state.project || a.project===state.project) && (!ip || ip.has(a.project)));
    if(asg.length===0){
      if(state.project || ip) return;
      const tr=document.createElement("tr");tr.className="grp-start";
      tr.innerHTML=`<td class="name-cell">${esc(p.name)}</td>
        <td><span class="pill un">Unallocated</span></td>
        <td class="num">-</td><td>-</td><td>-</td><td class="num">-</td>`;
      tb.appendChild(tr);return;
    }
    asg.forEach((a,i)=>{
      const tr=document.createElement("tr");
      if(i===0) tr.className="grp-start";
      const nameCell = i===0 ? `<td class="name-cell" rowspan="${asg.length}">${esc(p.name)}</td>` : "";
      tr.innerHTML=`${nameCell}
        <td>${esc(a.project)}</td>
        <td class="num">${a.allocation!=null?a.allocation+"%":"-"}</td>
        <td>${esc(a.baseline_finish)||"-"}</td>
        <td>${esc(a.forecast_end)||"-"}</td>
        <td class="num">${overdueCell(a.overdue)}</td>`;
      tb.appendChild(tr);
    });
  });
  tblCount.textContent="("+fp.length+" people)";
}

function overdueCell(od){
  if(od==null) return "-";
  if(od>0)  return `<span class="od-late">&#9660; ${od}</span>`;
  if(od<0)  return `<span class="od-ahead">&#9650; ${od}</span>`;
  return `<span class="od-on">&#9644; 0</span>`;
}

function topN(obj,n){return Object.fromEntries(Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,n));}
function esc(s){return (s==null?'':String(s)).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));}

function drawBar(id,obj,color,horizontal){
  let ent=Object.entries(obj).sort((a,b)=>b[1]-a[1]);
  if(charts[id])charts[id].destroy();
  charts[id]=new Chart(document.getElementById(id),{
    type:"bar",
    data:{labels:ent.map(e=>e[0]),datasets:[{data:ent.map(e=>e[1]),backgroundColor:color,
          borderRadius:3,maxBarThickness:26}]},
    options:{indexAxis:horizontal?'y':'x',responsive:true,maintainAspectRatio:false,
      layout:{padding:{right:28,top:14}},
      plugins:{legend:{display:false},
        datalabels:{anchor:'end',align:horizontal?'right':'top',
          color:'#1f2d27',font:{weight:'bold',size:11},formatter:v=>v}},
      scales:{x:{grid:{display:false},ticks:{font:{size:11}},beginAtZero:true},
              y:{grid:{display:false},ticks:{font:{size:11}}}}}
  });
}
function drawDonut(id,obj){
  const order=["Below 1 yr","1 - 3 yrs","3.1 - 7 yrs","7.1 - 10 yrs","10.1 - 15 yrs","15+ yrs","Unknown"];
  const ent=Object.entries(obj).sort((a,b)=>order.indexOf(a[0])-order.indexOf(b[0]));
  if(charts[id])charts[id].destroy();
  charts[id]=new Chart(document.getElementById(id),{
    type:"doughnut",
    data:{labels:ent.map(e=>e[0]),datasets:[{data:ent.map(e=>e[1]),
          backgroundColor:DONUT,borderWidth:1,borderColor:"#fff"}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:"55%",
      plugins:{legend:{position:"right",labels:{font:{size:10},boxWidth:12}},
        datalabels:{color:'#fff',font:{weight:'bold',size:12},formatter:v=>v>0?v:''}}}
  });
}

// ---- events ----
fProject.onchange=e=>{state.project=e.target.value;render();};
fIncharge.onchange=e=>{state.incharge=e.target.value;render();};
fMember.onchange=e=>{state.member=e.target.value;render();};
fDesig.onchange=e=>{state.desig=e.target.value;render();};
document.querySelectorAll("[data-alloc]").forEach(b=>b.onclick=()=>{
  document.querySelectorAll("[data-alloc]").forEach(x=>x.classList.remove("active"));
  b.classList.add("active");state.alloc=b.dataset.alloc;render();});
document.getElementById("locBtns").addEventListener("click",e=>{
  if(!e.target.dataset.loc)return;
  const wasActive=e.target.classList.contains("active");
  document.querySelectorAll("[data-loc]").forEach(x=>x.classList.remove("active"));
  if(!wasActive){e.target.classList.add("active");state.location=e.target.dataset.loc;}
  else{state.location="";}
  render();});
document.getElementById("resetBtn").onclick=()=>{
  state={project:"",incharge:"",member:"",desig:"",alloc:"",location:""};
  fProject.value="";fIncharge.value="";fMember.value="";fDesig.value="";
  document.querySelectorAll("[data-loc]").forEach(x=>x.classList.remove("active"));
  document.querySelectorAll("[data-alloc]").forEach(x=>x.classList.remove("active"));
  document.querySelector('[data-alloc=""]').classList.add("active");
  render();};

render();
</script>
</body>
</html>
"""


# --------------------------------------------------------------------------- #
def main():
    in_path = sys.argv[1] if len(sys.argv) > 1 else INPUT_PATH
    if not os.path.exists(in_path):
        sys.exit(f"ERROR: input not found:\n  {in_path}\n"
                 f"Edit INPUT_PATH or pass the path as an argument.")
    model = build_model(in_path)
    out_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(os.path.abspath(in_path)), OUTPUT_NAME)
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(render_html(model, os.path.basename(in_path)))
    except PermissionError:
        sys.exit(f"ERROR: cannot write output (is it open in a browser?):\n  {out_path}")
    s = model["summary"]
    print("Dashboard written to:", out_path)
    print(f"  Projects={s['project_count']}  Team={s['team_count']}  "
          f"Deployed cells={s['deployed_cells']}  Unallocated={s['unallocated_count']}  "
          f"Utilization={s['utilization']}%")


if __name__ == "__main__":
    main()
