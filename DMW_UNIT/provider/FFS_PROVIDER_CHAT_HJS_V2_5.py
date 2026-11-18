# DMW - Data Mart Workbench v1.0Beta (Cortex Provider Analyst) - Streamlit chat app
# Co-authored with CoCo
# v1.1: prompt-driven charts (pie / bar / line / heatmap) rendered as inline SVG in the chat

import re
import streamlit as st
import streamlit.components.v1 as components
import _snowflake
import json
import html as _html_lib
import pandas as pd
from snowflake.snowpark.context import get_active_session

SEMANTIC_MODEL_FILE = "@ADM_PRD.CMS_ADM_CGS_MAC.FHER_SV_STAGE/FFS_PROVIDER_DEMO.sv.yaml"
TIMEOUT_MS = 60_000

SAMPLE_QS = [
    "Show the top 20 highest-risk DME providers",
    "Which states have the most HIGH-risk providers?",
    "What is the fraud risk tier distribution?",
    "Show providers with composite score 90 or above",
    "List providers with Terminated or Suspended FAAIR status",
    "What is the average risk score by FAAIR type?",
    "Show HIGH-risk providers in Pennsylvania",
    "Which providers span ownership across 7 or more states?",
    "Show providers with eligibility flag 0 (not eligible)",
    "Top 10 states by average composite risk score?",
    "Show the risk tier distribution as a pie chart",
    "Bar chart of average risk score by FAAIR type",
    "Heatmap of provider counts by state and risk tier",
    "Now show that as a bar chart",
]

session = get_active_session()

# ── Chart intent detection (prompt-driven, no UI controls) ────────────────────
CHART_PATTERNS = [
    ("heat", re.compile(r"\bheat\s*-?\s*maps?\b", re.I)),
    ("pie",  re.compile(r"\b(pie|donut|doughnut)(\s+(chart|graph))?\b", re.I)),
    ("line", re.compile(r"\bline\s*(chart|graph|plot)\b|\btrend\s*(line|chart)\b|\bas\s+a\s+line\b", re.I)),
    ("bar",  re.compile(r"\b(bar|column)\s*(chart|graph|plot)\b|\bas\s+a\s+bar\b|\bhistogram\b", re.I)),
]
GENERIC_CHART_RE = re.compile(r"\b(chart|graph|plot|visuali[sz]e|visuali[sz]ation)\b", re.I)
CHART_NAMES = {"pie": "pie chart", "bar": "bar chart", "line": "line chart", "heat": "heatmap"}

# "show THIS as a pie chart" -> re-render the previous result instead of re-querying
REF_RE = re.compile(r"\b(this|that|these|those|it|above|previous|prior|last|same|current|results?)\b", re.I)

# Strips chart wording so Cortex Analyst only sees the data question
STRIP_RE = re.compile(
    r"\b(?:(?:as|in|into|with|using|via)\s+)?(?:an?\s+)?"
    r"(?:(?:pie|donut|doughnut|bar|column|line)\s*(?:chart|graph|plot)"
    r"|heat\s*-?\s*maps?|histogram|pie|donut|doughnut"
    r"|chart|graph|plot|visuali[sz]e|visuali[sz]ation)s?\b[,:]?",
    re.I,
)

def detect_chart_type(prompt):
    for ctype, rx in CHART_PATTERNS:
        if rx.search(prompt):
            return ctype
    if GENERIC_CHART_RE.search(prompt):
        return "bar"
    return None

def clean_for_analyst(prompt, ctype):
    s = STRIP_RE.sub(" ", prompt)
    s = re.sub(r"\s+", " ", s).strip(" ,.:;")
    s = re.sub(r"^(of|for|about)\s+", "Show ", s, flags=re.I)
    s = re.sub(r"\b(show|give|display|plot|draw|make|create)(\s+me)?\s+(of|for|about)\b",
               r"\1\2", s, flags=re.I)
    if len(s) < 8:
        s = prompt
    if ctype == "heat":
        hint = " Return two grouping columns (for example state and risk tier) and one numeric measure."
    else:
        hint = " Return one category column and numeric measure column(s)."
    return s + "." + hint

def last_result_df(messages):
    for m in reversed(messages):
        df = m.get("df")
        if m.get("role") == "analyst" and df is not None and not isinstance(df, str) and not df.empty:
            return df
    return None

# ── Chart data preparation ────────────────────────────────────────────────────
ID_RE = re.compile(r"(^|_)(ID|NPI|PTAN|TIN|ZIP|CCN|CD|CODE)($|_)")
MEASURE_RE = re.compile(r"COUNT|SCORE|AVG|TOTAL|AMT|AMOUNT|PCT|PERCENT|NUM|SUM|RATE")
MEAN_RE = re.compile(r"AVG|AVERAGE|MEAN|SCORE|PCT|PERCENT|RATE")
TIME_RE = re.compile(r"DATE|MONTH|YEAR|QUARTER|PERIOD|WEEK|DAY")
MAX_ROWS = 200

def _f(x):
    try:
        v = float(x)
    except Exception:
        return None
    if v != v or v in (float("inf"), float("-inf")):
        return None
    return round(v, 4)

def _nice(name):
    return str(name).replace("_", " ")

def _classify(df):
    cats, nums = [], []
    for c in df.columns:
        if ID_RE.search(str(c).upper()) or pd.api.types.is_datetime64_any_dtype(df[c]):
            cats.append(c)
            continue
        s = pd.to_numeric(df[c], errors="coerce")
        if len(df) and s.notna().mean() >= 0.8:
            nums.append(c)
        else:
            cats.append(c)
    return cats, nums

def _pick_series(df, nums, limit=3):
    chosen = [nums[0]]
    base = df[nums[0]].abs().max()
    if pd.notna(base) and base > 0:
        for c in nums[1:limit]:
            m = df[c].abs().max()
            if pd.notna(m) and m > 0 and 0.05 <= m / base <= 20:
                chosen.append(c)
    return chosen

def _heat_spec(df, cats, nums):
    meas = nums[0]
    if len(cats) >= 2:
        r, c = cats[0], cats[1]
        agg = "mean" if MEAN_RE.search(meas.upper()) else "sum"
        pv = df.pivot_table(index=r, columns=c, values=meas, aggfunc=agg)
        if pv.empty:
            return None, "Not enough data to build a heatmap."
        rows_order = pv.sum(axis=1, skipna=True).sort_values(ascending=False).index[:25]
        cols_order = pv.sum(axis=0, skipna=True).sort_values(ascending=False).index[:15]
        pv = pv.loc[rows_order, cols_order]
        rows = [str(x) for x in pv.index]
        cols = [str(x) for x in pv.columns]
        matrix = [[_f(v) for v in row] for row in pv.values.tolist()]
        colnorm = False
        title = _nice(meas) + " \u2014 " + _nice(r) + " \u00d7 " + _nice(c)
    else:
        use = nums[:12]
        d = df.head(25)
        if cats:
            rows = [str(x) for x in d[cats[0]].tolist()]
        else:
            rows = ["Row " + str(i + 1) for i in range(len(d))]
        cols = [_nice(c) for c in use]
        matrix = [[_f(v) for v in row] for row in d[use].values.tolist()]
        colnorm = len(use) > 1
        title = ", ".join(cols[:3]) + (" by " + _nice(cats[0]) if cats else "")
    return {"type": "heat", "title": title, "rows": rows, "cols": cols,
            "matrix": matrix, "colnorm": colnorm}, ""

def chart_spec(df, ctype):
    """Return (spec_dict, note). spec is None when the data cannot be charted."""
    if df is None or isinstance(df, str) or df.empty:
        return None, "No data available to chart."
    df = df.head(MAX_ROWS).copy()
    cats, nums = _classify(df)
    if not nums:
        return None, "This result has no numeric column to chart."
    nums.sort(key=lambda c: 0 if MEASURE_RE.search(str(c).upper()) else 1)
    for c in nums:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in cats:
        df[c] = df[c].astype(str)

    if ctype == "heat":
        return _heat_spec(df, cats, nums)

    if cats:
        lab = cats[0]
    else:
        lab = "Row"
        df[lab] = ["Row " + str(i + 1) for i in range(len(df))]

    cols = [nums[0]] if ctype == "pie" else _pick_series(df, nums)
    if ctype == "line" and TIME_RE.search(str(lab).upper()):
        df = df.sort_values(lab)
    if ctype in ("pie", "bar") and df[lab].duplicated().any():
        agg = "mean" if any(MEAN_RE.search(str(c).upper()) for c in cols) else "sum"
        df = df.groupby(lab, sort=False)[cols].agg(agg).reset_index()

    if ctype == "pie":
        meas = cols[0]
        d = df[[lab, meas]].dropna()
        d = d[d[meas] > 0].sort_values(meas, ascending=False)
        if d.empty:
            return None, "Pie charts need positive values; none were found in this result."
        labels = [str(x) for x in d[lab].tolist()]
        values = [_f(v) for v in d[meas].tolist()]
        if len(labels) > 10:
            other = _f(sum(v for v in values[9:] if v is not None))
            labels, values = labels[:9] + ["Other"], values[:9] + [other]
        colors = None
        if any(l.upper() in RISK_COLORS for l in labels):
            colors = [RISK_COLORS.get(l.upper(), "#9ca3af") for l in labels]
        return {"type": "pie", "title": _nice(meas) + " by " + _nice(lab),
                "labels": labels, "values": values, "colors": colors}, ""

    limit = 25 if ctype == "bar" else 60
    d = df.head(limit)
    return {"type": ctype,
            "title": ", ".join(_nice(c) for c in cols) + " by " + _nice(lab),
            "labels": [str(x) for x in d[lab].tolist()],
            "series": [{"name": _nice(c), "values": [_f(v) for v in d[c].tolist()]} for c in cols]}, ""

# ── Cortex Analyst helpers ────────────────────────────────────────────────────
def call_analyst(messages):
    return _snowflake.send_snow_api_request(
        "POST",
        "/api/v2/cortex/analyst/message",
        {},
        {},
        {"messages": messages, "semantic_model_file": SEMANTIC_MODEL_FILE},
        {},
        TIMEOUT_MS,
    )

def _parse_sse_blocks(content):
    accumulated = []
    for line in (content.splitlines() if isinstance(content, str) else []):
        line = line.strip()
        if not line.startswith("data:"):
            continue
        payload = line[5:].strip()
        if payload in ("[DONE]", ""):
            continue
        try:
            event = json.loads(payload)
        except Exception:
            continue
        msg = event.get("message")
        if msg and msg.get("content"):
            return msg["content"]
        for block in event.get("content", []):
            btype = block.get("type")
            if btype == "text_delta":
                delta = block.get("text_delta", "")
                if accumulated and accumulated[-1].get("type") == "text":
                    accumulated[-1]["text"] += delta
                else:
                    accumulated.append({"type": "text", "text": delta})
            elif block:
                accumulated.append(block)
    return accumulated or None

def extract_reply(raw):
    sql = None
    text_parts = []
    try:
        content = raw.get("content")
        blocks = None
        if isinstance(content, str):
            try:
                data = json.loads(content)
                msg = data.get("message")
                blocks = msg.get("content") if isinstance(msg, dict) else None
            except json.JSONDecodeError:
                pass
            if not blocks:
                blocks = _parse_sse_blocks(content)
        elif isinstance(content, list):
            if content and isinstance(content[0], dict) and "content" in content[0]:
                blocks = []
                for event in content:
                    blocks.extend(event.get("content", []))
            else:
                blocks = content
        elif isinstance(content, dict):
            msg = content.get("message")
            blocks = msg.get("content") if isinstance(msg, dict) else None
        if not blocks:
            return "No content in response.", None
        for block in blocks:
            btype = block.get("type")
            if btype == "text":
                val = block.get("text", "")
                if isinstance(val, str):
                    text_parts.append(val.strip())
            elif btype == "sql":
                if not sql:
                    sql = block.get("statement", "")
            elif btype == "tool_results":
                for tr in block.get("tool_results", []):
                    for item in tr.get("content", []):
                        if item.get("type") == "json":
                            j = item.get("json", {})
                            if "sql" in j and not sql:
                                sql = j["sql"]
        answer = " ".join(text_parts).strip()
        if not answer and sql:
            answer = "Query generated — see SQL and results below."
        elif not answer:
            answer = "No text in response."
        return answer, sql
    except Exception as exc:
        return "Error parsing response: " + str(exc), None

def run_sql(sql):
    try:
        return session.sql(sql).to_pandas()
    except Exception as exc:
        return str(exc)

# ── HTML / table rendering ────────────────────────────────────────────────────
RISK_COLORS = {"HIGH": "#d9534f", "MEDIUM": "#f0ad4e", "LOW": "#5cb85c"}

def df_to_html_table(df):
    risk_col = next((c for c in df.columns if c.upper() in ("RISK_TIER", "COMPOSITE_TIER")), None)
    score_cols = {c for c in df.columns if "SCORE" in c.upper()}
    headers = "".join("<th>" + _html_lib.escape(str(c)) + "</th>" for c in df.columns)
    rows = ""
    for _, row in df.iterrows():
        tier = str(row.get(risk_col, "")) if risk_col else ""
        row_color = RISK_COLORS.get(tier, "")
        cells = ""
        for col in df.columns:
            v = row[col]
            style = ""
            if col == risk_col and row_color:
                style = " style=\"color:" + row_color + ";font-weight:700;\""
            elif col in score_cols and row_color:
                style = " style=\"color:" + row_color + ";\""
            cells += "<td" + style + ">" + _html_lib.escape(str(v)) + "</td>"
        rows += "<tr>" + cells + "</tr>"
    return (
        "<div class=\"tbl-wrap\"><table>"
        "<thead><tr>" + headers + "</tr></thead>"
        "<tbody>" + rows + "</tbody>"
        "</table></div>"
    )

# Client-side SVG chart renderers (no external library). Raw string: JS escapes are kept as-is.
CHART_JS = r"""
var PAL=["#4a7ba7","#f0ad4e","#5cb85c","#d9534f","#8e6bbf","#2aa7a0","#e377a2","#7f7f7f","#bcbd22","#17becf","#9c6b4e","#6c8ebf"];
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function trunc(s,n){s=String(s);return s.length>n?s.slice(0,n-1)+'\u2026':s;}
function fmt(v){
  if(v===null||v===undefined)return '\u2013';
  var a=Math.abs(v);
  if(a>=1e6)return (v/1e6).toFixed(1)+'M';
  if(a>=1e4)return (v/1e3).toFixed(1)+'K';
  return (Math.round(v*100)/100).toLocaleString();
}
function svgOpen(w,h){
  return '<svg viewBox="0 0 '+w+' '+h+'" width="100%" style="max-width:'+w+'px;display:block" xmlns="http://www.w3.org/2000/svg" font-family="Segoe UI,Tahoma,sans-serif">';
}
function pieSvg(c){
  var W=560,H=250,cx=125,cy=125,r=105,tot=0,i,cols=c.colors||PAL;
  c.values.forEach(function(v){tot+=(v||0);});
  var out=svgOpen(W,H),a0=-Math.PI/2;
  for(i=0;i<c.values.length;i++){
    var v=c.values[i]||0,frac=v/tot,a1=a0+frac*2*Math.PI,col=cols[i%cols.length];
    var tip=esc(c.labels[i])+': '+fmt(v)+' ('+(frac*100).toFixed(1)+'%)';
    if(c.values.length===1){
      out+='<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="'+col+'"><title>'+tip+'</title></circle>';
    }else{
      var x0=cx+r*Math.cos(a0),y0=cy+r*Math.sin(a0),x1=cx+r*Math.cos(a1),y1=cy+r*Math.sin(a1);
      out+='<path d="M'+cx+','+cy+' L'+x0.toFixed(2)+','+y0.toFixed(2)+' A'+r+','+r+' 0 '+(a1-a0>Math.PI?1:0)+' 1 '+x1.toFixed(2)+','+y1.toFixed(2)+' Z" fill="'+col+'" stroke="#fff" stroke-width="1.5"><title>'+tip+'</title></path>';
    }
    a0=a1;
    var ly=22+i*20;
    out+='<rect x="265" y="'+(ly-9)+'" width="10" height="10" rx="2" fill="'+col+'"/>'
      +'<text x="281" y="'+ly+'" font-size="11" fill="#374151">'+esc(trunc(c.labels[i],26))+'</text>'
      +'<text x="545" y="'+ly+'" font-size="11" fill="#6B7280" text-anchor="end">'+fmt(v)+' \u00b7 '+(frac*100).toFixed(1)+'%</text>';
  }
  return out+'</svg>';
}
function barSvg(c){
  var n=c.labels.length,k=c.series.length,W=560,lw=140,rp=56,top=8,bot=24,bh=12,gap=8;
  var rowH=k*bh+gap,H=top+n*rowH+bot,pw=W-lw-rp,lo=0,hi=0,i,j,t;
  c.series.forEach(function(s){s.values.forEach(function(v){if(v!==null){if(v<lo)lo=v;if(v>hi)hi=v;}});});
  if(hi===lo)hi=lo+1;
  function X(v){return lw+(v-lo)/(hi-lo)*pw;}
  var out=svgOpen(W,H);
  for(t=0;t<=4;t++){
    var gv=lo+(hi-lo)*t/4,gx=X(gv).toFixed(1);
    out+='<line x1="'+gx+'" y1="'+top+'" x2="'+gx+'" y2="'+(H-bot)+'" stroke="#eef0f3"/>'
      +'<text x="'+gx+'" y="'+(H-bot+13)+'" font-size="9" fill="#9CA3AF" text-anchor="middle">'+fmt(gv)+'</text>';
  }
  for(i=0;i<n;i++){
    var y=top+i*rowH;
    out+='<text x="'+(lw-6)+'" y="'+(y+k*bh/2+3)+'" font-size="10.5" fill="#374151" text-anchor="end"><title>'+esc(c.labels[i])+'</title>'+esc(trunc(c.labels[i],22))+'</text>';
    for(j=0;j<k;j++){
      var v=c.series[j].values[i];
      if(v===null)continue;
      var x0=X(0),x1=X(v),bx=Math.min(x0,x1),bw=Math.max(Math.abs(x1-x0),1);
      out+='<rect x="'+bx.toFixed(1)+'" y="'+(y+j*bh)+'" width="'+bw.toFixed(1)+'" height="'+(bh-2)+'" rx="2" fill="'+PAL[j%PAL.length]+'"><title>'+esc(c.labels[i])+' \u2014 '+esc(c.series[j].name)+': '+fmt(v)+'</title></rect>'
        +'<text x="'+(v>=0?x1+4:x1-4).toFixed(1)+'" y="'+(y+j*bh+8)+'" font-size="9" fill="#6B7280" text-anchor="'+(v>=0?'start':'end')+'">'+fmt(v)+'</text>';
    }
  }
  return out+'</svg>';
}
function lineSvg(c){
  var n=c.labels.length,W=560,H=270,l=52,r=14,t=14,b=64,pw=W-l-r,ph=H-t-b,lo=Infinity,hi=-Infinity,i,j,g;
  c.series.forEach(function(s){s.values.forEach(function(v){if(v!==null){if(v<lo)lo=v;if(v>hi)hi=v;}});});
  if(lo===Infinity){lo=0;hi=1;}
  if(lo>0)lo=0;
  if(hi<=lo)hi=lo+1;
  function X(i){return n===1?l+pw/2:l+i*pw/(n-1);}
  function Y(v){return t+ph-(v-lo)/(hi-lo)*ph;}
  var out=svgOpen(W,H);
  for(g=0;g<=4;g++){
    var gv=lo+(hi-lo)*g/4,gy=Y(gv).toFixed(1);
    out+='<line x1="'+l+'" y1="'+gy+'" x2="'+(W-r)+'" y2="'+gy+'" stroke="#eef0f3"/>'
      +'<text x="'+(l-6)+'" y="'+(parseFloat(gy)+3)+'" font-size="9" fill="#9CA3AF" text-anchor="end">'+fmt(gv)+'</text>';
  }
  var step=Math.ceil(n/12);
  for(i=0;i<n;i+=step){
    var tx=X(i).toFixed(1),ty=H-b+14;
    out+='<text x="'+tx+'" y="'+ty+'" font-size="9.5" fill="#6B7280" text-anchor="end" transform="rotate(-35 '+tx+' '+ty+')">'+esc(trunc(c.labels[i],16))+'</text>';
  }
  for(j=0;j<c.series.length;j++){
    var col=PAL[j%PAL.length],pts=[],dots='';
    for(i=0;i<n;i++){
      var v=c.series[j].values[i];
      if(v===null)continue;
      var px=X(i).toFixed(1),py=Y(v).toFixed(1);
      pts.push(px+','+py);
      dots+='<circle cx="'+px+'" cy="'+py+'" r="3" fill="'+col+'"><title>'+esc(c.labels[i])+' \u2014 '+esc(c.series[j].name)+': '+fmt(v)+'</title></circle>';
    }
    if(pts.length>1)out+='<polyline points="'+pts.join(' ')+'" fill="none" stroke="'+col+'" stroke-width="2"/>';
    out+=dots;
  }
  return out+'</svg>';
}
function heatSvg(c){
  var R=c.rows.length,C=c.cols.length,lw=140,top=70,ch=20,i,j;
  var cw=Math.max(26,Math.min(64,(560-lw-10)/Math.max(C,1)));
  var W=Math.max(560,lw+C*cw+60),H=top+R*ch+8;
  var gLo=Infinity,gHi=-Infinity,cLo=[],cHi=[];
  for(j=0;j<C;j++){cLo[j]=Infinity;cHi[j]=-Infinity;}
  for(i=0;i<R;i++)for(j=0;j<C;j++){
    var v=c.matrix[i][j];
    if(v===null)continue;
    if(v<gLo)gLo=v;if(v>gHi)gHi=v;
    if(v<cLo[j])cLo[j]=v;if(v>cHi[j])cHi[j]=v;
  }
  function tOf(v,j){
    var lo=c.colnorm?cLo[j]:gLo,hi=c.colnorm?cHi[j]:gHi;
    return hi>lo?(v-lo)/(hi-lo):0.5;
  }
  function colr(t){
    var a=[240,245,250],b=[217,83,79],m=[];
    for(var k=0;k<3;k++)m.push(Math.round(a[k]+(b[k]-a[k])*t));
    return 'rgb('+m.join(',')+')';
  }
  var out=svgOpen(W,H);
  for(j=0;j<C;j++){
    var hx=(lw+j*cw+cw/2).toFixed(1);
    out+='<text x="'+hx+'" y="'+(top-8)+'" font-size="9.5" fill="#374151" transform="rotate(-35 '+hx+' '+(top-8)+')"><title>'+esc(c.cols[j])+'</title>'+esc(trunc(c.cols[j],16))+'</text>';
  }
  for(i=0;i<R;i++){
    var y=top+i*ch;
    out+='<text x="'+(lw-6)+'" y="'+(y+ch/2+3)+'" font-size="10.5" fill="#374151" text-anchor="end"><title>'+esc(c.rows[i])+'</title>'+esc(trunc(c.rows[i],22))+'</text>';
    for(j=0;j<C;j++){
      var v=c.matrix[i][j],x=lw+j*cw;
      if(v===null){
        out+='<rect x="'+x+'" y="'+y+'" width="'+(cw-2)+'" height="'+(ch-2)+'" rx="2" fill="#f3f4f6"/>';
        continue;
      }
      var t=tOf(v,j);
      out+='<rect x="'+x+'" y="'+y+'" width="'+(cw-2)+'" height="'+(ch-2)+'" rx="2" fill="'+colr(t)+'"><title>'+esc(c.rows[i])+' \u00d7 '+esc(c.cols[j])+': '+fmt(v)+'</title></rect>';
      if(cw>=34)out+='<text x="'+(x+(cw-2)/2).toFixed(1)+'" y="'+(y+ch/2+2)+'" font-size="9" text-anchor="middle" fill="'+(t>0.6?'#fff':'#374151')+'">'+fmt(v)+'</text>';
    }
  }
  return out+'</svg>';
}
function renderChart(c){
  var d=document.createElement('div');
  d.className='chart-wrap';
  var body=c.type==='pie'?pieSvg(c):c.type==='bar'?barSvg(c):c.type==='line'?lineSvg(c):heatSvg(c);
  var leg='';
  if((c.type==='bar'||c.type==='line')&&c.series.length>1){
    leg='<div class="chart-legend">'+c.series.map(function(s,i){
      return '<span><i style="background:'+PAL[i%PAL.length]+'"></i>'+esc(s.name)+'</span>';
    }).join('')+'</div>';
  }
  if(c.type==='heat'){
    leg='<div class="chart-legend"><span>Low</span><i class="grad"></i><span>High</span>'+(c.colnorm?'<span>(scaled per column)</span>':'')+'</div>';
  }
  d.innerHTML='<div class="chart-title">'+esc(c.title)+'</div>'+body+leg;
  return d;
}
"""

def build_chat_html(messages):
    msgs_data = []
    for m in messages:
        entry = {"role": m["role"], "text": m.get("text", ""),
                 "sql": m.get("sql") or "", "table": "", "error": "",
                 "chart": None, "chartNote": ""}
        df = m.get("df")
        if df is not None:
            if isinstance(df, str):
                entry["error"] = df
            elif not df.empty:
                if not m.get("hide_table"):
                    entry["table"] = df_to_html_table(df)
                if m.get("chart"):
                    try:
                        spec, note = chart_spec(df, m["chart"])
                    except Exception as exc:
                        spec, note = None, "Could not build the chart: " + str(exc)
                    entry["chart"] = spec
                    entry["chartNote"] = note
        elif m.get("chart"):
            entry["chartNote"] = "No result data available to chart."
        msgs_data.append(entry)
    msgs_json = json.dumps(msgs_data).replace("</", "<\\/")

    css = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f0f0f0;font-size:11px;}
.dmw-header{background:#4a7ba7;color:white;padding:11px 16px;
  border-bottom:2px solid #3a6b97;display:flex;align-items:center;
  gap:10px;position:relative;}
.dmw-titles{flex:1}
.dmw-titles h1{font-size:14px;font-weight:600;margin:0;letter-spacing:.1px;}
.dmw-titles p{font-size:10px;opacity:.88;margin-top:2px;}
.top-nav{display:flex;gap:4px;align-items:center;}
.top-nav a{color:white;text-decoration:none;font-size:11px;font-weight:600;
  padding:4px 10px;border-radius:3px;transition:background .2s;}
.top-nav a:hover{background:rgba(255,255,255,.18);}
.badge{background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.3);
  border-radius:20px;padding:3px 10px;font-size:10px;white-space:nowrap;margin-left:6px;}
.modal-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;
  background:rgba(0,0,0,.5);z-index:2000;justify-content:center;align-items:center;}
.modal-overlay.active{display:flex;}
.modal-box{background:white;border-radius:6px;box-shadow:0 8px 24px rgba(0,0,0,.25);
  max-width:480px;width:92%;max-height:78vh;overflow-y:auto;}
.modal-hdr{background:#4a7ba7;color:white;padding:12px 16px;border-radius:6px 6px 0 0;
  display:flex;justify-content:space-between;align-items:center;}
.modal-hdr h2{font-size:13px;font-weight:600;}
.mc{background:transparent;border:none;color:white;font-size:22px;cursor:pointer;
  width:28px;height:28px;display:flex;align-items:center;justify-content:center;
  border-radius:3px;transition:background .2s;}
.mc:hover{background:rgba(255,255,255,.18);}
.modal-body{padding:16px;font-size:12px;line-height:1.65;color:#333;}
.modal-body p{margin-bottom:10px;}
.modal-body code{background:#f0f0f0;padding:1px 5px;border-radius:3px;font-size:11px;}
.modal-body ul{margin:6px 0 10px 18px;}
.modal-body li{margin-bottom:4px;}
.chat-area{height:440px;overflow-y:auto;padding:13px;
  display:flex;flex-direction:column;gap:11px;background:#fff;}
.msg{display:flex;flex-direction:column;max-width:84%}
.msg.wide{width:96%;max-width:96%}
.msg.user{align-self:flex-end;align-items:flex-end}
.msg.analyst{align-self:flex-start;align-items:flex-start}
.msg-label{font-size:10px;font-weight:700;letter-spacing:.6px;
  text-transform:uppercase;color:#6B7280;margin-bottom:4px;}
.bubble{padding:9px 13px;border-radius:12px;
  font-size:12.5px;line-height:1.55;word-break:break-word;}
.msg.user .bubble{background:#4a7ba7;color:#fff;border-bottom-right-radius:3px;}
.msg.analyst .bubble{background:#fff;color:#1F2937;border:1px solid #E5E7EB;
  border-bottom-left-radius:3px;box-shadow:0 1px 3px rgba(0,0,0,.07);}
.sql-wrap{margin-top:7px;width:100%}
.sql-lbl{font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;
  color:#6B7280;margin-bottom:3px;display:flex;align-items:center;gap:4px;}
.sql-lbl::before{content:'';display:inline-block;width:7px;height:7px;
  background:#4a7ba7;border-radius:2px;}
.sql-code{background:#1e2537;color:#CDD5E0;padding:9px 12px;border-radius:6px;
  font-family:'Courier New',monospace;font-size:11px;line-height:1.45;
  white-space:pre-wrap;word-break:break-all;
  border-left:3px solid #4a7ba7;max-height:150px;overflow-y:auto;}
.tbl-wrap{margin-top:7px;overflow-x:auto;border-radius:5px;
  border:1px solid #E5E7EB;max-height:200px;overflow-y:auto;}
table{width:100%;border-collapse:collapse;font-size:11px;background:#fff}
thead{background:#4a7ba7;color:#fff;position:sticky;top:0}
th{padding:7px 9px;text-align:left;font-weight:600;white-space:nowrap}
td{padding:6px 9px;border-bottom:1px solid #F3F4F6;color:#374151}
tr:hover td{background:#f0f5fa}
tr:last-child td{border-bottom:none}
.err-box{margin-top:5px;padding:7px 11px;
  background:#FEF2F2;border:1px solid #FECACA;
  border-radius:5px;color:#DC2626;font-size:11px;}
.note-box{margin-top:5px;padding:6px 10px;background:#FFF8E6;border:1px solid #F5DFA3;
  border-radius:5px;color:#8a6d1d;font-size:11px;}
.chart-wrap{margin-top:7px;width:100%;background:#fff;border:1px solid #E5E7EB;
  border-radius:6px;padding:8px 10px;}
.chart-title{font-size:11px;font-weight:700;color:#374151;margin-bottom:4px;}
.chart-legend{display:flex;gap:12px;flex-wrap:wrap;font-size:10px;color:#6B7280;
  margin-top:4px;align-items:center;}
.chart-legend i{display:inline-block;width:10px;height:10px;border-radius:2px;
  margin-right:4px;vertical-align:-1px;}
.chart-legend i.grad{width:70px;margin:0 4px;
  background:linear-gradient(90deg,#f0f5fa,#d9534f);}
.empty-state{flex:1;display:flex;flex-direction:column;
  align-items:center;justify-content:center;
  color:#9CA3AF;gap:10px;padding:40px;text-align:center;}
.empty-state .ico{font-size:40px}
.empty-state p{font-size:13px;line-height:1.6}
.dmw-footer{padding:5px 16px;border-top:1px solid #ddd;background:#f8f8f8;
  display:flex;align-items:center;justify-content:space-between;
  font-size:11px;color:#666;}
.dot{width:8px;height:8px;border-radius:50%;background:#5cb85c;
  display:inline-block;margin-right:5px;}
"""

    about_body = (
        "<p>The <strong>DMW &ndash; Data Mart Workbench v1.0Beta</strong> "
        "(Cortex Provider Analyst) is powered by Snowflake Cortex Analyst and the "
        "<code>FFS_PROVIDER_DEMO</code> semantic view over "
        "<code>FFS_PROVIDER_DEMO_FLAT</code>.</p>"
        "<p>Ask plain-English questions about ~279&nbsp;K Medicare DME providers. "
        "The app converts your question to SQL, executes it, and returns a "
        "colour-coded results table.</p>"
        "<hr style=\"border:none;border-top:1px solid #e0e0e0;margin:12px 0;\">"
        "<p style=\"font-size:11px;color:#666;\">"
        "<strong>Risk tiers:</strong>&nbsp;"
        "<span style=\"color:#d9534f;font-weight:700;\">HIGH</span> score&nbsp;&ge;&nbsp;40 &nbsp;|&nbsp; "
        "<span style=\"color:#f0ad4e;font-weight:700;\">MEDIUM</span> 25&ndash;39 &nbsp;|&nbsp; "
        "<span style=\"color:#5cb85c;font-weight:700;\">LOW</span> &lt;&nbsp;25</p>"
        "<p style=\"font-size:11px;color:#666;\">The DMW is modeled on the CMS Episode Workbench. "
        "Long descriptions are AI-generated and should be independently verified.</p>"
    )

    help_body = (
        "<p>Type a question about DME provider fraud risk in the input box and press "
        "<strong>Send</strong>, or click a sample question in the sidebar.</p>"
        "<p><strong>Charts:</strong> just ask for one in the chat &mdash; "
        "<em>&ldquo;risk tier distribution as a pie chart&rdquo;</em>, "
        "<em>&ldquo;bar chart of average score by FAAIR type&rdquo;</em>, "
        "<em>&ldquo;line chart of&hellip;&rdquo;</em>, "
        "<em>&ldquo;heatmap of provider counts by state and risk tier&rdquo;</em>. "
        "To re-draw the last result, say <em>&ldquo;show this as a bar chart&rdquo;</em>.</p>"
        "<p><strong>Key columns:</strong></p>"
        "<ul>"
        "<li><code>FRAUD_RISK_SCORE</code> &mdash; base risk score 0&ndash;100</li>"
        "<li><code>MULTI_STATE_RISK_SCORE</code> &mdash; enhanced score including state footprint</li>"
        "<li><code>RISK_TIER</code> &mdash; HIGH / MEDIUM / LOW</li>"
        "<li><code>PRVDR_DME_FAAIR_TYPE_CD</code> &mdash; A=Standard, R=Revoked, S=Suspended, T=Terminated</li>"
        "<li><code>STATE_COUNT</code> &mdash; states spanned in ownership records</li>"
        "<li><code>CLAIM_COUNT</code> &mdash; matched Medicare claims</li>"
        "</ul>"
        "<p><strong>Contact:</strong> Reach out to Manuel Figallo for support.</p>"
    )

    page = (
        "<!DOCTYPE html><html lang=\"en\"><head>"
        "<meta charset=\"UTF-8\">"
        "<style>" + css + "</style>"
        "</head><body>"
        "<div class=\"modal-overlay\" id=\"aboutM\">"
        "<div class=\"modal-box\">"
        "<div class=\"modal-hdr\"><h2>About DMW Data Mart Workbench</h2>"
        "<button class=\"mc\" onclick=\"cm('aboutM')\">&times;</button></div>"
        "<div class=\"modal-body\">" + about_body + "</div>"
        "</div></div>"
        "<div class=\"modal-overlay\" id=\"helpM\">"
        "<div class=\"modal-box\">"
        "<div class=\"modal-hdr\"><h2>Help &amp; Support</h2>"
        "<button class=\"mc\" onclick=\"cm('helpM')\">&times;</button></div>"
        "<div class=\"modal-body\">" + help_body + "</div>"
        "</div></div>"
        "<div class=\"dmw-header\">"
        "<div class=\"dmw-titles\">"
        "<h1>DMW &ndash; Data Mart Workbench v1.0Beta -- Provider Risk Agent</h1>"
        "<p>Cortex Provider Analyst &middot; ADM_PRD.CMS_ADM_CGS_MAC"
        " &middot; ~279K DME Providers</p>"
        "</div>"
        "<div class=\"top-nav\">"
        "<a href=\"#\" onclick=\"om('aboutM');return false;\">About</a>"
        "<a href=\"#\" onclick=\"om('helpM');return false;\">Help</a>"
        "</div>"
        "<span class=\"badge\">&#10003; Session Active</span>"
        "</div>"
        "<div class=\"chat-area\" id=\"chatArea\"></div>"
        "<div class=\"dmw-footer\">"
        "<span><span class=\"dot\"></span>Cortex Analyst &bull; FFS_PROVIDER_DEMO</span>"
        "<span>ADM_PRD.CMS_ADM_CGS_MAC.FFS_PROVIDER_DEMO_FLAT</span>"
        "</div>"
        "<script>"
        "function om(id){document.getElementById(id).classList.add('active');}"
        "function cm(id){document.getElementById(id).classList.remove('active');}"
        "document.addEventListener('click',function(e){"
        "  document.querySelectorAll('.modal-overlay').forEach(function(ov){"
        "    if(e.target===ov)ov.classList.remove('active');"
        "  });"
        "});"
        "document.addEventListener('keydown',function(e){"
        "  if(e.key==='Escape') document.querySelectorAll('.modal-overlay')"
        "    .forEach(function(ov){ov.classList.remove('active');});"
        "});"
        + CHART_JS +
        "const msgs="
    )

    js_tail = (
        ";"
        "const area=document.getElementById('chatArea');"
        "if(!msgs.length){"
        "  area.innerHTML='<div class=\"empty-state\"><div class=\"ico\">&#128202;</div>"
        "<p>Ask a plain-English question about Medicare DME provider fraud risk.<br>"
        "Add &ldquo;as a pie chart&rdquo;, &ldquo;bar chart&rdquo;, &ldquo;line chart&rdquo; "
        "or &ldquo;heatmap&rdquo; to see it visually.</p></div>';"
        "}else{"
        "  msgs.forEach(function(m){"
        "    const wrap=document.createElement('div');"
        "    wrap.className='msg '+(m.role==='user'?'user':'analyst')+(m.chart?' wide':'');"
        "    const lbl=document.createElement('div');"
        "    lbl.className='msg-label';"
        "    lbl.textContent=m.role==='user'?'You':'Cortex Analyst';"
        "    const bub=document.createElement('div');"
        "    bub.className='bubble';"
        "    bub.textContent=m.text;"
        "    wrap.appendChild(lbl);wrap.appendChild(bub);"
        "    if(m.chart){"
        "      try{wrap.appendChild(renderChart(m.chart));}"
        "      catch(e){const nb=document.createElement('div');nb.className='note-box';"
        "        nb.textContent='Chart could not be drawn: '+e;wrap.appendChild(nb);}"
        "    }"
        "    if(m.chartNote){"
        "      const nd=document.createElement('div');nd.className='note-box';"
        "      nd.textContent=m.chartNote;wrap.appendChild(nd);"
        "    }"
        "    if(m.sql){"
        "      const sw=document.createElement('div');sw.className='sql-wrap';"
        "      const sl=document.createElement('div');sl.className='sql-lbl';"
        "      sl.textContent='Generated SQL';"
        "      const sc=document.createElement('div');sc.className='sql-code';"
        "      sc.textContent=m.sql;"
        "      sw.appendChild(sl);sw.appendChild(sc);wrap.appendChild(sw);"
        "    }"
        "    if(m.table){"
        "      const td=document.createElement('div');td.innerHTML=m.table;"
        "      wrap.appendChild(td);"
        "    }"
        "    if(m.error){"
        "      const ed=document.createElement('div');ed.className='err-box';"
        "      ed.textContent='SQL error: '+m.error;wrap.appendChild(ed);"
        "    }"
        "    area.appendChild(wrap);"
        "  });"
        "  area.scrollTop=area.scrollHeight;"
        "}"
        "</script>"
        "</body></html>"
    )

    return page + msgs_json + js_tail


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="DMW – Cortex Provider Analyst",
                   page_icon="📊", layout="wide")
st.markdown("""
<style>
  .block-container{padding-top:.75rem;padding-bottom:0}
  #MainMenu{visibility:hidden} footer{visibility:hidden} header{visibility:hidden}
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = []

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        "<div style=\"background:#4a7ba7;color:white;padding:8px 12px;border-radius:4px;"
        "font-weight:700;font-size:13px;margin-bottom:12px;\">"
        "Sample Questions</div>",
        unsafe_allow_html=True,
    )
    st.caption("Click any question to ask it:")
    for q in SAMPLE_QS:
        if st.button(q, use_container_width=True, key=q):
            st.session_state.pending_prompt = q
    st.divider()
    if st.button("🗑 Clear Conversation", use_container_width=True):
        st.session_state.messages = []
        st.rerun()
    st.divider()
    st.markdown(
        "<div style=\"font-size:11px;color:#555;line-height:1.8;\">"
        "<b>App:</b> DMW v1.0Beta (Cortex Provider Analyst)<br>"
        "<b>Semantic view:</b> FFS_PROVIDER_DEMO<br>"
        "<b>Table:</b> FFS_PROVIDER_DEMO_FLAT<br>"
        "<b>Providers:</b> ~279K DME providers<br>"
        "<b>Warehouse:</b> OPI_PRD_SAS_1_WH"
        "</div>",
        unsafe_allow_html=True,
    )
    st.divider()
    st.markdown(
        "<div style=\"font-size:10px;color:#888;line-height:1.8;\">"
        "<b>Risk tier legend:</b><br>"
        "<span style=\"color:#d9534f;font-weight:700;\">HIGH</span>&nbsp;&ge;&nbsp;40&emsp;"
        "<span style=\"color:#f0ad4e;font-weight:700;\">MEDIUM</span>&nbsp;25&ndash;39&emsp;"
        "<span style=\"color:#5cb85c;font-weight:700;\">LOW</span>&nbsp;&lt;&nbsp;25"
        "</div>",
        unsafe_allow_html=True,
    )

# ── Chat component ────────────────────────────────────────────────────────────
components.html(build_chat_html(st.session_state.messages), height=565, scrolling=False)

# ── Input form ────────────────────────────────────────────────────────────────
with st.form("chat_form", clear_on_submit=True):
    c1, c2 = st.columns([9, 1])
    with c1:
        typed = st.text_input("Ask about DME provider fraud risk...",
                              placeholder="e.g. Show top 20 highest-risk providers, or: heatmap of risk by state",
                              label_visibility="collapsed")
    with c2:
        submitted = st.form_submit_button("Send", use_container_width=True)

prompt = typed if submitted and typed.strip() else None
if "pending_prompt" in st.session_state:
    prompt = st.session_state.pop("pending_prompt")

if prompt:
    ctype = detect_chart_type(prompt)
    prev_df = last_result_df(st.session_state.messages)
    # "show this as a pie chart" -> redraw the previous result without re-querying
    reuse = bool(ctype and prev_df is not None and REF_RE.search(prompt))

    user_msg = {"role": "user", "text": prompt}
    if ctype and not reuse:
        user_msg["api_text"] = clean_for_analyst(prompt, ctype)
    if reuse:
        user_msg["local"] = True
    st.session_state.messages.append(user_msg)

    if reuse:
        st.session_state.messages.append({
            "role": "analyst",
            "text": "Rendered the previous result as a " + CHART_NAMES[ctype] + ".",
            "sql": "", "df": prev_df, "chart": ctype,
            "hide_table": True, "local": True,
        })
    else:
        api_msgs = [
            {"role": "user" if m["role"] == "user" else "analyst",
             "content": [{"type": "text", "text": m.get("api_text", m["text"])}]}
            for m in st.session_state.messages
            if m["role"] in ("user", "analyst") and not m.get("local")
        ]
        with st.spinner("Querying provider data..."):
            raw = call_analyst(api_msgs)
            answer, sql = extract_reply(raw)
            df = run_sql(sql) if sql else None
        reply = {"role": "analyst", "text": answer, "sql": sql or "", "df": df}
        if ctype:
            reply["chart"] = ctype
        st.session_state.messages.append(reply)
    st.rerun()