# -*- coding: utf-8 -*-
"""从 data.xlsx 生成 index.html — 醉梦六维属性排名（不含实力/集结进分）"""
from __future__ import annotations

from collections import Counter
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    raise SystemExit("缺少 openpyxl，请先: pip install openpyxl")

ROOT = Path(__file__).resolve().parent
DATA_XLSX = ROOT / "data.xlsx"
OUT_HTML = ROOT / "index.html"

# 六维等权归一后综合（也可直接用表内六维综合做排名展示）
# 属性分 = 六维综合 / 本组最大六维 * 100（仅统计六项齐全者）


def to_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('"', "").replace("\n", "").replace("\r", "")
    if not s or s.lower() == "none":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_str(v):
    if v is None:
        return ""
    return str(v).strip().replace("\n", "").replace("\r", "").replace('"', "")


def esc(s):
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def fmt(n, d=2):
    if n is None:
        return "-"
    if isinstance(n, float) and n == int(n):
        return str(int(n))
    if isinstance(n, (int, float)):
        return f"{n:.{d}f}".rstrip("0").rstrip(".")
    return str(n)


def fmt_power(n):
    if n is None:
        return "-"
    return f"{int(n):,}"


def load_members(path: Path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    members = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        name = to_str(row[1] if len(row) > 1 else None)
        if not name:
            continue
        m = {
            "seq": to_num(row[0]),
            "name": name,
            "power": to_num(row[2]) if len(row) > 2 else None,
            "rally": to_num(row[3]) if len(row) > 3 else None,
            "inf_hp": to_num(row[4]) if len(row) > 4 else None,
            "inf_def": to_num(row[5]) if len(row) > 5 else None,
            "arc_dmg": to_num(row[6]) if len(row) > 6 else None,
            "arc_atk": to_num(row[7]) if len(row) > 7 else None,
            "cav_atk": to_num(row[8]) if len(row) > 8 else None,
            "cav_dmg": to_num(row[9]) if len(row) > 9 else None,
            "six": to_num(row[10]) if len(row) > 10 else None,
        }
        # recompute six if attrs present
        attrs = [m["inf_hp"], m["inf_def"], m["arc_dmg"], m["arc_atk"], m["cav_atk"], m["cav_dmg"]]
        if all(a is not None for a in attrs):
            calc = sum(attrs)
            if m["six"] is None or m["six"] == 0:
                m["six"] = round(calc, 2)
            m["inf_sum"] = m["inf_hp"] + m["inf_def"]
            m["arc_sum"] = m["arc_atk"] + m["arc_dmg"]
            m["cav_sum"] = m["cav_atk"] + m["cav_dmg"]
            m["complete"] = True
        else:
            m["complete"] = False
            m["inf_sum"] = m["arc_sum"] = m["cav_sum"] = None
        members.append(m)
    return members, ws.title


def score_members(members):
    complete = [m for m in members if m["complete"] and m["six"] and m["six"] > 0]
    incomplete = [m for m in members if not m["complete"]]

    if not complete:
        raise SystemExit("没有六维齐全的成员，无法生成排名")

    max_six = max(m["six"] for m in complete)
    for m in complete:
        m["score"] = round(m["six"] / max_six * 100, 2)

    complete.sort(key=lambda x: (-x["score"], -(x["power"] or 0), x["name"]))
    for i, m in enumerate(complete, 1):
        m["rank"] = i

    blank = [
        m
        for m in incomplete
        if all(
            m[k] is None
            for k in ("inf_hp", "inf_def", "arc_dmg", "arc_atk", "cav_atk", "cav_dmg", "rally")
        )
    ]
    partial = [m for m in incomplete if m not in blank]

    for m in partial:
        issues = []
        for k, lab in (
            ("inf_hp", "缺步生"),
            ("inf_def", "缺步防"),
            ("arc_dmg", "缺弓破"),
            ("arc_atk", "缺弓攻"),
            ("cav_atk", "缺骑攻"),
            ("cav_dmg", "缺骑破"),
        ):
            if m[k] is None:
                issues.append(lab)
        m["issues"] = issues

    meta = {
        "title": "醉梦成员属性排名",
        "source": DATA_XLSX.name,
        "total": len(members),
        "complete": len(complete),
        "incomplete": len(incomplete),
        "avgScore": round(sum(m["score"] for m in complete) / len(complete), 2),
        "avgSix": round(sum(m["six"] for m in complete) / len(complete), 1),
        "max_six": max_six,
        "formula": f"属性分 = 六维综合 / {max_six} × 100（不含实力、集结加成）",
        "six_def": "六维 = 步生+步防+弓破+弓攻+骑攻+骑破",
    }
    return complete, blank, partial, meta


def render_html(complete, blank, partial, meta):
    max_score = complete[0]["score"]
    bars = []
    for m in complete[:10]:
        pct = round(m["score"] / max_score * 100, 1)
        bars.append(
            f'<div class="bar-row"><div class="bar-label">{esc(m["name"])}</div>'
            f'<div class="bar-track"><div class="bar-fill" style="width:{pct}%"></div></div>'
            f'<div class="bar-val">{fmt(m["score"])}</div></div>'
        )

    rank_rows = []
    for m in complete:
        tone = " top" if m["rank"] <= 10 else ""
        rank_rows.append(
            f'<tr class="{tone.strip()}" data-power="{m["power"] or -1}" data-rally="{m["rally"] or -1}" '
            f'data-six="{m["six"]}" data-inf="{m["inf_sum"]}" data-arc="{m["arc_sum"]}" data-cav="{m["cav_sum"]}" data-score="{m["score"]}">'
            f'<td class="num">{m["rank"]}</td>'
            f'<td>{esc(m["name"])}</td>'
            f'<td class="num">{fmt_power(m["power"])}</td>'
            f'<td class="num">{fmt(m["rally"])}</td>'
            f'<td class="num">{fmt(m["inf_sum"],1)}</td>'
            f'<td class="num">{fmt(m["arc_sum"],1)}</td>'
            f'<td class="num">{fmt(m["cav_sum"],1)}</td>'
            f'<td class="num">{fmt(m["six"],1)}</td>'
            f'<td class="num score">{fmt(m["score"])}</td></tr>'
        )

    top_inf = sorted(complete, key=lambda x: -x["inf_sum"])[:10]
    top_arc = sorted(complete, key=lambda x: -x["arc_sum"])[:10]
    top_cav = sorted(complete, key=lambda x: -x["cav_sum"])[:10]
    top_power = sorted([m for m in complete if m["power"]], key=lambda x: -x["power"])[:10]

    def side_rows(items, key_a, key_b=None):
        rows = []
        for i, m in enumerate(items, 1):
            if key_b:
                rows.append(
                    f"<tr><td>{i}</td><td>{esc(m['name'])}</td>"
                    f"<td class=num>{fmt(m[key_a],1)}</td><td class=num>{fmt(m[key_b],1)}</td></tr>"
                )
            else:
                rows.append(
                    f"<tr><td>{i}</td><td>{esc(m['name'])}</td>"
                    f"<td class=num>{fmt_power(m[key_a]) if key_a=='power' else fmt(m[key_a],1)}</td></tr>"
                )
        return "".join(rows)

    inf_rows = "".join(
        f"<tr><td>{i}</td><td>{esc(m['name'])}</td><td class=num>{fmt(m['inf_sum'],1)}</td>"
        f"<td class=num>{fmt(m['inf_hp'],1)}</td><td class=num>{fmt(m['inf_def'],1)}</td></tr>"
        for i, m in enumerate(top_inf, 1)
    )
    arc_rows = "".join(
        f"<tr><td>{i}</td><td>{esc(m['name'])}</td><td class=num>{fmt(m['arc_sum'],1)}</td>"
        f"<td class=num>{fmt(m['arc_atk'],1)}</td><td class=num>{fmt(m['arc_dmg'],1)}</td></tr>"
        for i, m in enumerate(top_arc, 1)
    )
    cav_rows = "".join(
        f"<tr><td>{i}</td><td>{esc(m['name'])}</td><td class=num>{fmt(m['cav_sum'],1)}</td>"
        f"<td class=num>{fmt(m['cav_atk'],1)}</td><td class=num>{fmt(m['cav_dmg'],1)}</td></tr>"
        for i, m in enumerate(top_cav, 1)
    )

    partial_rows = "".join(
        f"<tr><td class=num>{fmt(m['seq'],0)}</td><td>{esc(m['name'])}</td>"
        f"<td class=num>{fmt_power(m['power'])}</td>"
        f"<td>{esc('；'.join(m.get('issues') or []))}</td></tr>"
        for m in partial
    )

    blank_names = "、".join(esc(m["name"]) for m in blank) or "-"

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{esc(meta["title"])}</title>
<style>
  :root {{
    --bg:#0f1115; --panel:#171a21; --border:#2a2f3a; --text:#e8eaed;
    --muted:#9aa3b2; --dim:#6b7380; --accent:#5b8def; --row2:#151820;
  }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; font-family:"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif; background:var(--bg); color:var(--text); line-height:1.5; }}
  .wrap {{ max-width:1120px; margin:0 auto; padding:28px 20px 60px; }}
  h1 {{ font-size:26px; font-weight:650; margin:0 0 6px; }}
  h2 {{ font-size:18px; font-weight:600; margin:28px 0 10px; }}
  h3 {{ font-size:15px; font-weight:600; margin:18px 0 8px; color:var(--muted); }}
  .sub {{ color:var(--muted); font-size:13px; margin-bottom:22px; }}
  .stats {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:18px; }}
  .stat {{ background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:14px 16px; }}
  .stat .v {{ font-size:22px; font-weight:650; }}
  .stat .l {{ font-size:12px; color:var(--muted); margin-top:4px; }}
  .callout {{ background:#152033; border:1px solid #2a4060; border-radius:8px; padding:12px 14px; font-size:13px; color:#c5d4ea; margin-bottom:22px; }}
  .callout strong {{ color:var(--text); }}
  .grid2 {{ display:grid; grid-template-columns:1.4fr 1fr; gap:16px; }}
  .panel {{ background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:14px 16px; }}
  .bar-row {{ display:grid; grid-template-columns:120px 1fr 48px; gap:8px; align-items:center; margin:6px 0; font-size:12px; }}
  .bar-label {{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:var(--muted); }}
  .bar-track {{ height:10px; background:#232833; border-radius:4px; overflow:hidden; }}
  .bar-fill {{ height:100%; background:var(--accent); border-radius:4px; }}
  .bar-val {{ text-align:right; font-variant-numeric:tabular-nums; }}
  .tools {{ display:flex; flex-wrap:wrap; gap:8px; margin:10px 0 12px; align-items:center; }}
  .tools label {{ font-size:12px; color:var(--muted); }}
  select, input[type="search"] {{ background:#1c2028; color:var(--text); border:1px solid var(--border); border-radius:6px; padding:6px 10px; font-size:13px; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; background:var(--panel); border:1px solid var(--border); border-radius:8px; overflow:hidden; }}
  th, td {{ padding:8px 10px; border-bottom:1px solid var(--border); text-align:left; }}
  th {{ background:#1a1e27; color:var(--muted); font-weight:600; font-size:12px; position:sticky; top:0; z-index:1; }}
  tr:nth-child(even) td {{ background:var(--row2); }}
  tr.top td {{ color:#d8e6ff; }}
  td.num, th.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
  td.score {{ font-weight:650; color:var(--accent); }}
  .split3 {{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:14px; }}
  .table-scroll {{ max-height:560px; overflow:auto; border-radius:8px; }}
  .muted {{ color:var(--muted); font-size:13px; }}
  .dim {{ color:var(--dim); font-size:12px; margin-top:8px; }}
  @media (max-width:900px) {{ .stats,.grid2,.split3 {{ grid-template-columns:1fr; }} }}
  @media print {{
    body {{ background:#fff; color:#111; }}
    .panel,.stat,.callout,table {{ background:#fff!important; border-color:#ccc!important; color:#111!important; }}
    .table-scroll {{ max-height:none; overflow:visible; }}
  }}
</style>
</head>
<body>
<div class="wrap">
  <h1>{esc(meta["title"])}</h1>
  <div class="sub">数据源：{esc(meta["source"])} · 六维齐全 {meta["complete"]} 人 / 全表 {meta["total"]} 人 · 评分不含实力与集结加成</div>

  <div class="stats">
    <div class="stat"><div class="v">{meta["complete"]}</div><div class="l">六维齐全可排名</div></div>
    <div class="stat"><div class="v">{meta["incomplete"]}</div><div class="l">属性缺失未入榜</div></div>
    <div class="stat"><div class="v">{fmt(meta["avgScore"])}</div><div class="l">完整组平均属性分</div></div>
    <div class="stat"><div class="v">{fmt(meta["avgSix"],0)}</div><div class="l">完整组平均六维</div></div>
  </div>

  <div class="callout">
    <strong>评分公式</strong><br/>
    {esc(meta["formula"])}<br/>
    {esc(meta["six_def"])}
  </div>

  <div class="grid2">
    <div class="panel">
      <h2 style="margin-top:0">属性分 Top10</h2>
      {"".join(bars)}
    </div>
    <div class="panel">
      <h2 style="margin-top:0">实力参考 Top10（不进分）</h2>
      <table>
        <thead><tr><th>#</th><th>成员</th><th class="num">实力</th><th class="num">六维</th></tr></thead>
        <tbody>
          {"".join(f"<tr><td>{i}</td><td>{esc(m['name'])}</td><td class=num>{fmt_power(m['power'])}</td><td class=num>{fmt(m['six'],0)}</td></tr>" for i,m in enumerate(top_power,1))}
        </tbody>
      </table>
    </div>
  </div>

  <h2>完整六维排名</h2>
  <div class="tools">
    <label>排序
      <select id="sort">
        <option value="score">属性分</option>
        <option value="six">六维综合</option>
        <option value="inf">步兵合计</option>
        <option value="arc">弓兵合计</option>
        <option value="cav">骑兵合计</option>
        <option value="power">实力(参考)</option>
        <option value="rally">集结(参考)</option>
      </select>
    </label>
    <input type="search" id="q" placeholder="搜索游戏名…" style="min-width:180px" />
  </div>
  <div class="table-scroll">
    <table id="rankTable">
      <thead>
        <tr>
          <th class="num">#</th><th>游戏名</th><th class="num">实力</th><th class="num">集结%</th>
          <th class="num">步兵</th><th class="num">弓兵</th><th class="num">骑兵</th>
          <th class="num">六维</th><th class="num">属性分</th>
        </tr>
      </thead>
      <tbody>{"".join(rank_rows)}</tbody>
    </table>
  </div>

  <div class="split3" style="margin-top:24px">
    <div>
      <h2>步兵 Top10</h2>
      <table><thead><tr><th>#</th><th>成员</th><th class="num">合计</th><th class="num">步生</th><th class="num">步防</th></tr></thead>
      <tbody>{inf_rows}</tbody></table>
    </div>
    <div>
      <h2>弓兵 Top10</h2>
      <table><thead><tr><th>#</th><th>成员</th><th class="num">合计</th><th class="num">弓攻</th><th class="num">弓破</th></tr></thead>
      <tbody>{arc_rows}</tbody></table>
    </div>
    <div>
      <h2>骑兵 Top10</h2>
      <table><thead><tr><th>#</th><th>成员</th><th class="num">合计</th><th class="num">骑攻</th><th class="num">骑破</th></tr></thead>
      <tbody>{cav_rows}</tbody></table>
    </div>
  </div>

  <h2>未入榜</h2>
  <h3>六维缺项（{len(partial)}）</h3>
  <table>
    <thead><tr><th class="num">序号</th><th>游戏名</th><th class="num">实力</th><th>问题</th></tr></thead>
    <tbody>{partial_rows if partial_rows else "<tr><td colspan=4>无</td></tr>"}</tbody>
  </table>
  <h3>几乎无属性（{len(blank)}）</h3>
  <p class="muted">{blank_names}</p>
  <p class="dim">微信打开本页即可筛选排序。更新：替换 data.xlsx 后运行更新脚本或推送仓库。</p>
</div>
<script>
(() => {{
  const sortSel = document.getElementById("sort");
  const q = document.getElementById("q");
  const tbody = document.querySelector("#rankTable tbody");
  const rows = Array.from(tbody.querySelectorAll("tr")).map((tr, i) => ({{
    el: tr,
    name: tr.children[1].textContent,
    power: parseFloat(tr.dataset.power) || -1,
    rally: parseFloat(tr.dataset.rally) || -1,
    six: parseFloat(tr.dataset.six) || -1,
    inf: parseFloat(tr.dataset.inf) || -1,
    arc: parseFloat(tr.dataset.arc) || -1,
    cav: parseFloat(tr.dataset.cav) || -1,
    score: parseFloat(tr.dataset.score) || -1,
    orig: i,
  }}));
  function apply() {{
    const key = sortSel.value;
    const term = q.value.trim().toLowerCase();
    let list = rows.filter(r => !term || r.name.toLowerCase().includes(term));
    list.sort((a,b) => (b[key] !== a[key] ? b[key]-a[key] : a.orig-b.orig));
    tbody.innerHTML = "";
    list.forEach((r,i) => {{ r.el.children[0].textContent = String(i+1); tbody.appendChild(r.el); }});
  }}
  sortSel.addEventListener("change", apply);
  q.addEventListener("input", apply);
}})();
</script>
</body>
</html>
"""


def main():
    if not DATA_XLSX.exists():
        raise SystemExit(f"找不到 {DATA_XLSX}")
    members, _ = load_members(DATA_XLSX)
    complete, blank, partial, meta = score_members(members)
    OUT_HTML.write_text(render_html(complete, blank, partial, meta), encoding="utf-8")
    print(f"OK ranked={len(complete)} blankish={len(blank)} partial={len(partial)}")
    print("TOP10:")
    for m in complete[:10]:
        print(m["rank"], m["name"], m["score"], "six", m["six"])


if __name__ == "__main__":
    main()