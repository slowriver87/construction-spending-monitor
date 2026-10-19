"""Tests for scripts/update.py. Run: python -m pytest tests"""
import copy
import datetime as dt
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import update as u  # noqa: E402

PAGE = (ROOT / "index.html").read_text(encoding="utf-8")
FIX = Path(__file__).parent / "fixtures"


@pytest.fixture
def page():
    data, notes, hist_line = u.load_page(PAGE)
    return copy.deepcopy(data), copy.deepcopy(notes), u.parse_block(hist_line)


def sheet_rows():
    """Every cell of the real release.xlsx sheet 'Table1' (Aug 2026 release)."""
    return json.loads((FIX / "release_xlsx_table1.json").read_text(encoding="utf-8"))["rows"]


def table1_aug():
    return u.table1_from_sheet(sheet_rows())


# --- page -------------------------------------------------------------------

def test_blocks_round_trip_byte_identical(page):
    data, notes, _ = page
    assert u.render_page(PAGE, data, notes) == PAGE


def test_render_changes_only_data_lines(page):
    data, notes, _ = page
    data["latest"]["status"] = "revised"
    new = u.render_page(PAGE, data, notes)
    changed = [i for i, (a, b) in enumerate(zip(PAGE.splitlines(), new.splitlines())) if a != b]
    assert changed == [u.block_lines(PAGE.splitlines())["DATA"]]
    u.verify_page(PAGE, new)


def test_check5_rejects_change_outside_blocks():
    bad = PAGE.replace("For education only. Not financial advice.", "For education only.")
    with pytest.raises(u.CheckFailed, match="check 5"):
        u.verify_page(PAGE, bad)


def test_check6_rejects_hist_change():
    lines = PAGE.splitlines(keepends=True)
    i = u.block_lines(lines)["HIST"]
    lines[i] = lines[i].replace("2006–10 bust", "2006-10 bust")
    with pytest.raises(u.CheckFailed, match="check 6"):
        u.verify_page(PAGE, "".join(lines))


def test_check8_rejects_broken_script():
    with pytest.raises(u.CheckFailed, match="does not parse"):
        u.node_check("const a = {;")


def test_education_note_directly_under_header():
    after_header = PAGE.split("</header>", 1)[1].lstrip()
    assert after_header.startswith('<p class="disclaimer"><b>For education only. Not financial advice.</b>')


# --- Table 1 mapping ----------------------------------------------------------

def test_table1_mapping_matches_page(page):
    data, _, _ = page
    t1 = table1_aug()
    assert t1.month == "2026-08" and t1.yago == "2025-08"
    stored = {r[0]: r[1:] for r in data["months"]}
    for i, m in enumerate(t1.cols[:5]):
        assert [t1.values[f][i] for f in u.FIELDS] == stored[m]
    by_name = {s["name"]: s for s in data["latest"]["sectors"]}
    keys = {"Private nonresidential": "nonres", "Manufacturing": "mfg", "Power": "power",
            "Office (incl. data centers)": "office", "Commercial": "commercial", "Public": "pub",
            "Highway and street": "highway", "Educational": "educational"}
    for name, k in keys.items():
        s = by_name[name]
        assert (t1.get(k, "2026-08"), t1.get(k, "2026-07"), t1.get(k, "2025-08")) == (s["cur"], s["prev"], s["yago"])


def test_table1_uses_block_not_first_match():
    # "Residential", "Office", "Power", "Educational" and others appear in more than one block.
    t1 = table1_aug()
    assert t1.get("res", "2026-08") == 882270          # private, not 894,597 (Total block)
    assert t1.get("office", "2026-08") == 134119       # private, not 150,669 (Total) or 16,550 (Public)
    assert t1.get("educational", "2026-08") == 113073  # public, not 137,469 (Total) or 24,396 (Private)
    assert t1.get("mfg", "2026-08") == 168209          # private, not 170,701 (Total)


def test_footnote_digits_and_padding():
    assert u.clean_label("Residential2") == "Residential"
    assert u.clean_label("Total Private Construction1") == "Total Private Construction"
    assert u.clean_label("  Total Public Construction3 ") == "Total Public Construction"
    assert u.clean_label("    Manufacturing ") == "Manufacturing"
    assert u.clean_label("Power2") == "Power2"  # only the four known footnoted labels


def test_table1_xlsx_and_pdf_agree():
    x = table1_aug()
    p = u.table1_from_pdf_text((FIX / "release_pdf_table1.txt").read_text(encoding="utf-8"))
    assert (x.cols, x.values) == (p.cols, p.values)
    assert x.source == "release.xlsx" and p.source == "release.pdf"


def test_table1_xlsx_bytes_round_trip(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Table1"
    for r in sheet_rows():
        ws.append(r)
    wb.save(tmp_path / "release.xlsx")
    t1 = u.read_table1_xlsx((tmp_path / "release.xlsx").read_bytes())
    assert t1.values == table1_aug().values


def test_table1_missing_row_fails():
    rows = [r for r in sheet_rows() if not (isinstance(r[0], str) and r[0].strip() == "Commercial")]
    with pytest.raises(u.SourceError, match="commercial"):
        u.table1_from_sheet(rows)


def test_table1_duplicate_label_in_block_fails():
    rows = sheet_rows()
    i = max(i for i, r in enumerate(rows) if isinstance(r[0], str) and r[0].strip() == "Educational") + 1
    rows.insert(i, ["    Educational", 1, 1, 1, 1, 1, 1, 0.0, 0.0])  # second Educational in the Public block
    with pytest.raises(u.SourceError, match="twice"):
        u.table1_from_sheet(rows)


def test_broken_xlsx_falls_back_to_pdf(monkeypatch):
    pdf_text = (FIX / "release_pdf_table1.txt").read_text(encoding="utf-8")
    monkeypatch.setattr(u, "fetch", lambda url, **kw: url.encode())
    monkeypatch.setattr(u, "read_table1_xlsx", lambda b: u.table1_from_sheet(sheet_rows()[:20]))  # rows missing
    monkeypatch.setattr(u, "read_table1_pdf", lambda b: u.table1_from_pdf_text(pdf_text))
    said = []
    t1 = u.get_table1(said.append)
    assert t1.source == "release.pdf" and "failed" in said[0]


def test_table1_bad_columns_fail():
    rows = sheet_rows()
    h = next(i for i, r in enumerate(rows) if r[0] and r[0].strip() == "Type of Construction")
    rows[h] = list(rows[h])
    rows[h][3] = "May\n2026r"  # Jun column relabelled: months no longer consecutive
    with pytest.raises(u.SourceError, match="column months"):
        u.table1_from_sheet(rows)


def test_private_tie_check():
    t1 = table1_aug()
    u.check_table1(t1)
    t1.values["private_total"][0] += 6
    with pytest.raises(u.CheckFailed, match="check 2"):
        u.check_table1(t1)


# --- then vs. now and summary ---------------------------------------------------

def test_then_vs_now_matches_hand_written_august_note(page):
    data, _, hist = page
    tv = u.then_vs_now(data["months"], hist)
    assert (tv["k"], tv["peak_month"], tv["bust_peak_month"]) == (18, "2025-02", "2006-02")
    assert f"{tv['index']:.1f}" == "97.2" and f"{tv['bust_value']:.1f}" == "96.1"


def test_summary_from_current_page(page):
    data, _, _ = page
    s = u.build_summary(data)
    assert s == json.loads((ROOT / "data" / "summary.json").read_text(encoding="utf-8"))
    assert (s["mom"], s["yoy"], s["pct_below_peak"]) == (0.9, -1.7, 2.8)


# --- building a new month (synthetic September; numbers are made up) ---------------

def synthetic_sep(data, aug_total_bump=0):
    """Table 1 for a made-up Sep 2026 release, built from the stored months."""
    st = {r[0]: dict(zip(u.FIELDS, r[1:])) for r in data["months"]}
    sep = {"total": 2210000, "res": 885000, "nonres": 776000, "pub": 549000, "mfg": 167000}
    aug = dict(st["2026-08"])
    aug["res"] += aug_total_bump
    aug["total"] += aug_total_bump
    cols = ["2026-09", "2026-08", "2026-07", "2026-06", "2026-05", "2025-09"]
    rows_by_month = [sep, aug, st["2026-07"], st["2026-06"], st["2026-05"], st["2025-09"]]
    values = {f: [r[f] for r in rows_by_month] for f in u.FIELDS}
    values["private_total"] = [r["res"] + r["nonres"] for r in rows_by_month]
    for k, cur, prev, yago in (("power", 167000, 165900, 150000), ("office", 136000, 134119, 105000),
                               ("commercial", 116000, 116448, 122000), ("highway", 151000, 150610, 145000),
                               ("educational", 113500, 113073, 112000)):
        values[k] = [cur, prev, None, None, None, yago]
    return u.Table1(cols, values, "release.xlsx")


def text_for(t1, sig=(True, False, True, False)):
    m, p = t1.cols[0], t1.cols[1]
    ch = lambda k, s: u.Change(round(t1.get(k, m) / 1000, 1), round(u.pct(t1.get(k, m), t1.get(k, p)), 1), s)
    yoy = u.pct(t1.get("total", m), t1.get("total", t1.yago))
    return u.ReleaseText(month=m, released=dt.date(2026, 11, 2), total=ch("total", sig[0]),
                         yoy=u.Change(round(t1.get("total", t1.yago) / 1000, 1), round(yoy, 1), True),
                         ytd_months=9, ytd_cur_b=1650.2, ytd_prior_b=1700.9,
                         res=ch("res", sig[1]), nonres=ch("nonres", sig[2]), pub=ch("pub", sig[3]))


SCHEDULE = [("2026-09", dt.date(2026, 11, 2)), ("2026-10", dt.date(2026, 12, 1))]


def test_build_new_month(page):
    data, notes, hist = page
    t1 = synthetic_sep(data, aug_total_bump=3000)
    res = u.build_update(data, notes, hist, t1, text_for(t1), None, SCHEDULE, None)
    L = res.data["latest"]
    assert res.data["months"][-1] == ["2026-09", 2210000, 885000, 776000, 549000, 167000]
    assert res.added == ["2026-09"]
    assert [str(r) for r in res.revisions] == ["2026-08 total: 2,203,129 → 2,206,129 (+0.1%)",
                                               "2026-08 res: 882,270 → 885,270 (+0.3%)"]
    assert L["released"] == "Nov 2, 2026" and L["next"] == "Dec 1, 2026 (Oct data)"
    assert L["moe"] == {"total": True, "res": False, "nonres": True, "pub": False}
    assert L["ytd"] == {"cur": 1650.2, "prior": 1700.9, "label": "Jan–Sep 2026 vs Jan–Sep 2025"}
    assert L["sectorNote"] == "Selected categories from the September release. Census counts data centers under Office."
    assert [s["name"] for s in L["sectors"]] == ["Private nonresidential", "Manufacturing", "Power",
                                                 "Office (incl. data centers)", "Commercial", "Public",
                                                 "Highway and street", "Educational"]
    # hand-written notes kept unchanged, one automatic note added
    assert res.notes[:2] == notes and res.notes[2]["id"] == "2026-09" and res.notes[2]["auto"] is True
    note = res.notes[2]
    assert note["headline"] == "Total $2,210.0B, +0.2% m/m, −1.6% y/y"
    assert note["points"][0].endswith("Nominal dollars.")
    assert note["points"][1] == ("Private residential $885.0B, 0.0% m/m (not significant); private nonresidential "
                                 "$776.0B, +0.4% m/m; public $549.0B, +0.2% m/m (not significant).")
    assert note["points"][3].startswith("Revisions to total: Aug 2026 +$3.0B to $2,206.1B; Jul 2026 unrevised")
    assert "19 months after the Feb 2025 peak" in note["points"][4]
    assert note["points"][-1].startswith("Checks: ") and "FRED not reached" in note["points"][-1]
    new_html = u.render_page(PAGE, res.data, res.notes)
    u.verify_page(PAGE, new_html)
    assert u.load_page(new_html)[2] == u.load_page(PAGE)[2]


def test_nothing_new(page):
    data, notes, hist = page
    assert u.build_update(data, notes, hist, table1_aug(), None, None, None, None) is None


def test_text_disagreement_fails(page):
    data, notes, hist = page
    t1 = synthetic_sep(data)
    rt = text_for(t1)
    rt.total.level_b += 0.2
    with pytest.raises(u.CheckFailed, match="check 3"):
        u.build_update(data, notes, hist, t1, rt, None, SCHEDULE, None)


def test_month_mismatch_fails(page):
    data, notes, hist = page
    t1 = synthetic_sep(data)
    rt = text_for(t1)
    rt.month = "2026-08"
    with pytest.raises(u.CheckFailed, match="check 1"):
        u.build_update(data, notes, hist, t1, rt, None, SCHEDULE, None)


def test_tie_failure(page):
    data, notes, hist = page
    t1 = synthetic_sep(data)
    t1.values["total"][0] += 10
    with pytest.raises(u.CheckFailed, match="check 2"):
        u.build_update(data, notes, hist, t1, None, None, SCHEDULE, None)


def test_gap_fails(page):
    data, notes, hist = page
    data["months"] = [r for r in data["months"] if r[0] != "2024-03"]
    t1 = synthetic_sep(data)
    with pytest.raises(u.CheckFailed, match="check 7"):
        u.build_update(data, notes, hist, t1, None, None, SCHEDULE, None)


def test_without_release_text_carries_over(page):
    data, notes, hist = page
    t1 = synthetic_sep(data)
    res = u.build_update(data, notes, hist, t1, None, None, SCHEDULE, None)
    assert res.note is None and res.notes == notes
    assert res.data["latest"]["moe"] == data["latest"]["moe"]
    assert res.data["latest"]["ytd"] == data["latest"]["ytd"]
    assert any("carried over" in line for line in res.report)


def test_time_series_revision_and_big_move(page):
    data, notes, hist = page
    t1 = synthetic_sep(data)
    ts = {r[0]: dict(zip(u.FIELDS, r[1:])) for r in data["months"]}
    ts["2023-05"] = {**ts["2023-05"], "mfg": round(ts["2023-05"]["mfg"] * 1.08)}  # annual-revision style move
    res = u.build_update(data, notes, hist, t1, text_for(t1), ts, SCHEDULE, None)
    assert [(r.month, r.field) for r in res.big_revisions] == [("2023-05", "mfg")]
    msg = u.commit_message(res)
    assert "Moves over 5%" in msg and "2023-05 mfg" in msg
    assert "Other months revised: 1." in res.note["points"][3]


def test_fred_parse_and_compare(page):
    text = ("Title: Total Construction Spending: Total Construction in the United States\n"
            "Series ID: TTLCONS\nUnits: Millions of Dollars\n\nDATE        VALUE\n"
            "2026-07-01 2184474\n2026-08-01 2203129\n")
    assert u.parse_fred(text) == {"2026-07": 2184474, "2026-08": 2203129}
    data, notes, hist = page
    t1 = synthetic_sep(data)
    fred = {f: {"2026-09": t1.get(f, "2026-09")} for f in u.FIELDS}
    res = u.build_update(data, notes, hist, t1, text_for(t1), None, SCHEDULE, fred)
    assert "FRED agreed" in res.note["points"][-1]


# --- release text, schedule, time series (real Aug 2026 files) ------------------

RELEASE_HTML = (FIX / "release_text_aug2026.html").read_text(encoding="utf-8")


def test_release_text():
    rt = u.parse_release_text(RELEASE_HTML)
    assert (rt.month, rt.released) == ("2026-08", dt.date(2026, 10, 1))
    assert rt.total == u.Change(2203.1, 0.9, False)
    assert rt.yoy == u.Change(2242.0, -1.7, True)
    assert (rt.ytd_months, rt.ytd_cur_b, rt.ytd_prior_b) == (8, 1450.4, 1496.6)
    assert rt.res == u.Change(882.3, 1.1, False)
    assert rt.nonres == u.Change(773.0, 1.0, True)
    assert rt.pub == u.Change(547.8, 0.2, False)


def test_release_text_matches_stored_flags(page):
    data, _, _ = page
    rt = u.parse_release_text(RELEASE_HTML)
    assert {"total": rt.total.significant, "res": rt.res.significant, "nonres": rt.nonres.significant,
            "pub": rt.pub.significant} == data["latest"]["moe"]
    assert (rt.ytd_cur_b, rt.ytd_prior_b) == (data["latest"]["ytd"]["cur"], data["latest"]["ytd"]["prior"])
    u.check_against_text(table1_aug(), rt)


def test_release_text_private_sentence_only_in_private_section():
    # Moving the residential sentence into the Public section must not be accepted.
    i = RELEASE_HTML.index("Residential construction was at")
    j = RELEASE_HTML.index("revised July estimate of $872.7 billion.") + len("revised July estimate of $872.7 billion.")
    sentence = RELEASE_HTML[i:j]
    moved = RELEASE_HTML[:i] + RELEASE_HTML[j:]
    k = moved.index("Educational construction was at")
    moved = moved[:k] + sentence + " " + moved[k:]
    with pytest.raises(u.SourceError, match="private residential"):
        u.parse_release_text(moved)


def test_release_text_missing_sentence_fails():
    with pytest.raises(u.SourceError, match="public"):
        u.parse_release_text(RELEASE_HTML.replace("In August, the estimated", "In August the estimated"))


def test_schedule():
    sched = u.parse_schedule((FIX / "schedule_2026.html").read_text(encoding="utf-8"))
    assert sched[:4] == [("2026-09", dt.date(2026, 11, 2)), ("2026-10", dt.date(2026, 12, 1)),
                         ("2026-11", dt.date(2027, 1, 4)), ("2026-12", dt.date(2027, 2, 1))]
    assert sched[4] == ("2027-01", dt.date(2027, 3, 1)) and sched[-1] == ("2027-10", dt.date(2027, 12, 1))
    assert u.next_release(sched, "2026-08") == "Nov 2, 2026 (Sep data)"


def ts_sheets():
    fx = json.loads((FIX / "timeseries.json").read_text(encoding="utf-8"))
    return {"total": fx["totsatime"]["rows"], "private": fx["privsatime"]["rows"], "public": fx["pubsatime"]["rows"]}


def test_timeseries_matches_page(page):
    data, _, _ = page
    ts = u.timeseries_from_rows(ts_sheets())
    assert max(ts) == "2026-08"
    for r in data["months"]:
        assert [ts[r[0]][f] for f in u.FIELDS] == r[1:], r[0]


def test_timeseries_cross_check_fails_on_disagreement():
    sheets = ts_sheets()
    row = next(r for r in sheets["private"] if r[0] == "Jul-26r")
    row[2] += 1  # private residential, July
    with pytest.raises(u.SourceError, match="differ for 2026-07"):
        u.timeseries_from_rows(sheets)


def test_ts_header_and_month():
    assert u.ts_header("Total\n_x000D_Construction1") == "Total Construction"
    assert u.ts_header("Residential (inc. Improvements)2") == "Residential (inc. Improvements)"
    assert (u.ts_month("Aug-26p"), u.ts_month("Jan-93"), u.ts_month("p Preliminary r Revised")) == ("2026-08", "1993-01", None)


def test_full_run_on_real_files_with_august_removed(page):
    """Rebuild August from the real files on a page that stops at July: must reproduce the stored August."""
    data, notes, hist = page
    stored = copy.deepcopy(data)
    data["months"] = data["months"][:-1]
    data["latest"]["month"] = "2026-07"
    notes = [n for n in notes if n["id"] != "2026-08"]
    t1, rt = table1_aug(), u.parse_release_text(RELEASE_HTML)
    ts = u.timeseries_from_rows(ts_sheets())
    sched = u.parse_schedule((FIX / "schedule_2026.html").read_text(encoding="utf-8"))
    res = u.build_update(data, notes, hist, t1, rt, ts, sched, None)
    assert res.data["months"] == stored["months"] and res.revisions == [] and res.added == ["2026-08"]
    for k in ("month", "status", "released", "next", "moe", "ytd", "sectors"):
        assert res.data["latest"][k] == stored["latest"][k], k
    assert res.note["headline"] == "Total $2,203.1B, +0.9% m/m (not significant), −1.7% y/y"


def test_no_personal_information_in_repo():
    import re
    pat = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+|https?://(?:[\w-]+\.)?claude\.ai", re.I)
    for p in ROOT.rglob("*"):
        if p.is_file() and not any(part.startswith(".") and part not in (".github",) for part in p.relative_to(ROOT).parts) \
                and "__pycache__" not in p.parts and p.suffix in {".py", ".md", ".yml", ".json", ".txt", ".html"}:
            hits = [h for h in pat.findall(p.read_text(encoding="utf-8"))]
            assert not hits, f"{p}: {hits}"
