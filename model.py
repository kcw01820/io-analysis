import json
import math
import numpy as np
import openpyxl

WORKBOOKS = {}       # filename -> {"sheets": [...], "sheet_meta": {...}, "raw_rows": {...}}
IO_TABLES = {}       # level -> parsed dict
CLASSIFICATION = {   # parsed from (표)상품분류.xlsx
    "loaded": False,
    "items": [],
    "summary_codes": {}
}
LEONTIEF_CACHE = {}  # (level, model_type) -> dict

LEVEL_LABELS = {
    "major33": "대분류 (33부문)",
    "mid83": "중분류 (83부문)",
    "minor165": "소분류 (165부문)",
    "basic380": "기본부문 (380부문)"
}

NON_SECTOR_ROW_CODES = {
    "9519", "9520", "9590", "9610", "9621", "9622", "9630", "9690", "9790",
    "9801", "9802", "9819", "9812", "9811"
}

FD_CODE_NAMES = {
    "9111": "민간소비지출",
    "9112": "정부소비지출",
    "9121": "민간고정자본형성",
    "9122": "정부고정자본형성",
    "9131": "재고증감",
    "9132": "귀중품순취득",
    "9140": "수출"
}


def clean_str(v):
    if v is None:
        return ""
    return str(v).replace("\u3000", " ").strip()


def to_float(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        if math.isnan(v) or math.isinf(v):
            return 0.0
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "nan", "None"):
        return 0.0
    try:
        val = float(s)
        return 0.0 if (math.isnan(val) or math.isinf(val)) else val
    except ValueError:
        return 0.0


def safe_num(x, digits=None):
    if x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))):
        return 0.0
    val = float(x)
    return round(val, digits) if digits is not None else val


def json_dumps_safe(obj):
    return json.dumps(obj, ensure_ascii=False)


def detect_level(sector_count, filename):
    if sector_count == 33:
        return "major33"
    if sector_count == 83:
        return "mid83"
    if sector_count == 165:
        return "minor165"
    if sector_count == 380:
        return "basic380"
    if "대분류" in filename:
        return "major33"
    if "중분류" in filename:
        return "mid83"
    if "소분류" in filename:
        return "minor165"
    if "기본부문" in filename:
        return "basic380"
    return f"custom_{sector_count}"


def try_parse_classification(rows):
    header_idx = None
    for idx, r in enumerate(rows[:10]):
        joined = " ".join(clean_str(c) for c in r[:8])
        if "코드" in joined and "상품명" in joined:
            header_idx = idx
            break
    if header_idx is None:
        return False

    items = []
    summary_codes = {}
    cur_minor = ("", "")
    cur_major = ("", "")
    cur_div = ("", "")

    for r in rows[header_idx + 1:]:
        vals = [clean_str(c) for c in list(r) + [None] * 8][:8]
        b_code, b_name, mn_code, mn_name, mj_code, mj_name, dv_code, dv_name = vals
        if not b_code:
            continue
        if b_code.startswith("9"):
            summary_codes[b_code] = b_name
            continue
        if dv_code:
            cur_div = (dv_code, dv_name)
        if mj_code:
            cur_major = (mj_code, mj_name)
        if mn_code:
            cur_minor = (mn_code, mn_name)

        items.append({
            "basic_code": b_code,
            "basic_name": b_name,
            "minor_code": cur_minor[0],
            "minor_name": cur_minor[1],
            "major_code": cur_major[0],
            "major_name": cur_major[1],
            "div_code": cur_div[0],
            "div_name": cur_div[1]
        })

    if items:
        CLASSIFICATION["loaded"] = True
        CLASSIFICATION["items"] = items
        CLASSIFICATION["summary_codes"] = summary_codes
        return True
    return False


def try_parse_io_table(filename, sheet_name, rows):
    code_row_idx = None
    for idx, r in enumerate(rows[:15]):
        row_strs = [clean_str(c) for c in r]
        if "9090" in row_strs and ("9310" in row_strs or "9390" in row_strs):
            code_row_idx = idx
            break
    if code_row_idx is None or code_row_idx + 2 >= len(rows):
        return None

    code_row = [clean_str(c) for c in rows[code_row_idx]]
    name_row = [clean_str(c) for c in rows[code_row_idx + 1]]

    if "9090" not in code_row:
        return None
    col_9090 = code_row.index("9090")

    sector_col_indices = []
    sectors = []
    for c_idx in range(2, col_9090):
        sc = code_row[c_idx]
        sn = name_row[c_idx] if c_idx < len(name_row) else ""
        if sc:
            sector_col_indices.append(c_idx)
            sectors.append({"code": sc, "name": sn})

    n = len(sectors)
    if n == 0:
        return None

    col_map = {code_row[i]: i for i in range(col_9090, len(code_row)) if code_row[i]}

    Z = np.zeros((n, n), dtype=np.float64)
    sector_row_map = {}
    va_rows = {}
    va_totals = {}

    row_cursor = 0
    in_va = False
    for r in rows[code_row_idx + 2:]:
        r_code = clean_str(r[0]) if len(r) > 0 else ""
        r_name = clean_str(r[1]) if len(r) > 1 else ""
        if not r_code:
            continue
        if r_code in NON_SECTOR_ROW_CODES:
            in_va = True
            arr = np.array([to_float(r[c] if c < len(r) else 0.0) for c in sector_col_indices], dtype=np.float64)
            va_rows[r_code] = arr
            va_totals[r_code] = to_float(r[col_9090] if col_9090 < len(r) else 0.0)
            continue

        if not in_va and row_cursor < n:
            for j, c_idx in enumerate(sector_col_indices):
                Z[row_cursor, j] = to_float(r[c_idx] if c_idx < len(r) else 0.0)
            sector_row_map[row_cursor] = r
            row_cursor += 1

    if row_cursor != n:
        return None

    def extract_col_vec(code):
        c_idx = col_map.get(code)
        if c_idx is None:
            return np.zeros(n, dtype=np.float64)
        return np.array([
            to_float(sector_row_map[i][c_idx] if c_idx < len(sector_row_map[i]) else 0.0)
            for i in range(n)
        ], dtype=np.float64)

    fd_cols = {c: extract_col_vec(c) for c in ["9111", "9112", "9121", "9122", "9131", "9132", "9140", "9190"]}
    supply_cols = {c: extract_col_vec(c) for c in ["9090", "9290", "9310", "9321", "9322", "9329", "9331", "9332", "9390"]}

    meta_lines = [clean_str(rows[i][0]) for i in range(min(4, code_row_idx)) if rows[i] and clean_str(rows[i][0])]
    unit = "백만원"
    for ml in meta_lines:
        if "단위" in ml:
            unit = ml.split(":")[-1].strip() if ":" in ml else ml

    level = detect_level(n, filename)
    return {
        "level": level,
        "level_label": LEVEL_LABELS.get(level, f"기타 ({n}부문)"),
        "filename": filename,
        "sheet_name": sheet_name,
        "title": meta_lines[0] if len(meta_lines) > 0 else filename,
        "year": meta_lines[1] if len(meta_lines) > 1 else "2023년",
        "price_basis": meta_lines[2] if len(meta_lines) > 2 else "생산자가격",
        "unit": unit,
        "n": n,
        "sectors": sectors,
        "Z": Z,
        "fd_cols": fd_cols,
        "supply_cols": supply_cols,
        "va_rows": va_rows,
        "va_totals": va_totals
    }


def load_excel_file(filename, path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    wb_info = {
        "sheets": sheet_names,
        "raw_rows": {},
        "detected_type": "reference"
    }

    parsed_level = None
    for sname in sheet_names:
        ws = wb[sname]
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        wb_info["raw_rows"][sname] = rows

        # 1) 상품분류 시트 판별
        if not CLASSIFICATION["loaded"] and ("상품분류" in filename or any("기본부문" in clean_str(c) for r in rows[:5] for c in r[:8])):
            if try_parse_classification(rows):
                wb_info["detected_type"] = "classification"

        # 2) 투입산출표 시트 판별
        io_parsed = try_parse_io_table(filename, sname, rows)
        if io_parsed is not None:
            lvl = io_parsed["level"]
            IO_TABLES[lvl] = io_parsed
            parsed_level = lvl
            wb_info["detected_type"] = f"io_{lvl}"
            # 캐시 무효화
            for k in list(LEONTIEF_CACHE.keys()):
                if k[0] == lvl:
                    del LEONTIEF_CACHE[k]

    wb.close()
    WORKBOOKS[filename] = wb_info
    return json_dumps_safe({
        "filename": filename,
        "sheets": sheet_names,
        "detected_type": wb_info["detected_type"],
        "parsed_level": parsed_level
    })


def get_state_summary():
    levels = []
    for lvl in ["major33", "mid83", "minor165", "basic380"]:
        if lvl in IO_TABLES:
            t = IO_TABLES[lvl]
            levels.append({
                "level": lvl,
                "label": t["level_label"],
                "n": t["n"],
                "filename": t["filename"]
            })
    for lvl, t in IO_TABLES.items():
        if lvl not in ("major33", "mid83", "minor165", "basic380"):
            levels.append({"level": lvl, "label": t["level_label"], "n": t["n"], "filename": t["filename"]})

    return json_dumps_safe({
        "files": {k: {"sheets": v["sheets"], "type": v["detected_type"]} for k, v in WORKBOOKS.items()},
        "levels": levels,
        "classification_loaded": CLASSIFICATION["loaded"],
        "classification_count": len(CLASSIFICATION["items"])
    })


def get_sheet_preview(arg_json):
    req = json.loads(arg_json)
    fname = req.get("filename", "")
    sname = req.get("sheet_name", "")
    query = req.get("query", "").strip().lower()
    limit = int(req.get("limit", 80))

    if fname not in WORKBOOKS or sname not in WORKBOOKS[fname]["raw_rows"]:
        return json_dumps_safe({"meta": "시트를 찾을 수 없습니다.", "rows": []})

    rows = WORKBOOKS[fname]["raw_rows"][sname]
    total_rows = len(rows)
    max_cols = max((len(r) for r in rows[:20]), default=0)

    # 너무 넓은 열(380부문 등)은 앞 25개 열 위주로 보여주고 검색 시 매칭 열 정보 포함
    display_cols = min(max_cols, 18)
    filtered = []
    for idx, r in enumerate(rows):
        if query:
            row_text = " ".join(clean_str(c) for c in r if c is not None).lower()
            if query not in row_text:
                continue
        row_cells = []
        for c_i in range(display_cols):
            val = r[c_i] if c_i < len(r) else ""
            if isinstance(val, float) and val.is_integer():
                row_cells.append(f"{int(val):,}")
            elif isinstance(val, int):
                row_cells.append(f"{val:,}")
            else:
                row_cells.append(clean_str(val))
        filtered.append({"row_no": idx + 1, "cells": row_cells})
        if len(filtered) >= limit:
            break

    return json_dumps_safe({
        "meta": f"전체 {total_rows}행 × {max_cols}열 (표시: 상위 {display_cols}열, 최대 {limit}행)",
        "rows": filtered
    })


def search_classification(arg_json):
    req = json.loads(arg_json)
    query = req.get("query", "").strip().lower()
    limit = int(req.get("limit", 150))

    if not CLASSIFICATION["loaded"]:
        return json_dumps_safe({"total": 0, "items": []})

    matches = []
    for item in CLASSIFICATION["items"]:
        if query:
            text = " ".join([
                item["div_code"], item["div_name"],
                item["major_code"], item["major_name"],
                item["minor_code"], item["minor_name"],
                item["basic_code"], item["basic_name"]
            ]).lower()
            if query not in text:
                continue
        matches.append(item)
        if len(matches) >= limit:
            break

    return json_dumps_safe({
        "total": len(CLASSIFICATION["items"]),
        "matched": len(matches),
        "items": matches
    })


def compute_leontief(level, model_type="domestic_est"):
    key = (level, model_type)
    if key in LEONTIEF_CACHE:
        return LEONTIEF_CACHE[key]

    t = IO_TABLES[level]
    n = t["n"]
    Z = t["Z"]
    X = t["supply_cols"]["9310"]
    M = t["supply_cols"]["9329"]
    E = t["fd_cols"]["9140"]
    TotalD = t["supply_cols"]["9290"]
    VA = t["va_rows"].get("9690", np.zeros(n))
    Comp = t["va_rows"].get("9610", np.zeros(n))

    # 기술투입계수 행렬 A (총산출 0인 부문은 열 전체 0 처리)
    A = np.divide(Z, X[None, :], out=np.zeros_like(Z), where=X[None, :] > 0)

    # 국내수요 대비 수입비율 m_i = M_i / (총수요_i - 수출_i)
    dom_demand = TotalD - E
    m = np.divide(M, dom_demand, out=np.zeros(n), where=dom_demand > 0)
    m = np.clip(m, 0.0, 1.0)

    if model_type == "domestic_est":
        Ad = (1.0 - m)[:, None] * A
        Am = m[:, None] * A
    else:
        Ad = A
        Am = np.zeros_like(A)

    I = np.eye(n, dtype=np.float64)
    L = np.linalg.solve(I - Ad, I)

    # 부가가치율 및 피용자보수율
    v = np.divide(VA, X, out=np.zeros(n), where=X > 0)
    w = np.divide(Comp, X, out=np.zeros(n), where=X > 0)

    prod_mult = L.sum(axis=0)                  # 생산유발계수 (열합)
    va_mult = (v[:, None] * L).sum(axis=0)     # 부가가치유발계수
    wage_mult = (w[:, None] * L).sum(axis=0)   # 피용자보수유발계수
    imp_mult = (Am @ L).sum(axis=0)            # 간접 수입유발계수 (국내생산 1단위당 유발 수입)

    # 영향력계수(후방연쇄) 및 감응도계수(전방연쇄)
    col_sums = L.sum(axis=0)
    row_sums = L.sum(axis=1)
    mean_all = float(col_sums.mean()) if n > 0 else 1.0
    if mean_all == 0:
        mean_all = 1.0
    backward = col_sums / mean_all
    forward = row_sums / mean_all

    res = {
        "A": A, "Ad": Ad, "Am": Am, "L": L,
        "m": m, "v": v, "w": w,
        "prod_mult": prod_mult,
        "va_mult": va_mult,
        "wage_mult": wage_mult,
        "imp_mult": imp_mult,
        "backward": backward,
        "forward": forward
    }
    LEONTIEF_CACHE[key] = res
    return res


def get_structure_analysis(arg_json):
    req = json.loads(arg_json)
    level = req.get("level", "")
    if level not in IO_TABLES:
        return json_dumps_safe({"error": "선택한 분류의 투입산출표가 로드되지 않았습니다."})

    t = IO_TABLES[level]
    n = t["n"]
    X = t["supply_cols"]["9310"]
    M = t["supply_cols"]["9329"]
    TotalS = t["supply_cols"]["9390"]
    IntD = t["supply_cols"]["9090"]
    FD = t["fd_cols"]["9190"]
    E = t["fd_cols"]["9140"]
    VA = t["va_rows"].get("9690", np.zeros(n))
    IntIn = t["va_rows"].get("9590", np.zeros(n))

    sum_X = float(X.sum())
    sum_M = float(M.sum())
    sum_TotalS = float(TotalS.sum())
    sum_VA = float(VA.sum())
    sum_IntIn = float(IntIn.sum())
    sum_E = float(E.sum())
    sum_FD = float(FD.sum())

    # 최종수요 및 부가가치 항목별 합계
    fd_breakdown = []
    for c in ["9111", "9112", "9121", "9122", "9131", "9132", "9140"]:
        val = float(t["fd_cols"][c].sum())
        fd_breakdown.append({
            "code": c,
            "name": FD_CODE_NAMES[c],
            "value": safe_num(val, 0),
            "share": safe_num(val / sum_FD * 100 if sum_FD else 0.0, 2)
        })

    va_names = {"9610": "피용자보수", "9621": "영업잉여", "9622": "고정자본소모", "9630": "생산세(보조금공제)"}
    va_breakdown = []
    for c in ["9610", "9621", "9622", "9630"]:
        val = float(t["va_rows"].get(c, np.zeros(n)).sum())
        va_breakdown.append({
            "code": c,
            "name": va_names[c],
            "value": safe_num(val, 0),
            "share": safe_num(val / sum_VA * 100 if sum_VA else 0.0, 2)
        })

    rows = []
    for i, sec in enumerate(t["sectors"]):
        x_i = float(X[i])
        va_i = float(VA[i])
        ii_i = float(IntIn[i])
        e_i = float(E[i])
        m_i = float(M[i])
        ts_i = float(TotalS[i])
        rows.append({
            "code": sec["code"],
            "name": sec["name"],
            "output": safe_num(x_i, 0),
            "output_share": safe_num(x_i / sum_X * 100 if sum_X else 0.0, 2),
            "va": safe_num(va_i, 0),
            "va_rate": safe_num(va_i / x_i * 100 if x_i > 0 else 0.0, 2),
            "int_input_rate": safe_num(ii_i / x_i * 100 if x_i > 0 else 0.0, 2),
            "export": safe_num(e_i, 0),
            "export_rate": safe_num(e_i / x_i * 100 if x_i > 0 else 0.0, 2),
            "import": safe_num(m_i, 0),
            "import_rate": safe_num(m_i / ts_i * 100 if ts_i > 0 else 0.0, 2)
        })

    return json_dumps_safe({
        "meta": {
            "level_label": t["level_label"],
            "year": t["year"],
            "price_basis": t["price_basis"],
            "unit": t["unit"]
        },
        "kpi": {
            "total_supply": safe_num(sum_TotalS, 0),
            "total_output": safe_num(sum_X, 0),
            "total_va": safe_num(sum_VA, 0),
            "total_import": safe_num(sum_M, 0),
            "total_export": safe_num(sum_E, 0),
            "int_input_rate": safe_num(sum_IntIn / sum_X * 100 if sum_X else 0.0, 2),
            "va_rate": safe_num(sum_VA / sum_X * 100 if sum_X else 0.0, 2),
            "trade_ratio": safe_num((sum_E + sum_M) / sum_TotalS * 100 if sum_TotalS else 0.0, 2)
        },
        "fd_breakdown": fd_breakdown,
        "va_breakdown": va_breakdown,
        "rows": rows
    })


def get_sector_detail(arg_json):
    req = json.loads(arg_json)
    level = req.get("level", "")
    code = req.get("code", "")
    model_type = req.get("model_type", "domestic_est")

    if level not in IO_TABLES:
        return json_dumps_safe({"error": "분류표가 로드되지 않았습니다."})

    t = IO_TABLES[level]
    idx_list = [i for i, s in enumerate(t["sectors"]) if s["code"] == code]
    if not idx_list:
        idx = 0
    else:
        idx = idx_list[0]

    sec = t["sectors"][idx]
    Z = t["Z"]
    X = t["supply_cols"]["9310"]
    VA = t["va_rows"].get("9690", np.zeros(t["n"]))
    IntIn = t["va_rows"].get("9590", np.zeros(t["n"]))
    leo = compute_leontief(level, model_type)

    x_j = float(X[idx])
    va_j = float(VA[idx])
    ii_j = float(IntIn[idx])

    # 후방 투입구조 Top 10 (j열)
    col_vals = Z[:, idx]
    up_order = np.argsort(-col_vals)
    upstream = []
    for i in up_order[:12]:
        val = float(col_vals[i])
        if val == 0 and len(upstream) >= 10:
            break
        upstream.append({
            "code": t["sectors"][i]["code"],
            "name": t["sectors"][i]["name"],
            "value": safe_num(val, 0),
            "coeff": safe_num(leo["A"][i, idx], 4),
            "share": safe_num(val / ii_j * 100 if ii_j else 0.0, 2)
        })

    # 전방 배분구조 Top 10 (i행)
    row_vals = Z[idx, :]
    down_order = np.argsort(-row_vals)
    downstream = []
    for j in down_order[:12]:
        val = float(row_vals[j])
        if val == 0 and len(downstream) >= 10:
            break
        downstream.append({
            "code": t["sectors"][j]["code"],
            "name": t["sectors"][j]["name"],
            "value": safe_num(val, 0),
            "coeff": safe_num(val / x_j if x_j > 0 else 0.0, 4),
            "share": safe_num(val / x_j * 100 if x_j > 0 else 0.0, 2)
        })

    # 생산유발 파급 Top 10 (L[:, idx])
    l_col = leo["L"][:, idx]
    l_sum = float(l_col.sum())
    l_order = np.argsort(-l_col)
    induced_top = []
    for i in l_order[:12]:
        val = float(l_col[i])
        if val <= 0 and len(induced_top) >= 10:
            break
        induced_top.append({
            "code": t["sectors"][i]["code"],
            "name": t["sectors"][i]["name"],
            "coeff": safe_num(val, 4),
            "share": safe_num(val / l_sum * 100 if l_sum else 0.0, 2)
        })

    # 4분면 산점도 데이터 (전체 부문)
    scatter = []
    for i, s in enumerate(t["sectors"]):
        scatter.append({
            "code": s["code"],
            "name": s["name"],
            "forward": safe_num(leo["forward"][i], 3),
            "backward": safe_num(leo["backward"][i], 3),
            "selected": (i == idx)
        })

    return json_dumps_safe({
        "sector": sec,
        "zero_output_warning": (x_j <= 0),
        "kpi": {
            "output": safe_num(x_j, 0),
            "va": safe_num(va_j, 0),
            "prod_mult": safe_num(leo["prod_mult"][idx], 3),
            "va_mult": safe_num(leo["va_mult"][idx], 3),
            "imp_mult": safe_num(leo["imp_mult"][idx], 3),
            "backward": safe_num(leo["backward"][idx], 3),
            "forward": safe_num(leo["forward"][idx], 3)
        },
        "upstream": upstream,
        "downstream": downstream,
        "induced_top": induced_top,
        "scatter": scatter
    })


def run_simulation(arg_json):
    req = json.loads(arg_json)
    level = req.get("level", "")
    model_type = req.get("model_type", "domestic_est")
    shocks = req.get("shocks", [])  # [{code, amount}]

    if level not in IO_TABLES:
        return json_dumps_safe({"error": "분류표가 로드되지 않았습니다."})

    t = IO_TABLES[level]
    n = t["n"]
    leo = compute_leontief(level, model_type)

    code_to_idx = {s["code"]: i for i, s in enumerate(t["sectors"])}
    dY = np.zeros(n, dtype=np.float64)
    for sh in shocks:
        c = sh.get("code", "")
        amt = to_float(sh.get("amount", 0.0))
        if c in code_to_idx:
            dY[code_to_idx[c]] += amt

    dX = leo["L"] @ dY
    dX_indirect = dX - dY
    dV = leo["v"] * dX
    dW = leo["w"] * dX
    dM = (leo["Am"] @ dX)  # 중간재 유발 수입액

    total_dy = float(dY.sum())
    total_dx = float(dX.sum())
    total_dv = float(dV.sum())
    total_dw = float(dW.sum())
    total_dm = float(dM.sum())

    order = np.argsort(-np.abs(dX))
    results = []
    for i in order:
        if abs(dX[i]) < 1e-6 and abs(dY[i]) < 1e-6:
            continue
        results.append({
            "code": t["sectors"][i]["code"],
            "name": t["sectors"][i]["name"],
            "direct": safe_num(dY[i], 1),
            "indirect": safe_num(dX_indirect[i], 1),
            "total_prod": safe_num(dX[i], 1),
            "share": safe_num(dX[i] / total_dx * 100 if total_dx else 0.0, 2),
            "va": safe_num(dV[i], 1),
            "wage": safe_num(dW[i], 1),
            "imp": safe_num(dM[i], 1)
        })

    return json_dumps_safe({
        "kpi": {
            "total_dy": safe_num(total_dy, 1),
            "total_dx": safe_num(total_dx, 1),
            "prod_multiplier": safe_num(total_dx / total_dy if total_dy else 0.0, 3),
            "total_dv": safe_num(total_dv, 1),
            "va_ratio": safe_num(total_dv / total_dy if total_dy else 0.0, 3),
            "total_dw": safe_num(total_dw, 1),
            "total_dm": safe_num(total_dm, 1)
        },
        "results": results
    })


def verify_all_levels():
    rows = []
    for lvl in ["major33", "mid83", "minor165", "basic380"]:
        if lvl not in IO_TABLES:
            continue
        t = IO_TABLES[lvl]
        sum_out = float(t["supply_cols"]["9310"].sum())
        sum_in = float(t["va_rows"].get("9790", np.zeros(t["n"])).sum())
        sum_sup = float(t["supply_cols"]["9390"].sum())
        sum_dem = float(t["supply_cols"]["9290"].sum())
        sum_va = float(t["va_rows"].get("9690", np.zeros(t["n"])).sum())

        max_diff_io = float(np.max(np.abs(t["supply_cols"]["9310"] - t["va_rows"].get("9790", np.zeros(t["n"])))))
        max_diff_sd = float(np.max(np.abs(t["supply_cols"]["9390"] - t["supply_cols"]["9290"])))
        max_err = max(max_diff_io, max_diff_sd)

        rows.append({
            "level": lvl,
            "label": t["level_label"],
            "n": t["n"],
            "sum_output": safe_num(sum_out, 0),
            "sum_input": safe_num(sum_in, 0),
            "sum_supply_demand": safe_num(sum_sup, 0),
            "sum_va": safe_num(sum_va, 0),
            "max_error": safe_num(max_err, 2),
            "passed": max_err < 5.0
        })
    return json_dumps_safe({"rows": rows})


def get_drilldown(arg_json):
    req = json.loads(arg_json)
    div_code = req.get("div_code", "")
    target_level = req.get("target_level", "mid83")
    model_type = req.get("model_type", "domestic_est")

    if not CLASSIFICATION["loaded"]:
        return json_dumps_safe({"error": "(표)상품분류.xlsx가 로드되어야 다계층 분해가 가능합니다."})
    if target_level not in IO_TABLES:
        return json_dumps_safe({"error": f"{LEVEL_LABELS.get(target_level, target_level)} 파일이 로드되지 않았습니다."})

    # 대분류에 속한 대상 분류 코드 집합 추출
    field_map = {"mid83": "major_code", "minor165": "minor_code", "basic380": "basic_code"}
    target_field = field_map.get(target_level)
    if not target_field:
        return json_dumps_safe({"error": "지원하지 않는 하위 분류입니다."})

    valid_codes = {
        item[target_field]
        for item in CLASSIFICATION["items"]
        if item["div_code"] == div_code
    }

    t = IO_TABLES[target_level]
    leo = compute_leontief(target_level, model_type)
    X = t["supply_cols"]["9310"]
    VA = t["va_rows"].get("9690", np.zeros(t["n"]))
    E = t["fd_cols"]["9140"]

    matched_indices = [i for i, s in enumerate(t["sectors"]) if s["code"] in valid_codes]
    sub_total_x = float(sum(X[i] for i in matched_indices))

    rows = []
    for i in matched_indices:
        x_i = float(X[i])
        va_i = float(VA[i])
        e_i = float(E[i])
        rows.append({
            "code": t["sectors"][i]["code"],
            "name": t["sectors"][i]["name"],
            "output": safe_num(x_i, 0),
            "share": safe_num(x_i / sub_total_x * 100 if sub_total_x else 0.0, 2),
            "va_rate": safe_num(va_i / x_i * 100 if x_i > 0 else 0.0, 2),
            "export_rate": safe_num(e_i / x_i * 100 if x_i > 0 else 0.0, 2),
            "prod_mult": safe_num(leo["prod_mult"][i], 3),
            "va_mult": safe_num(leo["va_mult"][i], 3)
        })

    rows.sort(key=lambda r: r["output"], reverse=True)
    return json_dumps_safe({"rows": rows, "sub_total_output": safe_num(sub_total_x, 0)})
