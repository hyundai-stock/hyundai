"""현대·제네시스 재고 엑셀 → data.json 변환기

사용법:  python tools/convert.py <엑셀파일...>
  - 파일 이름은 상관없습니다. 첫 줄(제목줄)을 보고 종류를 자동으로 구분합니다.
      · '요청' '재고' 칸이 있으면      → 재고현황 (즉시출고 탭)
      · '판매조건계' 칸이 있으면       → 판매조건 재고 (할인재고 탭)
  - 한 종류만 올려도 됩니다. 안 올린 종류는 기존 data.json 내용을 그대로 유지합니다.
"""
import sys, os, re, json, datetime, collections
import openpyxl

OUT = os.path.join(os.getcwd(), "data.json")   # 저장소 맨 위 폴더에서 실행
KST = datetime.timezone(datetime.timedelta(hours=9))

# 시트 이름 → 화면에 보일 차종명 (오타·줄임말 보정)
MODEL_NAME = {
    "팰리": "팰리세이드", "아이오니5": "아이오닉5", "아반떼 신형": "아반떼(신형)",
}
GENESIS = re.compile(r"^(G\d|GV\d)")


def model_of(sheet):
    s = re.sub(r"\s*\(\d+\)\s*$", "", sheet).strip()
    return MODEL_NAME.get(s, s)


def clean(s):
    return re.sub(r"\s+", " ", str(s if s is not None else "")).strip()


def trim_of(s):
    t = clean(s)
    t = re.sub(r"(^| )\.(?= |$)", " ", t)      # 혼자 있는 '.' 제거
    return clean(t)


def num(v):
    return int(v) if isinstance(v, (int, float)) else 0


def man(v):
    return round(v / 10000)


def kind_of(header):
    h = [clean(x) for x in header]
    if "판매조건계" in h:
        return "sale"
    if "재고" in h and "요청" in h:
        return "now"
    return None


def read_now(ws, order):
    """재고현황: 판매코드|옵션코드|외장코드|내장코드|요청|재고|차종|옵션|외장|내장|가격"""
    m = model_of(ws.title)
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0] or clean(r[0]) == "합계":
            continue
        q = num(r[5])
        if q <= 0 or not isinstance(r[10], (int, float)):
            continue  # 재고 없이 '요청'만 있는 줄은 제외
        out.append(dict(m=m, t=trim_of(r[6]), o=clean(r[7]) or "무옵션", e=clean(r[8]), i=clean(r[9]),
                        p=int(r[10]), q=q, k=f"{clean(r[0])}|{clean(r[1])}|{clean(r[2])}|{clean(r[3])}"))
    order.setdefault(m, len(order))
    return out


def read_sale(ws, order):
    """판매조건: ''|구분|판매코드|옵션코드|외장코드|내장코드|파츠코드|판매조건계|기본|생산월|특별|한정|계약번호|출고예정일|차종|옵션|외장|내장|가격|비고"""
    m = model_of(ws.title)
    g = collections.OrderedDict()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not isinstance(r[18], (int, float)):
            continue
        row = dict(m=m, t=trim_of(r[14]), o=clean(r[15]) or "무옵션", e=clean(r[16]), i=clean(r[17]),
                   p=int(r[18]), s=num(r[7]), sb=num(r[8]), sm=num(r[9]), ss=num(r[10]), sl=num(r[11]),
                   n=clean(r[19]), k=f"{clean(r[2])}|{clean(r[3])}|{clean(r[4])}|{clean(r[5])}")
        key = tuple(row[f] for f in ("k", "p", "s", "sb", "sm", "ss", "sl", "n"))
        if key in g:
            g[key]["q"] += 1
        else:
            row["q"] = 1
            g[key] = row
    order.setdefault(m, len(order))
    out = []
    for v in g.values():
        for f in ("sb", "sm", "ss", "sl"):   # 0원인 항목은 저장하지 않음
            if not v[f]:
                v.pop(f)
        if not v["n"]:
            v.pop("n")
        out.append(v)
    return out


def main(paths):
    data = {}
    if os.path.exists(OUT):
        with open(OUT, encoding="utf-8") as f:
            data = json.load(f)
    now_s = datetime.datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    found = {}
    for p in paths:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        first = next(wb.worksheets[0].iter_rows(max_row=1, values_only=True), ())
        kind = kind_of(first)
        if not kind:
            print(f"⚠ 종류를 알 수 없어 건너뜀: {os.path.basename(p)}")
            continue
        if kind in found:
            print(f"⚠ 같은 종류 파일이 여러 개입니다. 마지막 파일을 사용: {os.path.basename(p)}")
        rows, order = [], {}
        for ws in wb.worksheets:
            rows += (read_now if kind == "now" else read_sale)(ws, order)
        found[kind] = rows
        print(f"✔ {os.path.basename(p)} → {'즉시출고' if kind == 'now' else '할인재고'}"
              f"  {len(rows)}개 사양 / {sum(r['q'] for r in rows):,}대")
    if not found:
        sys.exit("변환할 엑셀 파일이 없습니다.")

    for k, rows in found.items():
        data[k] = rows
        data.setdefault("time", {})[k] = now_s
    data.setdefault("now", [])
    data.setdefault("sale", [])
    data["updated"] = now_s

    # 즉시출고 줄에 '같은 사양 할인재고' 정보 붙이기
    best = {}
    for r in data["sale"]:
        b = best.setdefault(r["k"], [0, 0])
        b[0] += r["q"]
        b[1] = max(b[1], r["s"])
    for r in data["now"]:
        r.pop("ds", None)
        if r["k"] in best:
            r["ds"] = best[r["k"]]          # [할인재고 대수, 최대 할인액]
    data["genesis"] = sorted({r["m"] for k in ("now", "sale") for r in data[k] if GENESIS.match(r["m"])})

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print(f"저장 완료: data.json ({now_s} 기준)")


if __name__ == "__main__":
    main(sys.argv[1:])
