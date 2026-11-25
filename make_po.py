"""발주서 정리 일꾼 - 정리된 값(JSON)을 '우리 양식' 엑셀에 채운다.

사용법: python make_po.py <정리값.json> [설정.json]

저장 규칙 (설정.json에서 바꾼다)
  - 저장폴더: 결과를 모아 두는 곳. 설정.json이 있는 폴더 기준 (C:\\... 같은 전체 경로도 됨)
  - 폴더나누기: "월별"(2026-11/) / "거래처별"(사아물산/) / "없음"
  - 파일 이름: 발주일_거래처_발주번호_v버전[_확인필요N건].xlsx
  - 절대 덮어쓰지 않는다. 같은 발주서를 다시 정리하면 v2, v3…으로 새로 저장한다.
  - 작업기록.csv에 한 줄씩 남긴다.

이 코드는 값을 새로 만들지 않는다. 하는 일은 세 가지뿐이다.
  1) 비었거나 애매한 칸에 '확인 필요'를 적는다.
  2) 금액 = 수량 x 단가를 검산한다. 원본 금액과 다르면 '확인 필요'.
     (원본에 금액이 없으면 수량 x 단가로 계산해 넣고, 비고에 '계산값'이라고 적는다.)
  3) 합계를 원본 합계와 대조한다.
"""
import sys, json, re, csv, datetime as dt
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

U = "확인 필요"
RED = Font(color="C00000", bold=True)
YELLOW = PatternFill("solid", fgColor="FFF2CC")

HEAD_CELLS = {  # 우리 양식에서 각 값이 들어갈 칸
    "발주번호": "B3", "발주일": "B4", "거래처명": "B5", "담당자": "B6", "연락처": "B7",
    "납품장소": "G3", "납기일": "G4", "결제조건": "G5", "원본파일": "G6", "정리일": "G7", "비고": "B8",
}
HEAD_REQUIRED = ["발주번호", "발주일", "거래처명", "담당자", "연락처", "납품장소", "납기일", "결제조건"]
DATE_FIELDS = ["발주일", "납기일"]
ITEM_COLS = ["품목코드", "품명", "규격", "단위", "수량", "단가", "금액", "품목납기", "비고"]  # B~J
ITEM_REQUIRED = ["품명", "단위", "수량", "단가"]
FIRST_ROW, LAST_ROW, TOTAL_ROW = 11, 30, 31


def empty(v):
    return v is None or (isinstance(v, str) and v.strip() in ("", "-"))


def is_unknown(v):
    return empty(v) or (isinstance(v, str) and U in v)


def to_num(v):
    """숫자만 받는다. '1,200'처럼 쉼표가 있는 건 괜찮다. '5천', '약 300' 같은 건 받지 않는다."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str) and re.fullmatch(r"\s*-?[\d,]+(\.\d+)?\s*", v):
        s = v.replace(",", "").strip()
        return float(s) if "." in s else int(s)
    return None


def show(v):
    n = to_num(v)
    return f"{n:,}" if n is not None else str(v)


def ok_date(v):
    try:
        dt.date.fromisoformat(str(v)); return True
    except ValueError:
        return False


def safe(text):
    """파일 이름에 쓸 수 없는 글자와 '(가상)', '주식회사' 같은 말을 뺀다."""
    t = re.sub(r"\(가상\)|주식회사|\(주\)|㈜", "", str(text))
    t = re.sub(r'[\\/:*?"<>|\s]+', "", t)
    return t or "이름없음"


def plan_output(d, cfg, cfg_dir):
    """저장 규칙에 따라 저장 폴더와 파일 이름(버전 포함)을 정한다."""
    head = d.get("머리", {})
    root = Path(cfg.get("저장폴더", "정리완료"))
    if not root.is_absolute():
        root = cfg_dir / root
    day = head.get("발주일")
    day_ok = not is_unknown(day) and ok_date(day)
    vendor = "거래처미상" if is_unknown(head.get("거래처명")) else safe(head["거래처명"])
    po_no = "번호없음" if is_unknown(head.get("발주번호")) else safe(head["발주번호"])
    split = cfg.get("폴더나누기", "월별")
    if split == "월별":
        folder = root / (str(day)[:7] if day_ok else "날짜미상")
    elif split == "거래처별":
        folder = root / vendor
    else:
        folder = root
    base = f"{str(day).replace('-', '') if day_ok else '날짜미상'}_{vendor}_{po_no}"
    folder.mkdir(parents=True, exist_ok=True)
    used = [int(m.group(1)) for f in folder.glob(base + "_v*")
            if (m := re.match(re.escape(base) + r"_v(\d+)", f.name))]
    return root, folder, base, (max(used) + 1 if used else 1)


def write_log(root, row):
    log = root / "작업기록.csv"
    new = not log.exists()
    with open(log, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["정리시각", "원본파일", "정리값", "결과파일", "버전", "품목수", "합계", "확인필요"])
        w.writerow(row)


def main(src, cfg_path):
    cfg_path = Path(cfg_path)
    cfg = json.loads(cfg_path.read_text(encoding="utf-8")) if cfg_path.exists() else {}
    tpl = Path(cfg.get("양식", "우리양식_발주서.xlsx"))
    if not tpl.is_absolute():
        tpl = cfg_path.parent / tpl
    d = json.loads(Path(src).read_text(encoding="utf-8"))
    head, items = d.get("머리", {}), d.get("품목", [])
    checks = list(d.get("확인필요", []))  # 일꾼이 원본을 읽으며 미리 적어 둔 항목
    noted = {c.get("위치") for c in checks}

    def flag(where, original, why):
        if where not in noted:
            checks.append({"위치": where, "원본내용": original, "이유": why}); noted.add(where)

    wb = load_workbook(tpl)
    ws, wc = wb["발주서"], wb["확인 필요"]

    def put(cell, v):
        c = ws[cell] if isinstance(cell, str) else cell
        c.value = v
        if v == U:
            c.font, c.fill = RED, YELLOW

    # ---- 머리 정보 ----
    for k in HEAD_REQUIRED:
        v = head.get(k)
        if is_unknown(v):
            put(HEAD_CELLS[k], U); flag(k, "(없음)" if empty(v) else v, "원본에 없거나 애매함")
        elif k in DATE_FIELDS and not ok_date(v):
            put(HEAD_CELLS[k], U); flag(k, v, "날짜를 정확히 알 수 없음 (YYYY-MM-DD로 확정 필요)")
        else:
            put(HEAD_CELLS[k], v)
    put(HEAD_CELLS["원본파일"], d.get("원본파일", U))
    put(HEAD_CELLS["정리일"], dt.date.today().isoformat())
    put(HEAD_CELLS["비고"], head.get("비고") or "")

    # ---- 품목 ----
    if len(items) > LAST_ROW - FIRST_ROW + 1:
        sys.exit(f"품목이 {len(items)}개라 양식 칸(20줄)을 넘습니다. 양식을 늘려 주세요.")
    total, total_known = 0, True
    for i, it in enumerate(items, 1):
        r = FIRST_ROW + i - 1
        ws.cell(r, 1, i)
        vals = {k: it.get(k) for k in ITEM_COLS}
        notes = [vals["비고"]] if not empty(vals["비고"]) else []
        for k in ITEM_REQUIRED:
            if is_unknown(vals[k]):
                flag(f"{i}번 품목 {k}", "(없음)" if empty(vals[k]) else vals[k], "원본에 없거나 애매함")
                vals[k] = U
        q = to_num(vals["수량"]) if vals["수량"] != U else None
        p = to_num(vals["단가"]) if vals["단가"] != U else None
        for k, n in (("수량", q), ("단가", p)):
            if vals[k] != U and n is None:
                flag(f"{i}번 품목 {k}", vals[k], "숫자로 확정할 수 없음"); vals[k] = U
            elif n is not None:
                vals[k] = n
        amt_src = vals["금액"]
        if q is not None and p is not None:
            calc = q * p
            if empty(amt_src):
                vals["금액"] = calc; notes.append("금액은 수량×단가 계산값")
            else:
                a = to_num(amt_src)
                if a is None or abs(a - calc) > 0.5:
                    flag(f"{i}번 품목 금액", show(amt_src), f"원본 금액 {show(amt_src)} ≠ 수량×단가 {calc:,.0f}")
                    vals["금액"] = U
                else:
                    vals["금액"] = a
        else:
            a = to_num(amt_src) if not empty(amt_src) else None
            vals["금액"] = a if a is not None else U
            if a is None:
                flag(f"{i}번 품목 금액", "(없음)", "수량 또는 단가를 몰라 계산할 수 없음")
        if vals["금액"] == U:
            total_known = False
        else:
            total += vals["금액"]
        if not empty(vals["품목납기"]) and not is_unknown(vals["품목납기"]) and not ok_date(vals["품목납기"]):
            flag(f"{i}번 품목 납기", vals["품목납기"], "날짜를 정확히 알 수 없음"); vals["품목납기"] = U
        vals["비고"] = " / ".join(str(n) for n in notes)
        for j, k in enumerate(ITEM_COLS, 2):
            v = vals[k]
            put(ws.cell(r, j), "" if empty(v) else v)
            if k in ("수량", "단가", "금액") and isinstance(v, (int, float)):
                ws.cell(r, j).number_format = "#,##0"

    # ---- 합계 ----
    src_total = d.get("원본합계")
    if total_known:
        put(ws.cell(TOTAL_ROW, 8), total); ws.cell(TOTAL_ROW, 8).number_format = "#,##0"
        if src_total is not None and abs(to_num(src_total) - total) > 0.5:
            flag("합계", show(src_total), f"원본 합계 {show(src_total)} ≠ 품목 금액 합 {total:,.0f}")
    else:
        put(ws.cell(TOTAL_ROW, 8), U)
        flag("합계", show(src_total) if src_total is not None else "(없음)",
             f"금액을 모르는 품목이 있어 합계 확정 불가 (아는 것만 더하면 {total:,.0f})")

    # ---- 확인 필요 탭 ----
    for n, c in enumerate(checks, 1):
        wc.append([n, c.get("위치"), str(c.get("원본내용")), c.get("이유")])
        for cell in wc[wc.max_row]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    root, folder, base, ver = plan_output(d, cfg, cfg_path.parent)
    name = f"{base}_v{ver}" + (f"_확인필요{len(checks)}건" if checks else "")
    out = folder / f"{name}.xlsx"
    if out.exists():  # 버전 번호로 막지만, 혹시라도 같은 이름이 있으면 멈춘다
        sys.exit(f"이미 같은 파일이 있어 멈춥니다(덮어쓰지 않음): {out}")
    ws[HEAD_CELLS["정리일"]] = f"{dt.date.today().isoformat()} (v{ver})"
    wb.save(out)
    if cfg.get("작업기록", True):
        write_log(root, [dt.datetime.now().strftime("%Y-%m-%d %H:%M"), d.get("원본파일", ""), Path(src).name,
                         str(out.relative_to(root)), f"v{ver}", len(items),
                         f"{total:,.0f}" if total_known else U, len(checks)])

    print(f"완료: {out}")
    print(f"품목 {len(items)}개, 합계 {f'{total:,.0f}원' if total_known else U}")
    print(f"확인 필요 {len(checks)}건" + ("" if not checks else ":"))
    for c in checks:
        print(f"  - {c['위치']}: {c['이유']}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else Path(__file__).with_name("설정.json"))
