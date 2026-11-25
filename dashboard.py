# -*- coding: utf-8 -*-
"""
국내 상장사 실적 스크리닝 대시보드
실행:  streamlit run dashboard.py   (또는 3_dashboard.bat 더블클릭)
"""
import io
import os

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "financials.csv")
SALES, OP = "매출액(억원)", "영업이익(억원)"
C_SALES, C_OP, C_NEG = "#4C78A8", "#F58518", "#D64545"

st.set_page_config(page_title="실적 스크리닝", page_icon="📊", layout="wide")


@st.cache_data
def load():
    df = pd.read_csv(DATA, dtype={"종목코드": str, "결산월": str}, encoding="utf-8-sig")
    df["종목코드"] = df["종목코드"].str.zfill(6)
    for c in ("원본매출액(억)", "원본영업이익(억)"):
        if c not in df:
            df[c] = np.nan
    # 값이 하나도 없는 빈 줄은 제거 (아직 발표 전 기간 등)
    df = df.dropna(subset=[SALES, OP, "원본매출액(억)", "원본영업이익(억)"], how="all").copy()
    # 이상치 정리 (DART 공시 단위 오류)
    # - 500조원 초과, 또는 '회사 평소 규모의 1,000배 이상이면서 1조원 이상'인 값은 빈칸
    #   (매출이 이상하면 같은 기간 영업이익도 같은 오류라서 함께 빈칸)
    # - 음수 매출은 빈칸
    g = df.groupby("종목코드")
    usual = (g[SALES].transform(lambda s: s.abs().median())
             .fillna(g[OP].transform(lambda s: s.abs().median())).fillna(0))
    def broken(c):
        v = df[c].abs()
        return (v > 5e6) | ((v > usual * 1000) & (v > 1e4))
    bad_sales, bad_op = broken(SALES), broken(OP)
    df.loc[bad_sales, [SALES, OP]] = np.nan
    df.loc[bad_op, OP] = np.nan
    df.loc[df[SALES] < 0, SALES] = np.nan
    for c in ("원본매출액(억)", "원본영업이익(억)", "환율"):
        if c not in df:
            df[c] = np.nan
    if "매출기준" not in df:
        df["매출기준"] = np.where(df[SALES].notna(), "매출액", None)
    if "통화" not in df:
        df["통화"] = "KRW"
    df["통화"] = df["통화"].fillna("KRW")
    return df



def currency_label(df):
    """회사별 공시 통화 표시: 외화가 한 번이라도 있으면 'USD 환산' 등"""
    fx = df[df["통화"] != "KRW"].groupby("종목코드")["통화"].agg(lambda s: ",".join(sorted(set(s))) + " 환산")
    return pd.Series(df["종목코드"].unique()).to_frame("종목코드").assign(
        공시통화=lambda x: x["종목코드"].map(fx).fillna("KRW")).set_index("종목코드")["공시통화"]


def basis_label(df):
    """회사별 매출 기준 (가장 최근 값): 매출액 / 영업수익 / 이자·수수료·보험수익 합"""
    b = df.dropna(subset=["매출기준"]).sort_values(["연도", "분기"]).groupby("종목코드")["매출기준"].last()
    return b.reindex(df["종목코드"].unique()).fillna("-")


@st.cache_data(show_spinner="엑셀 파일 만드는 중...")
def build_excel(df):
    """전 종목 한 시트: 종목정보 | 연간 매출액 5년 | 연간 영업이익 5년 | 분기 매출액 20분기 | 분기 영업이익 20분기"""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    def latest(d, key, n):
        cnt = d.groupby(key)["종목코드"].nunique()
        last = cnt[cnt >= cnt.max() * 0.5].index.max()
        return sorted(k for k in d[key].unique() if k <= last)[-n:]

    info = df.drop_duplicates("종목코드", keep="last").set_index("종목코드")[["회사명", "시장", "결산월"]]
    info["공시통화"] = currency_label(df)
    info["매출기준"] = basis_label(df)
    a = df[df["분기"] == 0]
    years = latest(a, "연도", 5)
    q = df[df["분기"] > 0].assign(순서=lambda x: x["연도"] * 10 + x["분기"])
    qkeys = latest(q, "순서", 20)
    labels = [f"{k // 10}.{k % 10}Q" for k in qkeys]

    def wide(d, key, keys, col, names):
        w = d.pivot_table(index="종목코드", columns=key, values=col, aggfunc="first").reindex(columns=keys)
        w.columns = names
        return w

    blocks = [
        ("연간 매출액", wide(a, "연도", years, SALES, [str(y) for y in years])),
        ("연간 영업이익", wide(a, "연도", years, OP, [str(y) for y in years])),
        ("분기 매출액", wide(q, "순서", qkeys, SALES, labels)),
        ("분기 영업이익", wide(q, "순서", qkeys, OP, labels)),
    ]
    base = info.reset_index().sort_values(["시장", "회사명"]).reset_index(drop=True)
    data = pd.concat([base[["종목코드", "회사명", "시장", "결산월", "공시통화", "매출기준"]]]
                     + [b.reindex(base["종목코드"]).reset_index(drop=True) for _, b in blocks], axis=1)

    buf = io.BytesIO()
    data.to_excel(buf, sheet_name="전종목", index=False, startrow=1)
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb["전종목"]
    thin = Side(style="thin", color="BFBFBF")
    colors = {"연간 매출액": "DDEBF7", "연간 영업이익": "FCE4D6", "분기 매출액": "DDEBF7", "분기 영업이익": "FCE4D6"}

    def group(c1, c2, title):
        ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        ws.cell(row=1, column=c1, value=title)
        for c in range(c1, c2 + 1):
            for r in (1, 2):
                x = ws.cell(row=r, column=c)
                x.fill = PatternFill("solid", fgColor=colors.get(title, "D9D9D9"))
                x.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                x.font = Font(bold=True)
                x.alignment = Alignment(horizontal="center")

    group(1, 6, "종목 정보 (단위: 억원)")
    col = 7
    for title, b in blocks:
        group(col, col + b.shape[1] - 1, title)
        col += b.shape[1]
    n = len(data) + 2
    for c in range(7, col):
        ws.column_dimensions[get_column_letter(c)].width = 11
        for r in range(3, n + 1):
            ws.cell(row=r, column=c).number_format = "#,##0.0;[Red]-#,##0.0"
    for letter, w in zip("ABCDEF", (9, 22, 8, 7, 10, 20)):
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = "G3"
    ws.auto_filter.ref = f"A2:{get_column_letter(col - 1)}{n}"

    note = wb.create_sheet("설명")
    for i, line in enumerate([
        "출처: 금융감독원 DART (OpenDART 다중회사 주요계정 + 빈칸은 단일회사 전체 재무제표로 보완)",
        "단위: 억원 / 음수(적자)는 빨간색",
        "공시통화가 'USD 환산' 등인 회사: 외화 공시 금액을 기간 평균 환율(유럽중앙은행 기준)로 원화 환산한 값",
        "  → 원본 외화 금액과 적용 환율은 '외화공시_원본' 시트에 있음",
        "연결재무제표 우선, 연결이 없는 회사는 별도재무제표",
        "매출기준: 일반 기업은 '매출액'. 금융사는 매출 계정이 없어 '영업수익'(증권사 등) 또는",
        "  '이자·수수료·보험수익 합'(은행·금융지주·보험사)을 매출액 칸에 넣음 → 일반 기업과 직접 비교는 주의",
        "4분기 = 연간(사업보고서) − 3분기 누적",
        "빈칸 = 공시 데이터 없음(상장 전 등), 매출 계정을 찾지 못한 경우, 공시 단위 오류로 걸러낸 값",
        "결산월이 12월이 아닌 회사는 연도 기준이 다를 수 있음",
    ], 1):
        note.cell(row=i, column=1, value=line)
    note.column_dimensions["A"].width = 90

    fxrows = df[df["통화"] != "KRW"].copy()
    if len(fxrows):
        fxrows["기간"] = np.where(fxrows["분기"] == 0, "연간", fxrows["분기"].astype(str) + "Q")
        fxrows = fxrows.sort_values(["회사명", "연도", "분기"])[
            ["종목코드", "회사명", "연도", "기간", "통화", "원본매출액(억)", "원본영업이익(억)", "환율", SALES, OP]]
        fxrows.columns = ["종목코드", "회사명", "연도", "기간", "통화", "원본 매출액(억, 외화)", "원본 영업이익(억, 외화)",
                          "적용 환율(기간평균)", "환산 매출액(억원)", "환산 영업이익(억원)"]
        ws2 = wb.create_sheet("외화공시_원본", 1)
        ws2.append(list(fxrows.columns))
        for row in fxrows.itertuples(index=False):
            ws2.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
        for c in range(1, len(fxrows.columns) + 1):
            ws2.cell(row=1, column=c).font = Font(bold=True)
            ws2.cell(row=1, column=c).fill = PatternFill("solid", fgColor="E2EFDA")
            ws2.column_dimensions[get_column_letter(c)].width = 22 if c == 2 else 14
        for r in range(2, len(fxrows) + 2):
            for c in (6, 7):
                ws2.cell(row=r, column=c).number_format = "#,##0.00;[Red]-#,##0.00"
            ws2.cell(row=r, column=8).number_format = "#,##0.00"
            for c in (9, 10):
                ws2.cell(row=r, column=c).number_format = "#,##0.0;[Red]-#,##0.0"
        ws2.freeze_panes = "C2"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), f"실적_전종목_{labels[-1].replace('.', '')}.xlsx"

if not os.path.exists(DATA):
    st.error("financials.csv 가 없습니다. 먼저 2_collect.bat 을 실행해 데이터를 받아 주세요.")
    st.stop()

df = load()
names = df.drop_duplicates("종목코드").set_index("종목코드")[["회사명", "시장"]]


def growth(new, old):
    """증가율(%). 기준값이 0 이하면 계산하지 않음."""
    with np.errstate(divide="ignore", invalid="ignore"):
        g = np.where(old > 0, (new / old - 1) * 100, np.nan)
    return pd.Series(g, index=new.index)


def streak(op_wide):
    """마지막 기간까지 영업이익이 연속으로 늘어난 횟수"""
    arr = op_wide.to_numpy()
    out = []
    for row in arr:
        n = 0
        for i in range(len(row) - 1, 0, -1):
            if pd.notna(row[i]) and pd.notna(row[i - 1]) and row[i] > row[i - 1]:
                n += 1
            else:
                break
        out.append(n)
    return pd.Series(out, index=op_wide.index)


def trend(wide):
    return wide.apply(lambda r: [float(x) for x in r if pd.notna(x)], axis=1)


# ---------------------------------------------------------------- 사이드바
st.sidebar.header("보기")
mode = st.sidebar.radio("기준", ["연간", "분기"], horizontal=True)
markets = st.sidebar.multiselect("시장", ["코스피", "코스닥"], default=["코스피", "코스닥"])
query = st.sidebar.text_input("회사명 검색", placeholder="예: 삼성")

st.sidebar.header("스크리닝 조건")
min_sales = st.sidebar.number_input("최소 매출액 (억원)", min_value=0, value=0, step=100)
min_opm = st.sidebar.slider("영업이익률 최소 (%)", -50, 50, -50)
grow_label = "매출 연평균 성장률 최소 (%, 4년)" if mode == "연간" else "매출 전년동기 대비 증가율 최소 (%)"
min_growth = st.sidebar.slider(grow_label, -50, 100, -50)
only_profit = st.sidebar.checkbox("영업이익 흑자만")
only_turn = st.sidebar.checkbox("흑자전환 기업만")
min_streak = st.sidebar.slider("영업이익 연속 증가 (최소 횟수)", 0, 4 if mode == "연간" else 8, 0)

# ---------------------------------------------------------------- 기간 정리
if mode == "연간":
    d = df[df["분기"] == 0].copy()
    d["기간"] = d["연도"].astype(str)
    cnt = d.groupby("연도")["종목코드"].nunique()
    last = int(cnt[cnt >= cnt.max() * 0.5].index.max())
    periods = [str(y) for y in range(last - 4, last + 1)]
    prev, base = str(last - 1), str(last - 4)
else:
    d = df[df["분기"] > 0].copy()
    d["기간"] = d["연도"].astype(str) + "." + d["분기"].astype(str) + "Q"
    d["순서"] = d["연도"] * 10 + d["분기"]
    cnt = d.groupby("순서")["종목코드"].nunique()
    keys = sorted(cnt[cnt >= cnt.max() * 0.5].index)
    last_key = keys[-1]
    all_keys = sorted(d["순서"].unique())
    all_keys = [k for k in all_keys if k <= last_key][-20:]
    periods = [f"{k // 10}.{k % 10}Q" for k in all_keys]
    prev = f"{last_key // 10 - 1}.{last_key % 10}Q"  # 전년 동기
    base = prev

sales_w = d.pivot_table(index="종목코드", columns="기간", values=SALES, aggfunc="first").reindex(columns=periods)
op_w = d.pivot_table(index="종목코드", columns="기간", values=OP, aggfunc="first").reindex(columns=periods)
cur = periods[-1]

idx = sales_w.index.union(op_w.index)
sales_w, op_w = sales_w.reindex(idx), op_w.reindex(idx)
t = pd.DataFrame(index=idx)
t = t.join(names)
t["공시통화"] = currency_label(df).reindex(t.index).fillna("KRW")
t["매출기준"] = basis_label(df).reindex(t.index).fillna("-")
t["매출액"] = sales_w[cur]
t["영업이익"] = op_w[cur]
with np.errstate(divide="ignore", invalid="ignore"):
    t["영업이익률(%)"] = np.where(t["매출액"] > 0, t["영업이익"] / t["매출액"] * 100, np.nan)
if mode == "연간":
    with np.errstate(divide="ignore", invalid="ignore"):
        s0, s1 = sales_w[base], sales_w[cur]
        t["매출 성장률(%)"] = np.where((s0 > 0) & (s1 > 0), ((s1 / s0) ** (1 / 4) - 1) * 100, np.nan)
else:
    t["매출 성장률(%)"] = growth(sales_w[cur], sales_w[prev])
t["영업이익 증가율(%)"] = growth(op_w[cur], op_w[prev])
t["흑자전환"] = ((op_w[cur] > 0) & (op_w[prev] <= 0)).fillna(False).astype(bool)
t["연속증가"] = streak(op_w)
t["매출 추이"] = trend(sales_w)
t["영업이익 추이"] = trend(op_w)
t = t.dropna(subset=["매출액", "영업이익"], how="all")

# ---------------------------------------------------------------- 필터
f = t[t["시장"].isin(markets)]
if query:
    f = f[f["회사명"].str.contains(query, case=False, na=False)]
f = f[f["매출액"].fillna(0) >= min_sales]
if min_opm > -50:
    f = f[f["영업이익률(%)"] >= min_opm]
if min_growth > -50:
    f = f[f["매출 성장률(%)"] >= min_growth]
if only_profit:
    f = f[f["영업이익"] > 0]
if only_turn:
    f = f[f["흑자전환"].fillna(False).astype(bool)]
if min_streak:
    f = f[f["연속증가"] >= min_streak]
f = f.sort_values("매출액", ascending=False)

# ---------------------------------------------------------------- 화면
st.title("📊 국내 상장사 실적 스크리닝")
growth_note = f"{base}→{cur} 연평균" if mode == "연간" else f"{prev} 대비"
st.caption(f"기준 기간: **{cur}** · 매출 성장률은 {growth_note} · 영업이익 증가율은 {prev} 대비 · 단위: 억원 · 출처: DART")
xlsx_bytes, xlsx_name = build_excel(df)
st.download_button("📥 전종목 엑셀 다운로드 (연간 5년 + 분기 20분기)", xlsx_bytes, file_name=xlsx_name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

c1, c2, c3, c4 = st.columns(4)
c1.metric("조건에 맞는 기업", f"{len(f):,}개")
c2.metric("영업이익 흑자", f"{(f['영업이익'] > 0).sum():,}개")
c3.metric("흑자전환", f"{f['흑자전환'].sum():,}개")
c4.metric("영업이익률 중앙값", f"{f['영업이익률(%)'].median():.1f}%" if len(f) else "-")

show = f.reset_index()[["종목코드", "회사명", "시장", "공시통화", "매출기준", "매출액", "영업이익", "영업이익률(%)",
                         "매출 성장률(%)", "영업이익 증가율(%)", "매출 추이", "영업이익 추이"]]
st.caption("행을 클릭하면 아래에 해당 기업의 상세 실적이 나옵니다.")
event = st.dataframe(
    show, hide_index=True, width="stretch", height=480,
    on_select="rerun", selection_mode="single-row",
    column_config={
        "매출액": st.column_config.NumberColumn(f"매출액 ({cur})", format="localized"),
        "영업이익": st.column_config.NumberColumn(f"영업이익 ({cur})", format="localized"),
        "영업이익률(%)": st.column_config.NumberColumn(format="%.1f"),
        "매출 성장률(%)": st.column_config.NumberColumn(format="%.1f"),
        "영업이익 증가율(%)": st.column_config.NumberColumn(format="%.1f"),
        "매출 추이": st.column_config.LineChartColumn(f"매출 추이 ({periods[0]}~{cur})"),
        "영업이익 추이": st.column_config.LineChartColumn(f"영업이익 추이 ({periods[0]}~{cur})"),
    },
)
st.download_button("현재 목록 엑셀(CSV)로 받기",
                   show.drop(columns=["매출 추이", "영업이익 추이"]).to_csv(index=False).encode("utf-8-sig"),
                   file_name=f"screening_{mode}_{cur}.csv", mime="text/csv")

# ---------------------------------------------------------------- 상세
st.divider()
st.subheader("기업 상세")
if len(show) == 0:
    st.info("조건에 맞는 기업이 없습니다. 왼쪽 조건을 완화해 보세요.")
    st.stop()

rows = event.selection.rows if event and hasattr(event, "selection") else []
default_code = show.iloc[rows[0]]["종목코드"] if rows else show.iloc[0]["종목코드"]
options = show["종목코드"].tolist()
code = st.selectbox("기업 선택", options, index=options.index(default_code),
                    format_func=lambda c: f"{names.loc[c, '회사명']} ({c})")

one = df[df["종목코드"] == code]
detail_mode = st.radio("상세 기준", ["연간", "분기"], horizontal=True,
                       index=0 if mode == "연간" else 1, key="detail")
if detail_mode == "연간":
    one = one[one["분기"] == 0].sort_values("연도").tail(5)
    one["기간"] = one["연도"].astype(str)
else:
    one = one[one["분기"] > 0].sort_values(["연도", "분기"]).tail(20)
    one["기간"] = one["연도"].astype(str) + "." + one["분기"].astype(str) + "Q"

with np.errstate(divide="ignore", invalid="ignore"):
    one["영업이익률(%)"] = np.where(one[SALES] > 0, one[OP] / one[SALES] * 100, np.nan).round(1)

fig = go.Figure()
fig.add_bar(x=one["기간"], y=one[SALES], name="매출액", marker_color=C_SALES, opacity=0.85)
fig.add_scatter(x=one["기간"], y=one[OP], name="영업이익", yaxis="y2", mode="lines+markers",
                line=dict(color=C_OP, width=3),
                marker=dict(size=8, color=[C_NEG if v < 0 else C_OP for v in one[OP].fillna(0)]))
fig.update_layout(
    height=420, margin=dict(t=40, b=20, l=10, r=10), hovermode="x unified",
    legend=dict(orientation="h", y=1.12, x=0),
    yaxis=dict(title="매출액 (억원)", tickformat=",", rangemode="tozero"),
    yaxis2=dict(title="영업이익 (억원)", overlaying="y", side="right", tickformat=",", showgrid=False),
    xaxis=dict(type="category"),
)
st.plotly_chart(fig, width="stretch")

basis = one["기준"].dropna().unique()
sb = one["매출기준"].dropna().unique()
if len(sb) and any(x != "매출액" for x in sb):
    st.caption(f"⚠️ 금융사라 매출액 칸에 **{', '.join(sb)}**을(를) 넣었습니다. 일반 기업 매출액과 직접 비교할 때는 주의하세요.")
st.caption(f"재무제표 기준: {', '.join(basis) if len(basis) else '-'} · 결산월: {df[df['종목코드'] == code]['결산월'].iloc[0]}월")
st.dataframe(one.set_index("기간")[[SALES, OP, "영업이익률(%)"]].T, width="stretch")

fx_one = one[one["통화"] != "KRW"]
if len(fx_one):
    cur_names = ", ".join(sorted(fx_one["통화"].unique()))
    st.info(f"이 회사는 일부 기간을 **{cur_names}**로 공시했습니다. 위 숫자는 기간 평균 환율로 원화 환산한 값이고, "
            "아래는 공시 원본 금액입니다.")
    st.dataframe(fx_one.set_index("기간")[["통화", "원본매출액(억)", "원본영업이익(억)", "환율"]]
                 .rename(columns={"원본매출액(억)": "원본 매출액(억, 외화)", "원본영업이익(억)": "원본 영업이익(억, 외화)",
                                  "환율": "적용 환율"}), width="stretch")
