# -*- coding: utf-8 -*-
"""
국내 상장사 실적 스크리닝 대시보드
실행:  streamlit run dashboard.py   (또는 3_dashboard.bat 더블클릭)
"""
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
    # 이상치 정리: 음수 매출, 회사 평소 규모의 50배를 넘는 값(공시 단위 오류)은 빈칸 처리
    df.loc[df[SALES] < 0, SALES] = np.nan
    scale = df.groupby("종목코드")[SALES].transform(lambda s: s.abs().median())
    scale = scale.fillna(df.groupby("종목코드")[OP].transform(lambda s: s.abs().median()))
    for c in (SALES, OP):
        df.loc[df[c].abs() > scale * 50, c] = np.nan
    return df


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

c1, c2, c3, c4 = st.columns(4)
c1.metric("조건에 맞는 기업", f"{len(f):,}개")
c2.metric("영업이익 흑자", f"{(f['영업이익'] > 0).sum():,}개")
c3.metric("흑자전환", f"{f['흑자전환'].sum():,}개")
c4.metric("영업이익률 중앙값", f"{f['영업이익률(%)'].median():.1f}%" if len(f) else "-")

show = f.reset_index()[["종목코드", "회사명", "시장", "매출액", "영업이익", "영업이익률(%)",
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
st.caption(f"재무제표 기준: {', '.join(basis) if len(basis) else '-'} · 결산월: {df[df['종목코드'] == code]['결산월'].iloc[0]}월")
st.dataframe(one.set_index("기간")[[SALES, OP, "영업이익률(%)"]].T, width="stretch")
