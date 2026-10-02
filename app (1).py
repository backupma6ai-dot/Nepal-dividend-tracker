import io
import time

import pandas as pd
import streamlit as st
import requests

import sharesansar_dividends as sd

st.set_page_config(page_title="ShareSansar Dividend History", layout="wide")
st.title("ShareSansar Dividend History")
st.caption("Cash dividend and bonus share, last N fiscal years")

group = st.radio("List", ["companies", "funds"], horizontal=True)
options = list(sd.GROUPS[group])
chosen = st.multiselect("Symbols (leave empty = all in list)", options)
years = st.slider("Fiscal years", 1, 10, 5)
delay = st.slider("Delay between requests (sec)", 0.5, 5.0, 1.5, 0.5)

if st.button("Fetch dividend history", type="primary"):
    symbols = chosen or options
    session = requests.Session()
    frames, failures = [], []
    bar = st.progress(0.0)
    status = st.empty()

    for i, sym in enumerate(symbols, 1):
        status.text(f"Fetching {sym} ({i}/{len(symbols)})")
        try:
            df = sd.normalise(sd.fetch_dividends(session, sym))
            if "Fiscal Year" in df.columns:
                df["_k"] = df["Fiscal Year"].map(sd.fiscal_key)
                df = df.sort_values("_k", ascending=False).head(years).drop(columns="_k")
            else:
                df = df.head(years)
            df.insert(0, "Name", sd.GROUPS["all"].get(sym, ""))
            df.insert(0, "Symbol", sym)
            frames.append(df)
        except Exception as e:
            failures.append((sym, str(e)))
        bar.progress(i / len(symbols))
        time.sleep(delay)
    status.empty()

    if frames:
        out = pd.concat(frames, ignore_index=True)
        for col in ("Bonus Share (%)", "Cash Dividend (%)", "Total Dividend (%)"):
            if col in out.columns:
                out[col] = pd.to_numeric(
                    out[col].astype(str).str.replace("%", "").str.strip(), errors="coerce")
        front = ["Symbol", "Name", "Fiscal Year", "Bonus Share (%)",
                 "Cash Dividend (%)", "Total Dividend (%)"]
        out = out[[c for c in front if c in out.columns] +
                  [c for c in out.columns if c not in front]]
        st.session_state["result"] = out
    st.session_state["failures"] = failures

out = st.session_state.get("result")
if out is not None:
    st.subheader("Result")
    st.dataframe(out, use_container_width=True)

    if "Fiscal Year" in out.columns:
        c1, c2 = st.columns(2)
        for col, box, title in (("Cash Dividend (%)", c1, "Cash dividend (%)"),
                                ("Bonus Share (%)", c2, "Bonus share (%)")):
            if col in out.columns:
                pv = out.pivot_table(index="Symbol", columns="Fiscal Year",
                                     values=col, aggfunc="first")
                pv = pv[sorted(pv.columns, key=sd.fiscal_key, reverse=True)]
                box.subheader(title)
                box.dataframe(pv, use_container_width=True)

    buf = io.BytesIO()
    out.to_excel(buf, index=False)
    d1, d2 = st.columns(2)
    d1.download_button("Download CSV", out.to_csv(index=False).encode("utf-8-sig"),
                       "dividend_history.csv", "text/csv")
    d2.download_button("Download Excel", buf.getvalue(), "dividend_history.xlsx")

for sym, err in st.session_state.get("failures", []):
    st.warning(f"{sym}: {err}")
