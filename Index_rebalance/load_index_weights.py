import pandas as pd

data = pd.read_excel("IndexWeightData.xlsx", sheet_name=None)
names = list(data.keys())
sizes = {k: len(v) for k, v in data.items()}


def norm(s):
    return s.upper().replace(" ", "")

lookup = {norm(k): k for k in names}

def find(keys):
    for key in keys:
        nk = norm(key)
        for n, orig in lookup.items():
            if nk in n:
                return orig
    return None
#a          jfsjjsjsj
#kdjfjfkd
s300 = find(["SHSZ300", "CSI300"])
s500 = find(["CSI500"])
s1000 = find(["CSI1000"])
targets = [x for x in [s300, s500, s1000] if x]
if s1000:
    df1000 = data[s1000]
    print(len(df1000))
    print(df1000.head(20).to_string(index=False))
dfs = [data[name].copy().assign(Source=name) for name in targets]

combined = pd.concat(dfs, ignore_index=True)
print(len(combined))
if "Source" in combined.columns:
    print(combined["Source"].value_counts(dropna=False).to_string())
combined.to_excel("combined.xlsx", index=False)
if "FF_mkt_cap" in combined.columns:
    tmp = combined.sort_values("FF_mkt_cap", ascending=False)
    r = pd.Series(range(1, len(tmp) + 1), index=tmp.index)
    combined.loc[r.index, "FF_rank_desc"] = r
    prev_mask = combined.get("Source", pd.Series(index=combined.index, dtype=str)).astype(str).str.upper().str.contains("SHSZ300|CSI300")
    top240 = combined[combined["FF_rank_desc"] <= 240]
    buffer_prev = combined[(combined["FF_rank_desc"] <= 360) & prev_mask]
    sel_idx = pd.Index(top240.index).union(buffer_prev.index)
    if len(sel_idx) < 300:
        need = 300 - len(sel_idx)
        filler = combined[~combined.index.isin(sel_idx)].sort_values("FF_mkt_cap", ascending=False).head(need)
        sel_idx = sel_idx.union(filler.index)
    selected = combined.loc[sel_idx].sort_values("FF_mkt_cap", ascending=False)
    print(len(selected))
    print(selected.head(20).to_string(index=False))
    selected.to_excel('test.xlsx')