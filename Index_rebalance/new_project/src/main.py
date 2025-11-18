import pandas as pd

def main():

    # Read entire Excel file with all sheets
    excel_file = pd.ExcelFile('IndexWeightData.xlsx')
    # Get all sheet names
    sheet_names = excel_file.sheet_names
    print(f"Available sheets: {sheet_names}")

    # Read specific sheets
    df_300 = pd.read_excel('IndexWeightData.xlsx', sheet_name='SHSZ300')
    df_500 = pd.read_excel('IndexWeightData.xlsx', sheet_name='CSI500')
    df_1000 = pd.read_excel('IndexWeightData.xlsx', sheet_name='CSI1000')
    combined = pd.concat([df_300, df_500, df_1000], ignore_index=True)
    print(len(combined))
    # Rank by FF_mkt_cap descending
    ranked = combined.sort_values('FF_mkt_cap', ascending=False).copy()
    ranked['FF_rank_desc'] = range(1, len(ranked) + 1)
    # Buffer rule for CSI300 members with rank <= 360
    csi300_mask = ranked.get('Index', pd.Series(index=ranked.index, dtype=str)).astype(str).str.upper().str.contains('CSI300')
    top240 = ranked[ranked['FF_rank_desc'] <= 240]
    buffer_prev = ranked[(ranked['FF_rank_desc'] <= 360) & csi300_mask]
    selected_idx = pd.Index(top240.index).union(buffer_prev.index)
    if len(selected_idx) < 300:
        need = 300 - len(selected_idx)
        filler = ranked[~ranked.index.isin(selected_idx)].head(need)
        selected_idx = selected_idx.union(filler.index)
    selected = ranked.loc[selected_idx].sort_values('FF_mkt_cap', ascending=False)
    print(len(selected))
    print(selected.head(20).to_string(index=False))
    
    selected = selected[["Ticker","FF_mkt_cap","Weight","Index"]]
    selected.to_excel('test.xlsx', index=False)



if __name__ == "__main__":
    main()