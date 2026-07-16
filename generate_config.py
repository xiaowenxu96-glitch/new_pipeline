#!/usr/bin/env python3
"""Generate db_chemical2.yaml correctly."""

lines = []
lines.append('industry: "化工农药"')
lines.append('source_file: "D:/SWS/new_pipeline/化工农药/化工农药20260610_【申万化工细分行业及重点公司一周概览】_20260713090615.xlsx"')
lines.append('target_file: "D:/SWS/new_pipeline/化工农药/申万化工细分行业目标文档.xlsx"')
lines.append('output_dir: "D:/SWS/new_pipeline/化工农药/output"')
lines.append('')
lines.append('defaults:')
lines.append('  start_date: "2016-01-03"')
lines.append('  date_format: "yyyy-mm"')
lines.append('  start_row: 44')
lines.append('')
lines.append('sheets:')
lines.append('  # ============================================================')
lines.append('  # 原油与成品油')
lines.append('  # ============================================================')
lines.append('  - sheet_name: "原油与成品油"')
lines.append('    source_sheet: "原油与成品油"')
lines.append('    sections:')
lines.append('      - date_col: "AD"')
lines.append('        data_start_row: 44')
lines.append('        indicators:')
lines.append('          "AA000301905700": "AE"')
lines.append('          "AA000302354000": "AF"')
lines.append('      - date_col: "AG"')
lines.append('        data_start_row: 44')
lines.append('        indicators:')
lines.append('          "AA000293892000": "AL"')
lines.append('          "AA000293806300": "AM"')
lines.append('      - date_col: "AK"')
lines.append('        data_start_row: 44')
lines.append('        indicators:')
lines.append('          "AA000293860500": "AN"')
lines.append('      - date_col: "AR"')
lines.append('        data_start_row: 44')
lines.append('        indicators:')
lines.append('          "AA000293757300": "AO"')
lines.append('      - date_col: "AY"')
lines.append('        data_start_row: 44')
lines.append('        indicators: {}')
lines.append('    actions:')
lines.append('      - type: "chemical_write_sheet"')
lines.append('      - type: "chemical_write_formulas"')
lines.append('        start_row: 44')
lines.append('        anchor_cell: "AG4"')
lines.append('        reference_col: "AD"')

# columns (indented properly inside action)
clines = []
clines.append('        columns:')
clines.append('          - col: "C"')
clines.append('            format: "0.00_ "')
clines.append('            lookup:')
clines.append('              key_col: "B"')
clines.append('              start_col: "AD"')
clines.append('              end_col: "AF"')
clines.append('              return_index: 2')
clines.append('          - col: "D"')
clines.append('            format: "0.00_ "')
clines.append('            lookup:')
clines.append('              key_col: "B"')
clines.append('              start_col: "AD"')
clines.append('              end_col: "AF"')
clines.append('              return_index: 3')
clines.append('          - col: "I"')
clines.append('            formula: "=B{row}"')
clines.append('            format: "mm-dd-yy"')
clines.append('          - col: "J"')
clines.append('            format: "0.00_ "')
clines.append('            lookup:')
clines.append('              key_col: "I"')
clines.append('              start_col: "AK"')
clines.append('              end_col: "AO"')
clines.append('              return_index: 2')
clines.append('          - col: "K"')
clines.append('            format: "0.00_ "')
clines.append('            lookup:')
clines.append('              key_col: "I"')
clines.append('              start_col: "AK"')
clines.append('              end_col: "AO"')
clines.append('              return_index: 3')
clines.append('          - col: "L"')
clines.append('            format: "0.00_ "')
clines.append('            lookup:')
clines.append('              key_col: "I"')
clines.append('              start_col: "AK"')
clines.append('              end_col: "AO"')
clines.append('              return_index: 4')
clines.append('          - col: "M"')
clines.append('            format: "0.00_ "')
clines.append('            lookup:')
clines.append('              key_col: "I"')
clines.append('              start_col: "AK"')
clines.append('              end_col: "AO"')
clines.append('              return_index: 5')

# static cells
slines = []
slines.append('        static_cells:')
slines.append('          - cell: "E5"')
slines.append('            formula: "=AG4"')
slines.append('          - cell: "F5"')
slines.append('            formula: "=E5-7"')
slines.append('            format: "yyyy/m/d;@"')
slines.append('          - cell: "G5"')
slines.append('            formula: "=AG4"')
slines.append('            format: "yyyy/m/d;@"')
slines.append('          - cell: "H5"')
slines.append('            value: "2026-03-20"')
slines.append('            format: "yyyy/m/d;@"')
slines.append('          - cell: "AK5"')
slines.append('            formula: "=EDATE($E5,-1)+(4-WEEKDAY(EDATE($E5,-1),2))"')
slines.append('            format: "yyyy/m/d;@"')
slines.append('          - cell: "AL5"')
slines.append('            formula: "=EDATE($E5,-3)+(4-WEEKDAY(EDATE($E5,-3),2))"')
slines.append('            format: "yyyy/m/d;@"')
slines.append('          - cell: "AM5"')
slines.append('            formula: "=EDATE($E5,-12)+(4-WEEKDAY(EDATE($E5,-12),2))"')
slines.append('            format: "yyyy/m/d;@"')

def sc(cell_addr, formula_or_value, fmt=None):
    """Generate a static_cells entry."""
    key = "value" if cell_addr.startswith("H5") or cell_addr.startswith("M5") or cell_addr.startswith("N5") or cell_addr.startswith("O5") or cell_addr.startswith("P5") or cell_addr.startswith("Q5") or cell_addr.startswith("R5") or cell_addr.startswith("S5") or cell_addr.startswith("T5") or cell_addr.startswith("U5") or cell_addr.startswith("V5") or cell_addr.startswith("W5") or cell_addr.startswith("X5") or cell_addr.startswith("Y5") or cell_addr.startswith("Z5") or cell_addr.startswith("AA5") or cell_addr.startswith("AB5") or cell_addr.startswith("AC5") or cell_addr.startswith("AD5") else "formula"
    if key == "formula":
        s = f'          - cell: "{cell_addr}"\n            formula: \'{formula_or_value}\''
    else:
        s = f'          - cell: "{cell_addr}"\n            value: "{formula_or_value}"\n            format: "yyyy/m/d;@"'
    return s

# E6:E22 — references to data columns at row 44
e_44_cols = "C,D,F,G,H,J,K,L,M,Q,R,S,T,X,Y,Z,AA".split(",")
for i, dc in enumerate(e_44_cols):
    r = 6 + i
    slines.append(f'          - cell: "E{r}"')
    slines.append(f'            formula: "={dc}44"')

# F6:F22 — references to data columns at row 45
for i, dc in enumerate(e_44_cols):
    r = 6 + i
    slines.append(f'          - cell: "F{r}"')
    slines.append(f'            formula: "={dc}45"')

# G,H,I,J,K,L rows 6..22 — formulas
for r in range(6, 23):
    slines.append(f'          - cell: "G{r}"')
    slines.append(f'            formula: "=(E{r}-AI{r})/(AJ{r}-AI{r})"')
    slines.append(f'          - cell: "H{r}"')
    slines.append(f'            formula: "=(F{r}-AI{r})/(AJ{r}-AI{r})"')
    slines.append(f'          - cell: "I{r}"')
    slines.append(f'            formula: "=($E{r}-F{r})/ABS(F{r})"')
    slines.append(f'          - cell: "J{r}"')
    slines.append(f'            formula: "=($E{r}-AK{r})/ABS(AK{r})"')
    slines.append(f'          - cell: "K{r}"')
    slines.append(f'            formula: "=($E{r}-AL{r})/ABS(AL{r})"')
    slines.append(f'          - cell: "L{r}"')
    slines.append(f'            formula: "=($E{r}-AM{r})/ABS(AM{r})"')

# M6:M22 — AVERAGEIFS, different column per row
m_data_cols = "AE,AF,AH,AI,AJ,AL,AM,AN,AO,AS,AT,AU,AV,AZ,BA,BB,BC".split(",")
for i, mc in enumerate(m_data_cols):
    r = 6 + i
    slines.append(f'          - cell: "M{r}"')
    slines.append(f'            formula: \'=IFERROR(AVERAGEIFS(${mc}$49:${mc}$6000,$AD$49:$AD$6000,">="&$AD$49,$AD$49:$AD$6000,"<="&EOMONTH($AD$49,11)),"")\'')

# AI6:AJ22 — historical min/max, different column per row
ai_data_cols = "AE,AF,AH,AI,AJ,AL,AM,AN,AO,AS,AT,AU,AV,AZ,BA,BB,BC".split(",")
for i, mc in enumerate(ai_data_cols):
    r = 6 + i
    slines.append(f'          - cell: "AI{r}"')
    slines.append(f'            formula: "=MIN(IF(${mc}$44:${mc}$10000>0,${mc}$44:${mc}$10000))"')
    slines.append(f'            is_array: "AI{r}"')
    slines.append(f'          - cell: "AJ{r}"')
    slines.append(f'            formula: "=MAX(IF(${mc}$44:${mc}$10000>0,${mc}$44:${mc}$10000))"')
    slines.append(f'            is_array: "AJ{r}"')

# AK6:AM22 — VLOOKUP, increasing return_index per row
ak_indices = [2,3,5,6,7,9,10,11,12,16,17,18,19,23,24,25,26]
for i, idx in enumerate(ak_indices):
    r = 6 + i
    slines.append(f'          - cell: "AK{r}"')
    slines.append(f'            formula: "=VLOOKUP(AK5,B44:AA6000,{idx},FALSE)"')
    slines.append(f'          - cell: "AL{r}"')
    slines.append(f'            formula: "=VLOOKUP(AL5,B44:AA6000,{idx},FALSE)"')
    slines.append(f'          - cell: "AM{r}"')
    slines.append(f'            formula: "=VLOOKUP(AM5,B44:AA6000,{idx},FALSE)"')

# N6:U22 — 历史均价 (AVERAGEIFS per year)
# N5=2024, O5=2023, ..., U5=2017
# Each row uses a different data column (same as M column)
year_col_map = {
    "N": "YEAR(N$5)",
    "O": "YEAR(O$5)",
    "P": "YEAR(P$5)",
    "Q": "YEAR(Q$5)",
    "R": "YEAR(R$5)",
    "S": "YEAR(S$5)",
    "T": "YEAR(T$5)",
    "U": "YEAR(U$5)",
}
for year_col_letter, year_formula in year_col_map.items():
    for i, mc in enumerate(m_data_cols):
        r = 6 + i
        slines.append(f'          - cell: "{year_col_letter}{r}"')
        slines.append(f'            formula: \'=IFERROR(AVERAGEIFS(${mc}$44:${mc}$10000,$AD$44:$AD$10000,">="&DATE({year_formula},1,1),$AD$44:$AD$10000,"<="&DATE({year_formula},12,31)),"")\'')

# V6:AD22 — 历史分位数: (年历史均价 - AI) / (AJ - AI)
# V→M, W→N, X→O, Y→P, Z→Q, AA→R, AB→S, AC→T, AD→U
year_letter_map = {
    "V": "M", "W": "N", "X": "O", "Y": "P",
    "Z": "Q", "AA": "R", "AB": "S", "AC": "T", "AD": "U"
}
for col_letter, avg_col in year_letter_map.items():
    for i in range(17):
        r = 6 + i
        slines.append(f'          - cell: "{col_letter}{r}"')
        slines.append(f'            formula: "=({avg_col}{r}-AI{r})/(AJ{r}-AI{r})"')

all_lines = lines + clines + slines

with open("D:/SWS/new_pipeline/configs/db_chemical2.yaml", "w", encoding="utf-8") as f:
    f.write("\n".join(all_lines) + "\n")

print(f"Done! Total lines: {len(all_lines)}")
