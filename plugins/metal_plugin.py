import pandas as pd
from openpyxl.utils import column_index_from_string as col_letter_to_num


class MetalPlugin:
    """有色金属 Pipeline 插件：按 AA 指标代码匹配列写入数据"""

    @staticmethod
    def metal_write_sheet(context, params):
        ws = context['ws']
        reader = context['data_reader']
        sc = context['sheet_config']

        source_sheet = sc['source_sheet']
        data_start = sc.get('data_start_row', 6)
        date_col = sc.get('date_col', 'A')
        indicator_map = sc.get('indicators', {})

        df = reader.read_sheet_data(source_sheet)

        # AA 代码在源文件的第 2 行 (pandas row index 1)
        aa_row = df.iloc[1]

        # 建立 AA code → 源列索引 (0-based) 的映射
        aa_to_src_col = {}
        for c in range(1, len(df.columns)):
            code = str(aa_row.iloc[c]).strip() if pd.notna(aa_row.iloc[c]) else ''
            if code:
                aa_to_src_col[code] = c

        # 数据从第 11 行开始 (pandas row index 10)
        data_df = df.iloc[10:].reset_index(drop=True)
        if data_df.empty:
            print(f"    - [{source_sheet}] 无数据")
            return

        data_df.columns = [f'col_{i}' for i in range(len(data_df.columns))]

        # 解析日期（第 0 列）
        data_df['date'] = pd.to_datetime(data_df['col_0'], errors='coerce')
        data_df = data_df[data_df['date'].notna()].copy()
        # 从最新到最早排序
        data_df = data_df.sort_values('date', ascending=False).reset_index(drop=True)

        print(f"    - [{source_sheet}] {len(data_df)} 行数据, {len(indicator_map)} 个指标")

        # 写日期到目标日期列
        date_col_num = col_letter_to_num(date_col)
        for i, dv in enumerate(data_df['date']):
            c = ws.cell(row=data_start + i, column=date_col_num)
            c.value = dv.to_pydatetime() if isinstance(dv, pd.Timestamp) else dv
            c.number_format = 'yyyy-mm-dd'

        # 写指标数据到目标列
        written = 0
        for aa_code, target_col_letter in indicator_map.items():
            if aa_code not in aa_to_src_col:
                print(f"    !! 源中未找到 AA 代码: {aa_code}")
                continue

            src_col_idx = aa_to_src_col[aa_code]
            src_col_name = f'col_{src_col_idx}'
            tgt_col_num = col_letter_to_num(target_col_letter)

            for i, val in enumerate(data_df[src_col_name]):
                if pd.notna(val):
                    c = ws.cell(row=data_start + i, column=tgt_col_num)
                    try:
                        c.value = float(val)
                    except (ValueError, TypeError):
                        c.value = val
                    c.number_format = '0.00'
            written += 1

        print(f"    - [{source_sheet}] 已写入 {written} 列")
