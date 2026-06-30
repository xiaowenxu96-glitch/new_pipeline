import pandas as pd
from openpyxl.utils import column_index_from_string as col_letter_to_num


class ChemicalPlugin:
    """化工 Pipeline 插件：按 AA 指标代码读取源数据并写入目标文档"""

    @staticmethod
    def _find_indicator_date_column(ws, start_row, indicator_col_num):
        """向左扫描 header 行，找到最近的'日期'列号"""
        for c in range(indicator_col_num - 1, 0, -1):
            v = ws.cell(row=start_row - 1, column=c).value
            if v and str(v).strip() == '日期':
                return c
        return indicator_col_num - 1  # fallback

    @staticmethod
    def chemical_write_sheet(context, params):
        ws = context['ws']
        reader = context['data_reader']
        sheet_config = context['sheet_config']
        defaults = context['defaults']

        source_sheet = sheet_config['source_sheet']
        start_row = sheet_config.get('data_start_row') or sheet_config.get('start_row') or defaults.get('start_row', 44)
        indicators = sheet_config['indicators']

        # 按日期列分组: date_col_num → {dates, [(tgt_col_num, date_map), ...]}
        groups = {}
        for aa_code, target_col in indicators.items():
            ind_data = reader.read_indicator_data(source_sheet, aa_code)
            df = ind_data['data']
            if df.empty:
                continue
            df = df[df['日期'].notna()].copy()
            value_col = df.columns[-1]
            df = df[df[value_col].notna()].copy()
            if df.empty:
                continue

            tgt_col_num = col_letter_to_num(target_col)
            date_col_num = ChemicalPlugin._find_indicator_date_column(ws, start_row, tgt_col_num)

            date_val_map = {}
            for _, row in df.iterrows():
                d = row['日期']
                v = row[value_col]
                if pd.notna(d) and pd.notna(v):
                    date_val_map[d] = float(v)

            if date_col_num not in groups:
                groups[date_col_num] = {'dates': set(), 'indicators': []}
            groups[date_col_num]['dates'].update(date_val_map.keys())
            groups[date_col_num]['indicators'].append((tgt_col_num, date_val_map))

        if not groups:
            print(f"    - [{source_sheet}] 未找到任何数据")
            return

        total_rows = 0
        for date_col_num, group in groups.items():
            date_list = sorted(group['dates'], reverse=True)
            for i, date_val in enumerate(date_list):
                row = start_row + i
                c = ws.cell(row=row, column=date_col_num)
                c.value = date_val.to_pydatetime() if isinstance(date_val, pd.Timestamp) else date_val
                c.number_format = 'yyyy-mm-dd'
            for tgt_col_num, date_map in group['indicators']:
                for i, date_val in enumerate(date_list):
                    val = date_map.get(date_val)
                    if val is not None:
                        c = ws.cell(row=start_row + i, column=tgt_col_num)
                        c.value = val
                        c.number_format = '0.00'
            total_rows += len(date_list)

        print(f"    - [{source_sheet}] 写入 {len(groups)} 组日期列, {total_rows} 行, {len(indicators)} 个指标")
