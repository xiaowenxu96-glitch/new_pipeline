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
    def _process_indicators_for_date_col(context, date_col_str, indicators, data_start_row):
        """处理一组指标，写入指定的日期列。"""
        ws = context['ws']
        reader = context['data_reader']
        source_sheet = context['sheet_config']['source_sheet']
        start_row = data_start_row
        date_col_num = col_letter_to_num(date_col_str)

        # 收集所有指标数据
        all_dates = set()
        code_dfs = {}
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
            code_dfs[aa_code] = (df, value_col)
            all_dates.update(df['日期'])

        # 即便没有指标，也把日期列清一下（用户可能后续手动填）
        if not indicators or not code_dfs:
            return 0, 0

        date_list = sorted(all_dates, reverse=True)

        # 写入日期列
        for i, date_val in enumerate(date_list):
            row = start_row + i
            c = ws.cell(row=row, column=date_col_num)
            c.value = date_val.to_pydatetime() if isinstance(date_val, pd.Timestamp) else date_val
            c.number_format = 'yyyy-mm-dd'

        # 写入各指标值
        for aa_code, target_col in indicators.items():
            pair = code_dfs.get(aa_code)
            if pair is None:
                continue
            df, value_col = pair
            value_map = dict(zip(df['日期'], df[value_col]))
            tgt_col_num = col_letter_to_num(target_col)
            for i, date_val in enumerate(date_list):
                val = value_map.get(date_val)
                if val is not None:
                    c = ws.cell(row=start_row + i, column=tgt_col_num)
                    c.value = float(val)
                    c.number_format = '0.00'

        return len(date_list), len(indicators)

    @staticmethod
    def chemical_write_sheet(context, params):
        ws = context['ws']
        reader = context['data_reader']
        sheet_config = context['sheet_config']
        defaults = context['defaults']

        source_sheet = sheet_config['source_sheet']
        start_row = sheet_config.get('data_start_row') or sheet_config.get('start_row') or defaults.get('start_row', 44)

        # ---- sections 模式（推荐：每个 section 独立指定 date_col）----
        sections = sheet_config.get('sections')
        if sections:
            total_rows = 0
            total_indicators = 0
            for section in sections:
                sec_date_col = section['date_col']
                sec_indicators = section.get('indicators') or {}
                sec_start = section.get('data_start_row') or start_row
                rows, n_ind = ChemicalPlugin._process_indicators_for_date_col(
                    context, sec_date_col, sec_indicators, sec_start
                )
                total_rows += rows
                total_indicators += n_ind
                if n_ind > 0:
                    print(f"    - [{source_sheet}] section [{sec_date_col}] 写入 {rows} 行, {n_ind} 个指标")
            print(f"    - [{source_sheet}] 共写入 {len(sections)} sections, {total_rows} 行, {total_indicators} 个指标")
            return

        # ---- 旧格式兼容（单个 date_col + indicators）----
        indicators = sheet_config['indicators']
        date_col = sheet_config.get('date_col')

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
            if date_col:
                date_col_num = col_letter_to_num(date_col)
            else:
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

    @staticmethod
    def chemical_write_formulas(context, params):
        """写入公式列 + 静态单元格公式，完全由 YAML 配置控制。

        params:
            start_row:      数据起始行，默认 44
            anchor_cell:    B列首行锚定的目标，如 "AG4"
            reference_col:  用于自动检测数据行数的列，默认 "AD"
            data_end_col:   VLOOKUP 搜索范围的最右列，如 "AO"
            bottom_row:     VLOOKUP 搜索范围底行（可选，默认 ws.max_row）
            columns:        公式列配置列表，每个元素:
                - col: "I"                    # 写入的目标列
                  formula: "=B{row}"          # 公式模板，{row} → 行号
                  format: "mm-dd-yy"          # 可选，单元格格式
                  lookup:                     # 可选，生成 VLOOKUP
                      start_col: "AD"         #   搜索范围起始列
                      return_index: 2         #   VLOOKUP 第几列
            date_col:       日期公式列（默认 "B"），写入锚定+逐行减7
            date_format:    日期列格式（默认 "yyyy/m/d;@"）
            static_cells:   固定单元格公式列表（用于顶部汇总表），每个元素:
                - cell: "E6"                  # 单元格地址，如 E6
                  formula: "=C44"             # 公式内容
                  format: "0.00"              # 可选，单元格格式
        """
        ws = context['ws']
        sheet_config = context['sheet_config']
        defaults = context['defaults']
        source_sheet = sheet_config['source_sheet']

        start_row = params.get('start_row') or defaults.get('start_row', 44)
        anchor_cell = params.get('anchor_cell', '')
        columns = params.get('columns') or []
        data_end_col = params.get('data_end_col', 'AF')
        bottom_row = params.get('bottom_row') or ws.max_row
        static_cells = params.get('static_cells') or []

        # ---- 静态单元格公式（顶部汇总表）----
        for sc in static_cells:
            cell_addr = sc['cell']
            cell = ws[cell_addr]
            if 'value' in sc:
                # 静态值（日期等），解析日期字符串为 datetime 对象
                from datetime import datetime
                val = sc['value']
                try:
                    cell.value = datetime.strptime(val, '%Y-%m-%d')
                except (ValueError, TypeError, KeyError):
                    cell.value = val
                if 'format' in sc:
                    cell.number_format = sc['format']
                print(f"    - [静态] {cell_addr} = {sc['value']}")
            else:
                # 公式
                formula = sc['formula']
                if 'is_array' in sc:
                    from openpyxl.worksheet.formula import ArrayFormula
                    cell.value = ArrayFormula(sc['is_array'], formula)
                else:
                    cell.value = formula
                if 'format' in sc:
                    cell.number_format = sc['format']
                print(f"    - [静态] {cell_addr} = {formula}")

        # ---- 公式行组（顶部汇总表逐行公式，同名占位符自动替换）----
        # 格式: [{ rows: 6..22, cols: "E|F|G|H|...", template: "=...", row_ref: "44" }]
        for fr in params.get('formula_rows', []):
            rows_range = fr['rows']
            cols_str = fr['cols']
            col_list = cols_str.split('|')
            template = fr['template']
            # expand {row_ref+n} placeholders
            def expand_template(tpl, base_row):
                import re
                result = tpl
                # {r+n} -> actual row number
                result = re.sub(r'\{r\+(\d+)\}', lambda m: str(base_row + int(m.group(1))), result)
                # {r} -> base_row
                result = result.replace('{r}', str(base_row))
                # {R} -> row of this formula cell
                result = re.sub(r'\{R\+(\d+)\}', lambda m: str(base_row + int(m.group(1))), result)
                result = result.replace('{R}', str(base_row))
                return result

            if isinstance(rows_range, str) and '..' in rows_range:
                parts = rows_range.split('..')
                start_r = int(parts[0])
                end_r = int(parts[1])
                row_list = list(range(start_r, end_r + 1))
            else:
                row_list = rows_range if isinstance(rows_range, list) else [rows_range]

            for formula_row in row_list:
                expanded = expand_template(template, formula_row)
                for col_letter in col_list:
                    cell = ws[f'{col_letter}{formula_row}']
                    cell.value = expanded
                    if 'format' in fr:
                        cell.number_format = fr['format']
            print(f"    - [公式行] rows {rows_range}, cols [{cols_str}], template: {template[:60]}...")

        if not columns:
            print(f"    - [{source_sheet}] 无公式列配置，跳过列写入")
            return

        # 自动检测数据行数
        ref_col = params.get('reference_col', 'AD')
        ref_num = col_letter_to_num(ref_col)
        last_data_row = start_row
        for r in range(start_row, ws.max_row + 1):
            if ws.cell(row=r, column=ref_num).value is not None:
                last_data_row = r
        formula_count = last_data_row - start_row + 1

        # ---- 日期列：锚定 + 逐行减7 ----
        date_col_cfg = params.get('date_col', 'B')
        date_col_num = col_letter_to_num(date_col_cfg)
        date_fmt = params.get('date_format', 'yyyy/m/d;@')
        for i in range(formula_count):
            row = start_row + i
            cell = ws.cell(row=row, column=date_col_num)
            if i == 0:
                cell.value = f'={anchor_cell}'
            else:
                cell.value = f'={date_col_cfg}{row-1}-7'
            cell.number_format = date_fmt
        print(f"    - [{source_sheet}] {date_col_cfg}列写入 {formula_count} 行 [锚定: ={anchor_cell}]")

        # ---- 各公式列 ----
        for col_cfg in columns:
            target_col = col_cfg['col']
            target_num = col_letter_to_num(target_col)
            cell_fmt = col_cfg.get('format', 'General')
            lookup = col_cfg.get('lookup')

            for i in range(formula_count):
                row = start_row + i
                cell = ws.cell(row=row, column=target_num)

                if lookup:
                    # VLOOKUP 模式
                    l_start = lookup['start_col']
                    l_idx = lookup['return_index']
                    l_end = lookup.get('end_col', data_end_col)
                    l_key = lookup.get('key_col', date_col_cfg)
                    cell.value = f'=VLOOKUP({l_key}{row},{l_start}{row}:{l_end}{bottom_row},{l_idx},FALSE)'
                else:
                    # 自由公式模式，{row} → 实际行号
                    formula = col_cfg['formula'].replace('{row}', str(row))
                    cell.value = formula

                cell.number_format = cell_fmt

            desc = f'VLOOKUP {lookup["start_col"]}:{data_end_col}[{lookup["return_index"]}]' if lookup else col_cfg.get('formula', '')
            print(f"    - [{source_sheet}] {target_col}列写入 {formula_count} 行 [{desc}]")

        print(f"    - [{source_sheet}] 共写入 {len(columns)+1} 个公式列, {formula_count} 行")