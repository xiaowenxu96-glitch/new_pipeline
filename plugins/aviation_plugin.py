import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string as column_letter_to_number, get_column_letter as column_number_to_letter

# 为了重用之前的复杂代码，直接引入旧的 DataProcessor (作为示例平滑迁移，实际可全盘重构进这里)
from core_engine.data_processor import DataProcessor 

class AviationPlugin:
    @staticmethod
    def _find_section_rows(ws):
        quarter_start_row = None
        ytd_row = None
        full_year_row = None
        last_label_row = 0

        for row_idx in range(1, ws.max_row + 1):
            raw_value = ws.cell(row=row_idx, column=1).value
            if raw_value is None:
                continue

            value = str(raw_value).strip()
            if not value:
                continue

            last_label_row = row_idx
            if value == "季度":
                quarter_start_row = row_idx
            elif "YTD" in value.upper():
                ytd_row = row_idx
            elif "全年" in value:
                full_year_row = row_idx

        return {
            "quarter_start_row": quarter_start_row,
            "ytd_row": ytd_row,
            "full_year_row": full_year_row,
            "last_label_row": last_label_row,
        }

    @staticmethod
    def _clear_columns_below_row(ws, cols_to_clear, start_row):
        if not start_row or start_row > ws.max_row:
            return

        for col in cols_to_clear:
            col_num = column_letter_to_number(col)
            for row_idx in range(start_row, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_num).value = None

    @staticmethod
    def _get_last_formula_row(ws):
        section_rows = AviationPlugin._find_section_rows(ws)
        return section_rows["last_label_row"], section_rows
    
    @staticmethod
    def aviation_write_airline_sheet(context, params):
        """
        处理基础指标数据写入，对齐到第一个指标日期，并生成季度和YTD统计。
        context: 包含 wb (openpyxl Workbook), ws, data_reader
        params: 从 YAML 中解析出的该 action 的配置
        """
        ws = context['ws']
        reader = context['data_reader']
        sheet_config = context['sheet_config']
        
        source_sheet = sheet_config['source_sheet']
        indicator_col_map = sheet_config['indicators']
        start_row = params.get('start_row', 4)
        start_date = params.get('start_date', '2014-01-01')
        date_format = params.get('date_format', 'yyyy-mm')
        
        # 1. 从缓存中获取所需所有指标数据
        indicator_dfs = {}
        for indicator in indicator_col_map:
            ind_data = reader.read_indicator_data(source_sheet, indicator)
            df = ind_data["data"]
            # 只保留start_date及以后数据
            valid_dates = pd.to_datetime(df['日期'], errors='coerce')
            df = df[valid_dates.notna() & (valid_dates >= start_date)].copy()
            df = df.sort_values('日期', ascending=True).reset_index(drop=True)
            indicator_dfs[indicator] = df

        # 2. 写入第一个指标的时间列和数据
        first_indicator = list(indicator_col_map.keys())[0]
        first_col = indicator_col_map[first_indicator]
        first_df = indicator_dfs[first_indicator]
        
        for i, row in first_df.iterrows():
            ws.cell(row=start_row + i, column=1, value=row['日期'])  # A列
            ws.cell(row=start_row + i, column=column_letter_to_number(first_col), value=row[first_df.columns[-1]])
            ws.cell(row=start_row + i, column=1).number_format = date_format

        # 3. 写入其他指标数据（不写时间列）
        for indicator, col in list(indicator_col_map.items())[1:]:
            df = indicator_dfs[indicator]
            # 对齐到第一个指标的日期
            date_list = first_df['日期'].tolist()
            value_map = dict(zip(df['日期'], df[df.columns[-1]]))
            for i, date in enumerate(date_list):
                value = value_map.get(date, None)
                ws.cell(row=start_row + i, column=column_letter_to_number(col), value=value)

        # 4. 添加季度统计（从2018年开始）
        empty_row = start_row + len(first_df)
        ws.cell(row=empty_row, column=1, value=None)
        
        title_row = empty_row + 1
        ws.cell(row=title_row, column=1, value="季度")
        ws.cell(row=title_row, column=1).font = openpyxl.styles.Font(bold=True)
        
        # 过滤2018年及以后的数据
        valid_dates = pd.to_datetime(first_df['日期'], errors='coerce')
        df_2018 = first_df[valid_dates.notna() & (valid_dates.dt.year >= 2018)].copy()
        df_2018.reset_index(drop=True, inplace=True)
        
        df_all_cols = pd.DataFrame({'日期': df_2018['日期']})
        for indicator, col in indicator_col_map.items():
            indicator_df = indicator_dfs[indicator]
            valid_ind_dates = pd.to_datetime(indicator_df['日期'], errors='coerce')
            filtered_df = indicator_df[valid_ind_dates.notna() & (valid_ind_dates.dt.year >= 2018)].copy()
            date_list = df_2018['日期'].tolist()
            value_map = dict(zip(filtered_df['日期'], filtered_df[filtered_df.columns[-1]]))
            df_all_cols[col] = [value_map.get(date, None) for date in date_list]

        processor = DataProcessor(reader)
        quarter_title_row, quarter_results = processor.add_quarterly_stats(
            ws=ws,
            start_row=start_row,
            data_df=df_all_cols,
            value_cols=list(indicator_col_map.values())
        )

        current_row = title_row + 1
        for quarter_data in quarter_results:
            if quarter_data['quarter_label']:
                ws.cell(row=current_row, column=1, value=quarter_data['quarter_label'])
                for col, value in quarter_data['data'].items():
                    cell = ws.cell(row=current_row, column=column_letter_to_number(col), value=value)
                    if value is not None:
                        original_col = ws.cell(row=start_row, column=column_letter_to_number(col))
                        cell.number_format = original_col.number_format
            current_row += 1

        last_quarter_row = current_row - 1
        while last_quarter_row > title_row:
            if ws.cell(row=last_quarter_row, column=1).value and 'Q' in str(ws.cell(row=last_quarter_row, column=1).value):
                break
            last_quarter_row -= 1

        last_ytd_row = processor.add_ytd_stats(
            ws=ws,
            last_quarter_row=last_quarter_row,
            data_df=df_all_cols,  
            value_cols=list(indicator_col_map.values())
        )

        target_decimal_cols = ['AL', 'AO', 'AR', 'AU']
        for col in target_decimal_cols:
            if col in indicator_col_map.values():
                col_num = column_letter_to_number(col)
                for row in range(start_row, last_ytd_row + 1):
                    cell = ws.cell(row=row, column=col_num)
                    if cell.value is not None:
                        cell.number_format = '0.00'
                        
        print(f"    - 完成航空基础数据写入与季度统计")

    @staticmethod
    def aviation_apply_yoy_formulas(context, params):
        ws = context['ws']
        target_cols = params['target_cols']
        start_row = params.get('start_row', 16)

        last_formula_row, section_rows = AviationPlugin._get_last_formula_row(ws)
        full_year_row = section_rows["full_year_row"]
        AviationPlugin._clear_columns_below_row(ws, target_cols, last_formula_row + 1)

        data_2017_rows = {}
        for row_idx in range(start_row, last_formula_row + 1):
            raw_value = ws.cell(row=row_idx, column=1).value
            cell_value = str(raw_value).strip() if raw_value is not None else ""
            if cell_value.startswith("2017-"):
                try:
                    month = int(cell_value.split("-")[1])
                    data_2017_rows[month] = row_idx
                except (ValueError, IndexError):
                    continue

        for col in target_cols:
            col_num = column_letter_to_number(col)
            source_col = column_number_to_letter(col_num - 1)
            
            for row in range(start_row, last_formula_row + 1):
                formula = ""
                current_cell = ws.cell(row=row, column=column_letter_to_number(source_col))
                raw_date_value = ws.cell(row=row, column=1).value
                date_value = str(raw_date_value).strip() if raw_date_value is not None else ""
                if not date_value or current_cell.value in [None, ""]:
                    continue
                
                if "年" in date_value:
                    if full_year_row is not None and row > full_year_row:
                        prev_row = row - 1
                        if prev_row >= 1:
                            formula = f"={source_col}{row}/{source_col}{prev_row}-1"
                    else:
                        try:
                            year = int(date_value.replace("年", ""))
                            if year <= 2018:
                                formula = ""
                            else:
                                prev_row = row - 2
                                if prev_row >= 1:
                                    formula = f"={source_col}{row}/{source_col}{prev_row}-1"
                        except (ValueError, IndexError):
                            prev_row = row - 2
                            if prev_row >= 1:
                                formula = f"={source_col}{row}/{source_col}{prev_row}-1"
                elif "18" in date_value and "Q" in date_value:
                    try:
                        quarter = int(date_value[date_value.find("Q") + 1])
                        q_month_rows = []
                        for month in range((quarter - 1) * 3 + 1, (quarter - 1) * 3 + 4):
                            if month in data_2017_rows:
                                q_month_rows.append(data_2017_rows[month])

                        if len(q_month_rows) == 3:
                            sum_formula = f"(SUM({source_col}{q_month_rows[0]}:{source_col}{q_month_rows[2]}))"
                            formula = f"={source_col}{row}/{sum_formula}-1"
                    except (ValueError, IndexError):
                        formula = ""
                elif "Q" in date_value:
                    prev_row = row - 5
                    if prev_row >= 1:
                        formula = f"={source_col}{row}/{source_col}{prev_row}-1"
                else:
                    prev_row = row - 12
                    if prev_row >= 1:
                        formula = f"={source_col}{row}/{source_col}{prev_row}-1"
                        
                if formula:
                    cell = ws.cell(row=row, column=col_num)
                    cell.value = formula
                    cell.number_format = "0.00%"
        print(f"    - 完成航空同比上年计算: {target_cols}")

    @staticmethod
    def aviation_apply_yoy_diff_formulas(context, params):
        ws = context['ws']
        target_cols = params['target_cols']
        start_row = params.get('start_row', 16)
        last_formula_row, _ = AviationPlugin._get_last_formula_row(ws)
        AviationPlugin._clear_columns_below_row(ws, target_cols, last_formula_row + 1)
            
        for col in target_cols:
            col_num = column_letter_to_number(col)
            source_col = column_number_to_letter(col_num - 1)
            for row in range(start_row, last_formula_row + 1):
                formula = ""
                raw_date_value = ws.cell(row=row, column=1).value
                date_value = str(raw_date_value).strip() if raw_date_value is not None else ""
                if not date_value:
                    continue
                if "年" in date_value:
                    if "2018年" in date_value:
                        continue
                    prev_row = row - 2
                    if prev_row >= 1:
                        formula = f"={source_col}{row}-{source_col}{prev_row}"
                else:
                    prev_row = row - 12
                    if prev_row >= 1:
                        formula = f"={source_col}{row}-{source_col}{prev_row}"
                if formula:
                    cell = ws.cell(row=row, column=col_num)
                    cell.value = formula
                    cell.number_format = "0.00"
        print(f"    - 完成航空同比差值计算: {target_cols}")

    @staticmethod
    def aviation_apply_yoy19_formulas(context, params):
        ws = context['ws']
        target_cols = params['target_cols']
        start_row = params.get('start_row', 112)
        last_formula_row, section_rows = AviationPlugin._get_last_formula_row(ws)
        full_year_row = section_rows["full_year_row"]
        AviationPlugin._clear_columns_below_row(ws, target_cols, last_formula_row + 1)

        for col in target_cols:
            col_num = column_letter_to_number(col)
            source_col = column_number_to_letter(col_num - 2)
            
            # 找到2019年相关行
            row_2019_base = 64
            row_2019 = None
            row_2019_after_full_year = None
            row_2019Q1 = row_2019Q2 = row_2019Q3 = row_2019Q4 = None
            full_year_row_19 = None
            
            for r in range(1, ws.max_row + 1):
                val = str(ws.cell(row=r, column=1).value)
                if val == "2019年":
                    if full_year_row_19 and r > full_year_row_19:
                        row_2019_after_full_year = r
                    else:
                        row_2019 = r
                elif val == "2019Q1": row_2019Q1 = r
                elif val == "2019Q2": row_2019Q2 = r
                elif val == "2019Q3": row_2019Q3 = r
                elif val == "2019Q4": row_2019Q4 = r
                elif "全年" in val: full_year_row_19 = r

            for row in range(start_row, last_formula_row + 1):
                formula = ""
                raw_date_value = ws.cell(row=row, column=1).value
                date_value = str(raw_date_value).strip() if raw_date_value is not None else ""
                if not date_value:
                    continue
                
                if "Q" in date_value and "2019" not in date_value:
                    quarter = date_value[-1]
                    q_row_map = {"1": row_2019Q1, "2": row_2019Q2, "3": row_2019Q3, "4": row_2019Q4}
                    row_19_q = q_row_map.get(quarter)
                    if row_19_q:
                        formula = f"={source_col}{row}/{source_col}{row_19_q}-1"
                elif "年" in date_value:
                    year_str = "".join(filter(str.isdigit, date_value))
                    if year_str and int(year_str) >= 2023:
                        if full_year_row and row > full_year_row and row_2019_after_full_year:
                            formula = f"={source_col}{row}/{source_col}{row_2019_after_full_year}-1"
                        elif row_2019:
                            formula = f"={source_col}{row}/{source_col}{row_2019}-1"
                elif "19" not in date_value:
                    month = None
                    try:
                        month = pd.to_datetime(date_value).month
                    except: pass
                    if month:
                        base_row = row_2019_base + month - 1
                        formula = f"={source_col}{row}/{source_col}{base_row}-1"

                if formula:
                    cell = ws.cell(row=row, column=col_num)
                    cell.value = formula
                    cell.number_format = "0.00%"
        print(f"    - 完成航空同比19计算: {target_cols}")

    @staticmethod
    def aviation_apply_diff19_formulas(context, params):
        ws = context['ws']
        target_cols = params['target_cols']
        start_row = params.get('start_row', 112)
        last_formula_row, section_rows = AviationPlugin._get_last_formula_row(ws)
        full_year_row = section_rows["full_year_row"]
        AviationPlugin._clear_columns_below_row(ws, target_cols, last_formula_row + 1)

        for col in target_cols:
            col_num = column_letter_to_number(col)
            source_col = column_number_to_letter(col_num - 2)
            
            row_2019_base = 64
            row_2019 = None
            row_2019_after_full_year = None
            row_2019Q1 = row_2019Q2 = row_2019Q3 = row_2019Q4 = None
            full_year_row_19 = None
            
            for r in range(1, ws.max_row + 1):
                val = str(ws.cell(row=r, column=1).value)
                if val == "2019年":
                    if full_year_row_19 and r > full_year_row_19:
                        row_2019_after_full_year = r
                    else:
                        row_2019 = r
                elif val == "2019Q1": row_2019Q1 = r
                elif val == "2019Q2": row_2019Q2 = r
                elif val == "2019Q3": row_2019Q3 = r
                elif val == "2019Q4": row_2019Q4 = r
                elif "全年" in val: full_year_row_19 = r

            for row in range(start_row, last_formula_row + 1):
                formula = ""
                raw_date_value = ws.cell(row=row, column=1).value
                date_value = str(raw_date_value).strip() if raw_date_value is not None else ""
                if not date_value:
                    continue
                
                if "Q" in date_value and "2019" not in date_value:
                    quarter = date_value[-1]
                    q_row_map = {"1": row_2019Q1, "2": row_2019Q2, "3": row_2019Q3, "4": row_2019Q4}
                    row_19_q = q_row_map.get(quarter)
                    if row_19_q:
                        formula = f"={source_col}{row}-{source_col}{row_19_q}"
                elif "年" in date_value:
                    year_str = "".join(filter(str.isdigit, date_value))
                    if year_str and int(year_str) >= 2023:
                        if full_year_row and row > full_year_row and row_2019_after_full_year:
                            formula = f"={source_col}{row}-{source_col}{row_2019_after_full_year}"
                        elif row_2019:
                            formula = f"={source_col}{row}-{source_col}{row_2019}"
                elif "19" not in date_value:
                    month = None
                    try:
                        month = pd.to_datetime(date_value).month
                    except: pass
                    if month:
                        base_row = row_2019_base + month - 1
                        formula = f"={source_col}{row}-{source_col}{base_row}"

                if formula:
                    cell = ws.cell(row=row, column=col_num)
                    cell.value = formula
                    cell.number_format = "0.00"
        print(f"    - 完成航空差值型同比19计算: {target_cols}")

    @staticmethod #清 2014–2017 年这些行里的同比/同比19列（让它们保持空白）。
    def aviation_clear_early_years_data(context, params):
        ws = context['ws']
        target_years = ["2014年", "2015年", "2016年", "2017年"]
        cols_to_clear = params.get('yoy_cols', []) + params.get('yoy19_cols', [])
        
        for row_idx in range(1, ws.max_row + 1):
            cell_value = str(ws.cell(row=row_idx, column=1).value)
            if cell_value in target_years:
                for col in cols_to_clear:
                    ws.cell(row=row_idx, column=column_letter_to_number(col)).value = None
        print(f"    - 完成清除2014-2017年早年无用公式")

    @staticmethod #全年行以下的数据格式调整(保留一位小数)
    def aviation_adjust_format_after_full_year(context, params):
        ws = context['ws']
        cols = params.get('yoy_cols', []) + params.get('yoy19_cols', [])
        
        full_year_row = None
        for row_idx in range(1, ws.max_row + 1):
            if "全年" in str(ws.cell(row=row_idx, column=1).value):
                full_year_row = row_idx
                break
                
        if full_year_row:
            for col in cols:
                col_num = column_letter_to_number(col)
                for row in range(full_year_row + 1, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_num)
                    if cell.value is not None:
                        cell.number_format = "0.0%"
        print(f"    - 完成全年行以下的数据格式调整(保留一位小数)")

    @staticmethod # 清除YTD中2018年的同比差值
    def aviation_clear_ytd_2018_diff_data(context, params):
        ws = context['ws']
        yoy_diff_cols = params.get('yoy_diff_cols', [])
        
        ytd_row = full_year_row = None
        for row_idx in range(1, ws.max_row + 1):
            val = str(ws.cell(row=row_idx, column=1).value)
            if "YTD" in val.upper(): ytd_row = row_idx
            elif "全年" in val: full_year_row = row_idx
            
        if ytd_row and full_year_row:
            for row_idx in range(ytd_row, full_year_row):
                if str(ws.cell(row=row_idx, column=1).value) == "2018年":
                    for col in yoy_diff_cols:
                        ws.cell(row=row_idx, column=column_letter_to_number(col)).value = None
        print(f"    - 完成清除YTD中2018年的同比差值")

    @staticmethod # 清除季度/YTD/全年数据的AX, AY, AZ列的残留（模板特殊列）
    def aviation_clear_ax_ay_az_columns(context, params):
        ws = context['ws']
        cols_to_clear = ["AX", "AY", "AZ"]
        
        quarter_start_row = ytd_row = full_year_row = None
        for row_idx in range(1, ws.max_row + 1):
            val = str(ws.cell(row=row_idx, column=1).value)
            if val == "季度": quarter_start_row = row_idx
            elif "YTD" in val.upper(): ytd_row = row_idx
            elif "全年" in val: full_year_row = row_idx
            
        for row_idx in range(1, ws.max_row + 1):
            val = str(ws.cell(row=row_idx, column=1).value)
            should_clear = False
            if quarter_start_row and ytd_row and quarter_start_row < row_idx < ytd_row and "Q" in val:
                should_clear = True
            elif ytd_row and full_year_row and ytd_row < row_idx < full_year_row and "年" in val:
                should_clear = True
            elif full_year_row and row_idx > full_year_row and "年" in val:
                should_clear = True
                
            if should_clear:
                for col in cols_to_clear:
                    cell = ws.cell(row=row_idx, column=column_letter_to_number(col))
                    # 无论原先有没有值，直接清除内容并把公式清空，防止出现 #NAME? 等错误
                    cell.value = None
        print(f"    - 完成清除季度/YTD/全年数据的AX, AY, AZ列")
