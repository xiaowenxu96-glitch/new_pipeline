import yaml
import os
import shutil
import openpyxl
from datetime import datetime
from core_engine.data_reader import DataReader
from core_engine.transformers import apply_formula
from plugins.aviation_plugin import AviationPlugin
from plugins.macro_plugin import MacroPlugin
from plugins.electronics_plugin import ElectronicsPlugin
from plugins.baijiu_plugin import BaijiuPlugin
from plugins.metal_plugin import MetalPlugin

class PipelineEngine:
    def __init__(self, config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
            
        self.source_file = self.config['source_file']
        self.target_file = self.config['target_file']
        self.output_dir = self.config.get('output_dir', './output')
        
        # 动作注册表
        self.actions_registry = {
            'apply_formula': apply_formula,
            # 航空动作
            'aviation_write_airline_sheet': AviationPlugin.aviation_write_airline_sheet,
            'aviation_apply_yoy_formulas': AviationPlugin.aviation_apply_yoy_formulas,
            'aviation_apply_yoy_diff_formulas': AviationPlugin.aviation_apply_yoy_diff_formulas,
            'aviation_apply_yoy19_formulas': AviationPlugin.aviation_apply_yoy19_formulas,
            'aviation_apply_diff19_formulas': AviationPlugin.aviation_apply_diff19_formulas,
            'aviation_clear_early_years_data': AviationPlugin.aviation_clear_early_years_data,
            'aviation_adjust_format_after_full_year': AviationPlugin.aviation_adjust_format_after_full_year,
            'aviation_clear_ytd_2018_diff_data': AviationPlugin.aviation_clear_ytd_2018_diff_data,
            'aviation_clear_ax_ay_az_columns': AviationPlugin.aviation_clear_ax_ay_az_columns,
            # 宏观动作
            'write_indicator_group': MacroPlugin.macro_write_indicator_group,
            'create_pivot_table': MacroPlugin.macro_create_pivot_table,
            'macro_create_festival_pivot': MacroPlugin.macro_create_festival_pivot,
            'macro_create_chuxi_pivot': MacroPlugin.macro_create_chuxi_pivot,
            'macro_create_weekly_pivot': MacroPlugin.macro_create_weekly_pivot,
            'macro_create_yearly_date_scaffold': MacroPlugin.macro_create_yearly_date_scaffold,
            # 电子动作
            'electronics_write_sheet': ElectronicsPlugin.electronics_write_sheet,
            'electronics_update_chart_ranges': ElectronicsPlugin.electronics_update_chart_ranges,
            # 白酒动作
            'baijiu_write_sheet': BaijiuPlugin.baijiu_write_sheet,
            'baijiu_finalize_charts': BaijiuPlugin.baijiu_finalize_charts,
            # 有色金属动作
            'metal_write_sheet': MetalPlugin.metal_write_sheet,
        }
        
    def _create_backup(self):
        os.makedirs(self.output_dir, exist_ok=True)
        file_name = os.path.basename(self.target_file)
        name, ext = os.path.splitext(file_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(self.output_dir, f"{name}_{timestamp}{ext}")
        shutil.copy2(self.target_file, backup_file)
        print(f"[{self.config['industry']}] 已创建目标文件副本: {backup_file}")
        return backup_file
        
    def run(self):
        print(f"\n=============================================")
        print(f">> 开始执行 Pipeline: {self.config['industry']}")
        print(f"=============================================\n")
        
        # 1. 预加载所有数据到内存 (DataReader自带缓存机制)
        reader = DataReader(self.source_file)
        
        # 2. 创建输出副本并在内存中打开
        backup_file = self._create_backup()
        wb = openpyxl.load_workbook(backup_file)
        
        # 3. 遍历并执行每个 sheet 的任务
        defaults = self.config.get('defaults', {})
        for sheet_config in self.config['sheets']:
            sheet_name = sheet_config['sheet_name']
            print(f"\n>> 开始处理工作表: [{sheet_name}]")
            
            if sheet_name not in wb.sheetnames:
                print(f"!! 警告：模板中不存在工作表 {sheet_name}，跳过。")
                continue
                
            ws = wb[sheet_name]
            
            # 构建上下文环境
            context = {
                'wb': wb,
                'ws': ws,
                'data_reader': reader,
                'sheet_config': sheet_config,
                'defaults': defaults
            }
            
            # 执行配置中的 actions
            for action_cfg in sheet_config.get('actions', []):
                action_type = action_cfg['type']
                if action_type in self.actions_registry:
                    # 混合默认参数
                    params = {**defaults, **action_cfg}
                    self.actions_registry[action_type](context, params)
                else:
                    print(f"!! 未知动作: {action_type}")
                    
            # 执行后处理 post_processes
            for post_cfg in sheet_config.get('post_processes', []):
                action_type = post_cfg['action']
                if action_type in self.actions_registry:
                    params = post_cfg.get('params', {})
                    self.actions_registry[action_type](context, params)
                else:
                    print(f"⚠️ 未知后处理动作: {action_type}")
                    
        # 4. 一次性保存结果
        print(f"\n[Saving] 正在保存文件...")
        wb.save(backup_file)

        # 5. 执行 finalize_actions（在 save 之后，用于修复 openpyxl 的序列化问题）
        for finalize_cfg in self.config.get('finalize_actions', []):
            action_type = finalize_cfg['action']
            if action_type in self.actions_registry:
                fctx = {'filepath': backup_file, 'config': self.config}
                self.actions_registry[action_type](fctx, finalize_cfg.get('params', {}))
            else:
                print(f"⚠️ 未知 finalize 动作: {action_type}")

        print(f"[OK] Pipeline 执行完成！文件已保存至: {backup_file}\n")
