import pandas as pd
import os
from datetime import datetime, timedelta
import sys

def update_excel_fees(df):
    """
    整合脚本1：修改收费标准
    输入：DataFrame
    输出：修改后的DataFrame
    """
    print("[步骤1] 正在执行：修改收费标准...")
    
    # 确保必要的列存在
    col_area = '管理区'
    col_fee = '收费标准'
    if col_area not in df.columns or col_fee not in df.columns:
        raise ValueError(f"错误：Excel 文件中未找到必需的列 '{col_area}' 或 '{col_fee}'。")

    # 定义修改规则字典
    rules = { 
        # (管理区, 原收费标准) -> 新收费标准
        ("白晶谷15-2组团_A区", "高层物业费_2.2元（BJG_WYF_GC_2.2）"): "高层物业费",
        ("白晶谷26组团_A区", "高层物业费_2元（BJG_WYF_GC_2）"): "高层物业费",
        ("白晶谷26组团_A区", "别墅物业费_3.2元（BJG_WFY_BS_3.2）"): "别墅物业费",
        ("白晶谷32组团_A区", "商品物业费_4元（BJG_SP_WYF_4元）"): "商铺物业费",
        ("白晶谷32组团_A区", "叠拼物业费_2.6元（BJG_WYF_DP_2.6）"): "叠拼物业费",
        ("白晶谷32组团_A区", "合院物业费_3.9元（BJG_WYF_HY_3.9）"): "联排合院物业费",
        ("白晶谷32组团_A区", "联排物业费_3.9元（BJG_WFY_LP_3.9）"): "联排合院物业费",
        
        ("太阳谷13组团_A区", "商铺物业费_4元（TYG_WYF_SP_4）"): "商铺物业费",
        ("太阳谷13组团_A区", "洋房物业费_2.6元（TYG_WYF_YF_2.6）"): "高层物业费",
        ("太阳谷13组团_B区", "高层物业费_2.6元（TYG_WYF_GC_2.6）"): "高层物业费",
        ("太阳谷13组团_B区", "合院物业费_4元（TYG_WFY_HY_4）"): "别墅物业费",
        ("太阳谷13组团_B区", "合院物业费_太阳谷_4元（TYG_WFY_HY_N_4）"): "别墅物业费",
        ("太阳谷13组团_B区", "别墅物业费_4元（TYG_WFY_BS_4）"): "别墅物业费",
        ("太阳谷13组团_B区", "别墅物业费_太阳谷_4元（TYG_WFY_BS_N_4）"): "别墅物业费",
        
        ("太阳谷30组团_A区", "高层物业费_2.6元（TYG_WYF_GC_2.6）"): "高层物业费",
        ("太阳谷30组团_A区", "洋房物业费_2.6元（TYG_WYF_YF_2.6）"): "高层物业费",
        ("太阳谷30组团_B区", "商铺物业费_4元（TYG_WYF_SP_4）"): "商铺物业费",
        ("太阳谷30组团_B区", "高层物业费_2.6元（TYG_WYF_GC_2.6）"): "高层物业费",
        
        ("太阳谷7组团_A区", "商铺物业费_3.9元（TYG_WYF_SP_3.9）"): "商铺物业费",
        ("太阳谷7组团_A区", "高层物业费_2元（TYG_WYF_GC_2）"): "高层物业费",
        ("太阳谷7组团_B区", "别墅物业费_3.5元（TYG_WFY_BS_3.5）"): "别墅物业费",
        ("太阳谷7组团_G区", "高层物业费_2元（TYG_WYF_GC_2）"): "高层物业费",
        ("太阳谷7组团_F区", "商铺物业费_3.9元（TYG_WYF_SP_3.9）"): "商铺物业费",
        ("太阳谷7组团_F区", "高层物业费_2元（TYG_WYF_GC_2）"): "高层物业费",
        
        ("悦龙东郡一组团_二期", "高层物业费_2元（WYF_GC_2）"): "高层物业费",
        ("悦龙东郡一组团_二期", "商铺物业费_4元（WYF_SP_4）"): "商铺物业费",
        ("悦龙东郡一组团_二期", "商铺物业费_2元（WYF_SP_2）"): "商铺物业费",
        ("悦龙东郡一组团_三期", "高层物业费_2元（WYF_GC_2）"): "高层物业费",
        ("悦龙东郡一组团_三期", "商铺物业费_4元（WYF_SP_4）"): "商铺物业费",
        ("悦龙东郡一组团_四期", "高层物业费_2元（WYF_GC_2）"): "高层物业费",
        ("悦龙东郡一组团_五期", "商铺物业费_4元（WYF_SP_4）"): "商铺物业费",
        ("悦龙东郡一组团_五期", "高层物业费_2元（WYF_GC_2）"): "高层物业费",
        ("悦龙东郡一组团_六期A区", "商铺物业费_4元（WYF_SP_4）"): "商铺物业费",
        ("悦龙东郡一组团_六期A区", "高层物业费_2元（WYF_GC_2）"): "高层物业费",
        
        ("悦龙东郡二组团_A区", "高层物业费_2元（WYF_GC_2）"): "高层物业费",
        ("悦龙东郡二组团_C区", "高层物业费_2元（WYF_GC_2）"): "高层物业费",
        
        ("悦龙南山5组团_A区", "高层物业费_2元（YLNS_WYF_GC_2）"): "高层物业费",
        ("悦龙南山5组团_A区", "商铺物业费_4元（YLNS_WYF_SP_4）"): "商铺物业费",
    }
    
    count_updated = 0
    for index, row in df.iterrows():
        area = row[col_area]
        fee = row[col_fee]
        
        # 处理可能的空白字符
        if isinstance(area, str):
            area = area.strip()
        if isinstance(fee, str):
            fee = fee.strip()
            
        key = (area, fee)
        if key in rules:
            df.at[index, col_fee] = rules[key]
            count_updated += 1
            
    print(f"[步骤1] 完成，共修改了 {count_updated} 条记录。")
    return df

def modify_management_area(df):
    """
    整合脚本2：修改管理区字段
    输入：DataFrame
    输出：修改后的DataFrame
    """
    print("[步骤2] 正在执行：修改管理区字段...")
    
    # 自动检测管理区列名
    management_col = None
    for col in df.columns:
        if '管理区' in col:
            management_col = col
            break
            
    if management_col is None:
        raise ValueError("错误：未找到包含'管理区'的列名。")

    # 定义修改规则
    def apply_modifications(area):
        if pd.isna(area):
            return area
            
        # 规则1：太阳谷7组团系列
        if area in ['太阳谷7组团_A区', '太阳谷7组团_B区', '太阳谷7组团_F区', '太阳谷7组团_G区']:
            return '太阳谷7组团'
            
        # 规则2：太阳谷13组团系列
        if area in ['太阳谷13组团_A区', '太阳谷13组团_B区']:
            return '太阳谷13组团'
            
        # 规则3：太阳谷30组团系列
        if area in ['太阳谷30组团_A区', '太阳谷30组团_B区']:
            return '太阳谷30组团'
            
        # 规则4：白晶谷系列
        if area == '白晶谷15-2组团_A区':
            return '白晶谷15-2组团'
        if area == '白晶谷26组团_A区':
            return '白晶谷26组团'
        if area == '白晶谷32组团_A区':
            return '白晶谷32组团'
            
        # 规则5：悦龙东郡一组团系列
        if area in ['悦龙东郡一组团_二期', '悦龙东郡一组团_三期', '悦龙东郡一组团_四期', 
                   '悦龙东郡一组团_五期', '悦龙东郡一组团_六期A区']:
            return '悦龙东郡一组团'
            
        # 规则6：悦龙东郡二组团系列
        if area in ['悦龙东郡二组团_A区', '悦龙东郡二组团_C区']:
            return '悦龙东郡二组团'
            
        # 如果没有匹配到规则，则返回原始值
        return area

    # 应用修改规则
    df[management_col] = df[management_col].apply(apply_modifications)
    print("[步骤2] 完成。")
    return df

def analyze_excel_data(df):
    """
    整合脚本3：统计分析
    输入：DataFrame
    输出：保存统计结果
    """
    print("[步骤3] 正在执行：统计分析...")
    
    # 检查必要列是否存在
    required_columns = ['组织机构', '管理区', '收费标准', '收全日期', '已收金额']
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"缺少必要的列: {missing_cols}")

    # --- 数据预处理 ---
    df['收全日期'] = pd.to_datetime(df['收全日期'], errors='coerce')
    df['已收金额'] = pd.to_numeric(df['已收金额'], errors='coerce').fillna(0)

    # 过滤用于常规统计的有效数据 (有收全日期)
    valid_data = df.dropna(subset=['收全日期'])
    
    # --- 时间范围计算 ---
    # 获取统计基准日期 (今天 00:00:00)，确保使用当前日期
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=4)
    weekday = today.weekday()
    
    # 1. 本周的开始日期
    start_of_this_week = today - timedelta(days=weekday)
    
    # 2. 上周范围：上周一 00:00:00 到 上周日 23:59:59
    start_of_last_week = start_of_this_week - timedelta(days=7)
    end_of_last_week = (start_of_this_week - timedelta(days=1)).replace(hour=23, minute=59, second=59)
    
    # 3. 上上周末：上上周日 23:59:59
    end_of_week_before_last = (start_of_last_week - timedelta(days=1)).replace(hour=23, minute=59, second=59)
    
    # 4. 本月范围：本月1号 00:00:00 到 本月最后一天 23:59:59
    start_of_current_month = today.replace(day=1)
    if today.month == 12:
        end_of_current_month = today.replace(year=today.year+1, month=1, day=1) - timedelta(seconds=1)
    else:
        end_of_current_month = today.replace(month=today.month+1, day=1) - timedelta(seconds=1)
    
    print("------------------")
    print("统计时间：")
    print(f"上一周：{start_of_last_week.strftime('%Y-%m-%d')}至{end_of_last_week.strftime('%Y-%m-%d')}")
    print(f"本月：{start_of_current_month.strftime('%Y-%m-%d')}至{end_of_current_month.strftime('%Y-%m-%d')}")
    print(f"累计（包含空白收全日期，并统计至上上周末）：2026年1月1日至{end_of_week_before_last.strftime('%Y-%m-%d')}")
    print("------------------")
    
    # --- 统计执行 ---

    # 统计1：上一周 (基于 已收金额)
    last_week_data = valid_data[(valid_data['收全日期'] >= start_of_last_week) & (valid_data['收全日期'] <= end_of_last_week)]
    last_week_summary = last_week_data.groupby(['组织机构', '管理区', '收费标准'])['已收金额'].sum().reset_index()
    last_week_summary['时间段'] = "本周累计收费"
    last_week_summary['具体日期'] = f"{start_of_last_week.strftime('%Y-%m-%d')}至{end_of_last_week.strftime('%Y-%m-%d')}"

    # 统计2：本月 (基于 已收金额)
    current_month_data = valid_data[(valid_data['收全日期'] >= start_of_current_month) & (valid_data['收全日期'] <= end_of_current_month)]
    current_month_summary = current_month_data.groupby(['组织机构', '管理区', '收费标准'])['已收金额'].sum().reset_index()
    current_month_summary['时间段'] = "月度累计收费"
    current_month_summary['具体日期'] = f"{start_of_current_month.strftime('%Y-%m-%d')}至{end_of_current_month.strftime('%Y-%m-%d')}"

    # 统计3：累计（包含空白收全日期，并统计至上上周末）
    # 条件1: 收全日期为空白
    blank_date_data = df[df['收全日期'].isna()]
    # 条件2: 收全日期在上上周末之前
    past_date_data = df[df['收全日期'] <= end_of_week_before_last]
    # 合并两种情况的数据，并去除重复项
    cumulative_data_to_process = pd.concat([blank_date_data, past_date_data]).drop_duplicates().reset_index(drop=True)
    
    period_cumulative_summary = cumulative_data_to_process.groupby(['组织机构', '管理区', '收费标准'])['已收金额'].sum().reset_index()
    period_cumulative_summary['时间段'] = "截止上周累计收费"
    period_cumulative_summary['具体日期'] = f"2026-01-01至{end_of_week_before_last.strftime('%Y-%m-%d')}"

    # --- 合并与保存 ---
    all_summary = pd.concat([
        last_week_summary, 
        current_month_summary, 
        period_cumulative_summary
    ], ignore_index=True)
    
    # 调整列顺序
    cols = ['组织机构', '管理区', '收费标准', '时间段', '具体日期', '已收金额']
    all_summary = all_summary[cols]
    all_summary = all_summary.sort_values(['组织机构', '管理区', '收费标准', '时间段']).reset_index(drop=True)
    
    # 保存结果
    output_file = f"整合统计结果_{today.strftime('%Y%m%d')}.xlsx"
    with pd.ExcelWriter(output_file) as writer:
        all_summary.to_excel(writer, sheet_name='汇总结果', index=False)
        last_week_summary.to_excel(writer, sheet_name='本周累计收费', index=False)
        current_month_summary.to_excel(writer, sheet_name='月度累计收费', index=False)
        period_cumulative_summary.to_excel(writer, sheet_name='截止上周累计收费', index=False)
        
    print(f"[步骤3] 完成，结果已保存至 {output_file}")
    return all_summary

def main():
    # 检查是否传入了参数
    if len(sys.argv) < 2:
        print("错误：脚本需要传入文件路径作为参数。")
        print("请在 SKILL.md 中检查调用逻辑。")
        return
    
    input_path = sys.argv[1] # 获取第一个参数（文件路径）
    print(f"正在处理文件: {input_path}")
    
    # 检查文件是否存在
    if not os.path.exists(input_path):
        print(f"错误：找不到文件 '{input_path}'")
        return
        
    try:
        # 读取原始数据
        df = pd.read_excel(input_path)
        print(f"成功读取 {len(df)} 行数据。")
        
        # 2. 执行清洗流程
        df = update_excel_fees(df)
        df = modify_management_area(df)
        
        # 3. 执行统计
        analyze_excel_data(df) # 注意：这个函数内部会自己保存文件
        
        print("\n=== 全部流程执行完毕！===")
        
    except Exception as e:
        print(f"程序执行出错: {str(e)}") # 增加了错误提示，方便调试

if __name__ == "__main__":
    main()