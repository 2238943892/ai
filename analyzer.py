import pandas as pd
import gradio as gr
import os
from pathlib import Path

# ── 配置 ──────────────────────────────────────────────
REPORT_FOLDER = "./reports"
# ─────────────────────────────────────────────────────

def extract_max_etac_data(file_objs):
    if not file_objs:
        return None, "请先上传文件（支持 .xlsx, .xls, .csv）"
        
    results = []
    
    for file_obj in file_objs:
        filepath = file_obj.name
        filename = Path(filepath).name
        
        # 1. 扩展检查逻辑：增加对 .csv 的支持
        ext = filename.lower()
        if not (ext.endswith('.xlsx') or ext.endswith('.xls') or ext.endswith('.csv')):
            continue
            
        try:
            # 2. 根据文件后缀选择不同的读取方式
            if ext.endswith('.csv'):
                try:
                    df = pd.read_csv(filepath, encoding='utf-8')
                except:
                    # 针对某些测试软件导出的 CSV 可能使用 GBK 编码
                    df = pd.read_csv(filepath, encoding='gbk')
            else:
                df = pd.read_excel(filepath)
            
            # 清理列名（去掉空格和特殊换行符）
            df.columns = df.columns.astype(str).str.strip()
            cols = df.columns
            
            # 3. 智能匹配列名：针对钙钛矿电池参数进行优化
            # 优先找包含 "Etac" 的，找不到再找包含 "PCE" 或 "Eff" 的
            etac_col = next((c for c in cols if any(k in c for k in ["Etac", "etac", "PCE", "Eff"])), None)
            jsc_col = next((c for c in cols if any(k in c for k in ["Jsc", "jsc", "JSC"])), None)
            voc_col = next((c for c in cols if any(k in c for k in ["Voc", "voc", "VOC"])), None)
            ff_col = next((c for c in cols if any(k in c for k in ["Fill Factor", "FF", "ff", "FF(%)"])), None)
            
            if etac_col:
                # 确保该列是数字类型
                df[etac_col] = pd.to_numeric(df[etac_col], errors='coerce')
                # 删掉数值为空的行
                df = df.dropna(subset=[etac_col])
                
                if not df.empty:
                    # 找到 Etac(%) 最大的那一行
                    max_idx = df[etac_col].idxmax()
                    best_row = df.loc[max_idx]
                    
                    results.append({
                        "文件名": filename,
                        "Etac(%)": best_row[etac_col],
                        "Jsc(mA/cm²)": best_row[jsc_col] if jsc_col else "未找到",
                        "Voc(V)": best_row[voc_col] if voc_col else "未找到",
                        "Fill Factor(%)": best_row[ff_col] if ff_col else "未找到"
                    })
            
        except Exception as e:
            print(f"处理 {filename} 失败: {e}")
            continue

    if not results:
        return None, "未能识别到有效数据。请检查：1.文件是否有内容 2.表头是否有 'Etac' 等字样。"
        
    # 4. 排序并生成结果
    summary_df = pd.DataFrame(results)
    summary_df = summary_df.sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    
    # 保存汇总表
    Path(REPORT_FOLDER).mkdir(parents=True, exist_ok=True)
    output_filename = Path(REPORT_FOLDER) / "参数提取汇总表.xlsx"
    summary_df.to_excel(output_filename, index=False)
    
    return summary_df, str(output_filename)

# 构建网页界面
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("# 🦞 器件参数极速提取器 (支持 Excel & CSV)")
    gr.Markdown("专门针对钙钛矿电池数据优化。只需把测试导出的整个文件夹文件拖进来，我会自动挑出效率最高的一行并排序。")
    
    with gr.Row():
        file_input = gr.File(label="📥 批量上传文件 (可全选拖入 .csv 或 .xlsx)", file_count="multiple")
    
    with gr.Row():
