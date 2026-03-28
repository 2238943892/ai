import pandas as pd
import gradio as gr
import os
from pathlib import Path

# ── 配置 ──────────────────────────────────────────────
REPORT_FOLDER = "./reports"
# ─────────────────────────────────────────────────────

def extract_max_etac_data(file_objs):
    if not file_objs:
        return None, "请先上传 Excel 文件"
        
    results = []
    
    for file_obj in file_objs:
        filepath = file_obj.name
        filename = Path(filepath).name
        
        if not (filename.endswith('.xlsx') or filename.endswith('.xls')):
            continue
            
        try:
            # 读取 Excel 文件
            df = pd.read_excel(filepath)
            
            # 获取所有列名并转为字符串
            cols = df.columns.astype(str)
            
            # 智能匹配列名（忽略大小写和乱码符号）
            etac_col = next((c for c in cols if "Etac" in c or "etac" in c), None)
            jsc_col = next((c for c in cols if "Jsc" in c or "jsc" in c), None)
            voc_col = next((c for c in cols if "Voc" in c or "voc" in c), None)
            ff_col = next((c for c in cols if "Fill Factor" in c or "FF" in c or "ff" in c), None)
            
            if etac_col and df[etac_col].notna().any():
                # 确保该列是数字类型，遇到无法转换的转为空值
                df[etac_col] = pd.to_numeric(df[etac_col], errors='coerce')
                
                # 找到 Etac(%) 最大的那一行
                max_idx = df[etac_col].idxmax()
                best_row = df.loc[max_idx]
                
                results.append({
                    "文件名": filename,
                    "Etac(%)": best_row[etac_col] if etac_col else None,
                    "Jsc(mA/cm²)": best_row[jsc_col] if jsc_col else None,
                    "Voc(V)": best_row[voc_col] if voc_col else None,
                    "Fill Factor(%)": best_row[ff_col] if ff_col else None
                })
            else:
                # 如果没有找到 Etac 列
                results.append({
                    "文件名": filename,
                    "Etac(%)": -999,  # 用负数垫底
                    "Jsc(mA/cm²)": "未找到列",
                    "Voc(V)": "未找到列",
                    "Fill Factor(%)": "未找到列"
                })
                
        except Exception as e:
            print(f"处理 {filename} 失败: {e}")
            continue

    if not results:
        return None, "未能从上传的文件中提取到有效数据，请检查表头。"
        
    # 转换为数据框
    summary_df = pd.DataFrame(results)
    
    # 过滤掉没有 Etac 数据的文件，并按 Etac(%) 从大到小排序
    summary_df = summary_df[summary_df["Etac(%)"] != -999]
    summary_df = summary_df.sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    
    # 保存汇总后的 Excel 文件用于下载
    Path(REPORT_FOLDER).mkdir(parents=True, exist_ok=True)
    output_filename = Path(REPORT_FOLDER) / "Etac效率排序汇总表.xlsx"
    summary_df.to_excel(output_filename, index=False)
    
    return summary_df, str(output_filename)

# 构建网页界面
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("# ⚡️ 器件参数极速提取器 (按 Etac 排序)")
    gr.Markdown("批量上传 Excel 文件，自动提取每份文件中 `Etac(%)` 最高的一行数据（包含 Jsc, Voc, Fill Factor），并按效率从大到小生成汇总表。")
    
    with gr.Row():
        file_input = gr.File(label="📥 批量上传 Excel 文件 (可全选拖入)", file_count="multiple")
    
    with gr.Row():
        submit_btn = gr.Button("开始极速提取并排序", variant="primary")
        
    with gr.Row():
        output_table = gr.Dataframe(label="📊 数据提取结果 (已按 Etac 降序排列)")
        
    with gr.Row():
        output_file = gr.File(label="📤 下载完整 Excel 汇总表")
        
    submit_btn.click(
        fn=extract_max_etac_data,
        inputs=file_input,
        outputs=[output_table, output_file]
    )

if __name__ == "__main__":
    port = int(os.getenv("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
