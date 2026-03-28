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
        
        ext = filename.lower()
        if not (ext.endswith('.xlsx') or ext.endswith('.xls') or ext.endswith('.csv')):
            continue
            
        try:
            # 根据格式读取数据
            if ext.endswith('.csv'):
                try:
                    df = pd.read_csv(filepath, encoding='utf-8')
                except:
                    df = pd.read_csv(filepath, encoding='gbk')
            else:
                df = pd.read_excel(filepath)
            
            # 统一清理列名
            df.columns = df.columns.astype(str).str.strip()
            cols = df.columns
            
            # 智能匹配钙钛矿电池关键参数列
            etac_col = next((c for c in cols if any(k in c for k in ["Etac", "etac", "PCE", "Eff"])), None)
            jsc_col = next((c for c in cols if any(k in c for k in ["Jsc", "jsc", "JSC"])), None)
            voc_col = next((c for c in cols if any(k in c for k in ["Voc", "voc", "VOC"])), None)
            ff_col = next((c for c in cols if any(k in c for k in ["Fill Factor", "FF", "ff", "FF(%)"])), None)
            
            if etac_col:
                # 强制转换效率为数字
                df[etac_col] = pd.to_numeric(df[etac_col], errors='coerce')
                df = df.dropna(subset=[etac_col])
                
                if not df.empty:
                    # 核心逻辑：锁定 Etac(%) 最大的一行
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
        return None, "未能识别到有效数据。请确认：1.表头包含 'Etac' 2.数据行非空"
        
    # 按 Etac(%) 从高到低排序
    summary_df = pd.DataFrame(results)
    summary_df = summary_df.sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    
    # 保存结果到云端
    Path(REPORT_FOLDER).mkdir(parents=True, exist_ok=True)
    output_filename = Path(REPORT_FOLDER) / "器件参数排序汇总表.xlsx"
    summary_df.to_excel(output_filename, index=False)
    
    return summary_df, str(output_filename)

# ── 网页界面布局 (请确保缩进严格一致) ──────────────────
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("# 🦞 器件参数极速提取器 (支持 Excel & CSV)")
    gr.Markdown("专门针对钙钛矿电池数据优化。只需把测试导出的整个文件夹文件拖进来，我会自动挑出效率最高的一行并排序。")
    
    with gr.Row():
        file_input = gr.File(label="📥 批量上传文件 (可全选拖入 .csv 或 .xlsx)", file_count="multiple")
    
    with gr.Row():
        submit_btn = gr.Button("开始极速分析", variant="primary")
        
    with gr.Row():
        output_table = gr.Dataframe(label="📊 提取结果 (按效率降序排列)")
        
    with gr.Row():
        output_file = gr.File(label="📤 下载整理好的 Excel 汇总表")
        
    submit_btn.click(
        fn=extract_max_etac_data,
        inputs=file_input,
        outputs=[output_table, output_file]
    )

if __name__ == "__main__":
    port = int(os.getenv("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
