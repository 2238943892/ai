import pandas as pd
import gradio as gr
import os
import io
from pathlib import Path

# ── 配置 ──────────────────────────────────────────────
REPORT_FOLDER = "./reports"
# ─────────────────────────────────────────────────────

def load_data_robustly(filepath):
    """鲁棒地读取 CSV 或 Excel，尝试跳过元数据行找到表头"""
    ext = Path(filepath).suffix.lower()
    
    if ext == '.csv':
        for encoding in ['utf-8', 'gbk', 'latin-1', 'utf-16']:
            try:
                # 尝试读取并寻找表头所在行（前 20 行找核心列名）
                df_raw = pd.read_csv(filepath, encoding=encoding, sep=None, engine='python', on_bad_lines='skip')
                # 清理空白行列
                df_raw = df_raw.dropna(how='all').dropna(axis=1, how='all')
                
                # 查找包含 Etac 或 PCE 的潜在行
                for i in range(min(20, len(df_raw))):
                    row_values = df_raw.iloc[i].astype(str).tolist()
                    if any(k in "".join(row_values) for k in ["Etac", "etac", "PCE", "Eff"]):
                        # 重新读入，以该行为表头
                        df = pd.read_csv(filepath, encoding=encoding, sep=None, engine='python', skiprows=i, on_bad_lines='skip')
                        return df
                return df_raw
            except:
                continue
    else:
        try:
            return pd.read_excel(filepath)
        except:
            return None
    return None

def extract_max_etac_data(file_objs):
    if not file_objs:
        return None, "请先上传文件（支持 .xlsx, .xls, .csv）"
        
    results = []
    
    for file_obj in file_objs:
        filepath = file_obj.name
        filename = Path(filepath).name
        
        try:
            df = load_data_robustly(filepath)
            if df is None or df.empty:
                continue
            
            # 统一清理列名
            df.columns = df.columns.astype(str).str.strip()
            cols = df.columns
            
            # 智能模糊匹配钙钛矿参数
            etac_col = next((c for c in cols if any(k in c for k in ["Etac", "etac", "PCE", "Eff"])), None)
            jsc_col = next((c for c in cols if any(k in c for k in ["Jsc", "jsc", "JSC"])), None)
            voc_col = next((c for c in cols if any(k in c for k in ["Voc", "voc", "VOC"])), None)
            ff_col = next((c for c in cols if any(k in c for k in ["Fill Factor", "FF", "ff", "FF(%)"])), None)
            
            if etac_col:
                df[etac_col] = pd.to_numeric(df[etac_col], errors='coerce')
                df = df.dropna(subset=[etac_col])
                
                if not df.empty:
                    max_idx = df[etac_col].idxmax()
                    best_row = df.loc[max_idx]
                    
                    results.append({
                        "文件名": filename,
                        "Etac(%)": best_row[etac_col],
                        "Jsc(mA/cm²)": best_row[jsc_col] if jsc_col else "未找到",
                        "Voc(V)": best_row[voc_col] if voc_col else "未找到",
                        "Fill Factor(%)": best_row[ff_col] if ff_col else "未找到"
                    })
        except:
            continue

    if not results:
        return None, "无法识别到数据。请确认：1.表头包含 'Etac' 2.数据行非空"
        
    summary_df = pd.DataFrame(results).sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    
    Path(REPORT_FOLDER).mkdir(parents=True, exist_ok=True)
    out_file = Path(REPORT_FOLDER) / "参数提取汇总表.xlsx"
    summary_df.to_excel(out_file, index=False)
    
    return summary_df, str(out_file)

with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("# 🦞 器件参数极速提取器")
    gr.Markdown("批量全选你的数据拖进来，会自动挑选出每份文件效率最高的一行数据。")
    with gr.Row():
        file_input = gr.File(label="📥 批量上传文件 (支持 .csv / .xlsx)", file_count="multiple")
    with gr.Row():
        submit_btn = gr.Button("开始提取并排序", variant="primary")
    with gr.Row():
        output_table = gr.Dataframe(label="📊 提取结果 (按效率降序)")
    with gr.Row():
        output_file = gr.File(label="📤 下载 Excel 汇总表")
        
    submit_btn.click(fn=extract_max_etac_data, inputs=file_input, outputs=[output_table, output_file])

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", 7860)))
