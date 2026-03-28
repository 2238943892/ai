import pandas as pd
import gradio as gr
import os
import io
from pathlib import Path

# 确保存放报告的文件夹存在
REPORT_FOLDER = "./reports"
os.makedirs(REPORT_FOLDER, exist_ok=True)

def process_single_file(filepath):
    """尝试多种方式读取并寻找效率最高的一行"""
    filename = Path(filepath).name
    ext = Path(filepath).suffix.lower()
    
    df = None
    # 1. 尝试读取文件
    if ext == '.csv':
        for enc in ['utf-8', 'gbk', 'gb2312', 'latin-1']:
            try:
                # 针对测试仪 CSV，先不设表头，读取全部内容
                temp_df = pd.read_csv(filepath, encoding=enc, sep=None, engine='python', on_bad_lines='skip', header=None)
                if not temp_df.empty:
                    # 寻找包含效率关键字的行作为表头
                    for i, row in temp_df.iterrows():
                        row_str = "".join(row.astype(str))
                        if any(k in row_str for k in ["Etac", "PCE", "Eff"]):
                            df = pd.read_csv(filepath, encoding=enc, sep=None, engine='python', skiprows=i)
                            break
                if df is not None: break
            except: continue
    else:
        try: df = pd.read_excel(filepath)
        except: return None

    if df is None or df.empty: return None

    # 2. 清理列名并匹配参数
    df.columns = df.columns.astype(str).str.strip()
    cols = df.columns
    
    # 效率匹配：Etac, PCE, Eff
    e_col = next((c for c in cols if any(k in c for k in ["Etac", "PCE", "Eff"])), None)
    # 电流匹配：Jsc
    j_col = next((c for c in cols if "Jsc" in c or "JSC" in c), None)
    # 电压匹配：Voc
    v_col = next((c for c in cols if "Voc" in c or "VOC" in c), None)
    # 填充因子匹配：FF, Fill Factor
    f_col = next((c for c in cols if any(k in c for k in ["FF", "Fill Factor", "ff"])), None)

    if e_col:
        df[e_col] = pd.to_numeric(df[e_col], errors='coerce')
        df = df.dropna(subset=[e_col])
        if not df.empty:
            best = df.loc[df[e_col].idxmax()]
            return {
                "文件名": filename,
                "Etac(%)": best[e_col],
                "Jsc(mA/cm²)": best[j_col] if j_col else "N/A",
                "Voc(V)": best[v_col] if v_col else "N/A",
                "Fill Factor(%)": best[f_col] if f_col else "N/A"
            }
    return None

def main_handler(file_objs):
    if not file_objs: return None, "未收到文件"
    
    final_results = []
    for f in file_objs:
        res = process_single_file(f.name)
        if res: final_results.append(res)
    
    if not final_results:
        return None, "❌ 处理结束：在上传的文件中没找到包含 'Etac' 或 'PCE' 列名的有效数据。"

    # 排序并导出
    summary_df = pd.DataFrame(final_results).sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    out_path = os.path.join(REPORT_FOLDER, "钙钛矿参数排序汇总.xlsx")
    summary_df.to_excel(out_path, index=False)
    
    return summary_df, out_path

# 网页界面：极简稳定版
with gr.Blocks() as demo:
    gr.Markdown("### 🦞 钙钛矿电池数据极速排序工具")
    with gr.Row():
        files = gr.File(label="📥 拖入 CSV 或 Excel (支持批量)", file_count="multiple")
    btn = gr.Button("🚀 开始分析并排序", variant="primary")
    with gr.Row():
        table = gr.Dataframe(label="📊 效率排名表")
    with gr.Row():
        download = gr.File(label="📤 下载汇总结果")
    
    btn.click(fn=main_handler, inputs=files, outputs=[table, download])

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", 7860)))
