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
        # 增加更多的编码尝试，防止钙钛矿测试仪导出的特殊格式乱码
        for enc in ['utf-8', 'gbk', 'gb2312', 'latin-1', 'cp1252']:
            try:
                # 针对测试仪 CSV，先不设表头，读取全部内容
                temp_df = pd.read_csv(filepath, encoding=enc, sep=None, engine='python', on_bad_lines='skip', header=None)
                if not temp_df.empty:
                    # 寻找包含效率关键字的行作为表头（全表搜索）
                    for i, row in temp_df.iterrows():
                        row_str = "".join(row.astype(str))
                        if any(k in row_str for k in ["Etac", "PCE", "Eff"]):
                            df = pd.read_csv(filepath, encoding=enc, sep=None, engine='python', skiprows=i)
                            break
                if df is not None: break
            except: continue
    else:
        try: 
            df = pd.read_excel(filepath)
        except: 
            return None

    if df is None or df.empty: return None

    # 2. 清理列名并匹配参数
    df.columns = df.columns.astype(str).str.strip()
    cols = df.columns
    
    # 智能匹配：效率、电流、电压、填充因子
    e_col = next((c for c in cols if any(k in c for k in ["Etac", "PCE", "Eff"])), None)
    j_col = next((c for c in cols if any(k in c for k in ["Jsc", "jsc", "JSC"])), None)
    v_col = next((c for c in cols if any(k in c for k in ["Voc", "voc", "VOC"])), None)
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
    if not file_objs: 
        return None, None, "⚠️ 未收到文件，请先上传文件。"
    
    final_results = []
    for f in file_objs:
        # Gradio 传进来的是对象，路径在 .name 里
        res = process_single_file(f.name)
        if res: final_results.append(res)
    
    if not final_results:
        return None, None, "❌ 匹配失败：在上传的文件中没找到包含 'Etac' 或 'PCE' 的列。请检查文件格式。"

    # 排序并导出
    summary_df = pd.DataFrame(final_results).sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    out_path = os.path.abspath(os.path.join(REPORT_FOLDER, "汇总排序结果.xlsx"))
    summary_df.to_excel(out_path, index=False)
    
    return summary_df, out_path, f"✅ 处理成功！共处理 {len(final_results)} 个器件数据。"

# 网页界面布局
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("## 🦞 钙钛矿器件数据极速提取与排序")
    gr.Markdown("将测试导出的 CSV 或 Excel 文件夹内容全选拖入，系统自动提取每份文件的最高效率行并降序排列。")
    
    with gr.Row():
        files = gr.File(label="📥 批量上传数据文件", file_count="multiple")
    
    btn = gr.Button("🚀 开始分析并排序", variant="primary")
    
    # 状态提示框：专门显示成功或报错信息，防止系统崩溃
    status_msg = gr.Textbox(label="运行状态", interactive=False)
    
    with gr.Row():
        table = gr.Dataframe(label="📊 效率排名预览")
    
    with gr.Row():
        download = gr.File(label="📤 下载汇总 Excel 表")
    
    # 这里的 outputs 顺序必须和 main_handler 返回值顺序一致
    btn.click(fn=main_handler, inputs=files, outputs=[table, download, status_msg])

if __name__ == "__main__":
    # Railway 部署必须监听 0.0.0.0 和环境变量中的 PORT
    demo.launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", 7860)))
