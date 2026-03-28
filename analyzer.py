import pandas as pd
import gradio as gr
import os
from pathlib import Path

# 确保存放汇总表的目录存在
REPORT_FOLDER = "reports"
os.makedirs(REPORT_FOLDER, exist_ok=True)

def load_data_flexibly(filepath):
    """智能寻找表头并读取数据 (支持 CSV/Excel)"""
    ext = Path(filepath).suffix.lower()
    df = None
    
    if ext == '.csv':
        # 尝试所有可能的编码，防止乱码
        for enc in ['utf-8', 'gbk', 'gb2312', 'latin-1', 'cp1252']:
            try:
                # 1. 先不设表头读入，寻找包含关键字的行
                temp = pd.read_csv(filepath, encoding=enc, header=None, sep=None, engine='python', on_bad_lines='skip')
                for i, row in temp.iterrows():
                    row_str = " ".join(row.astype(str))
                    # 匹配钙钛矿常见的效率关键字
                    if any(k in row_str for k in ["Etac", "PCE", "Eff"]):
                        # 2. 以这一行为表头重新读取
                        df = pd.read_csv(filepath, encoding=enc, skiprows=i, sep=None, engine='python', on_bad_lines='skip')
                        break
                if df is not None: break
            except: continue
    elif ext in ['.xlsx', '.xls']:
        try:
            df = pd.read_excel(filepath)
        except: pass
    return df

def analyze_files(file_objs):
    if not file_objs:
        return None, None, "⚠️ 请先上传文件。"
    
    all_data = []
    status_msg = ""
    
    for f in file_objs:
        filename = Path(f.name).name
        df = load_data_flexibly(f.name)
        
        if df is not None and not df.empty:
            # 清理列名空格
            df.columns = df.columns.astype(str).str.strip()
            cols = df.columns
            
            # 自动寻找对应的列名（模糊匹配）
            eff_col = next((c for c in cols if any(k in c for k in ["Etac", "PCE", "Eff"])), None)
            jsc_col = next((c for c in cols if any(k in c for k in ["Jsc", "jsc", "JSC"])), None)
            voc_col = next((c for c in cols if any(k in c for k in ["Voc", "voc", "VOC"])), None)
            ff_col = next((c for c in cols if any(k in c for k in ["FF", "Fill Factor", "ff"])), None)
            
            if eff_col:
                # 转换效率为数字
                df[eff_col] = pd.to_numeric(df[eff_col], errors='coerce')
                valid_df = df.dropna(subset=[eff_col])
                
                if not valid_df.empty:
                    # 找到 Etac(%) 最大的一行
                    max_idx = valid_df[eff_col].idxmax()
                    best_row = valid_df.loc[max_idx]
                    
                    all_data.append({
                        "文件名": filename,
                        "Etac(%)": best_row[eff_col],
                        "Jsc(mA/cm²)": best_row[jsc_col] if jsc_col else "N/A",
                        "Voc(V)": best_row[voc_col] if voc_col else "N/A",
                        "Fill Factor(%)": best_row[ff_col] if ff_col else "N/A"
                    })
        else:
            status_msg += f"无法解析: {filename}\n"

    if not all_data:
        return None, None, "❌ 错误：所有文件中都没找到 'Etac' 或 'PCE' 表头。请检查文件内容。"

    # 创建汇总表并排序
    summary_df = pd.DataFrame(all_data).sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    
    # 保存结果到本地文件供下载
    out_path = os.path.abspath(os.path.join(REPORT_FOLDER, "Summary_Result.xlsx"))
    summary_df.to_excel(out_path, index=False)
    
    success_msg = f"✅ 处理成功！已从 {len(all_data)} 个文件中提取到最高效率。"
    return summary_df, out_path, success_msg

# 构建网页界面
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("## ☀️ 钙钛矿器件参数全自动排序工具")
    gr.Markdown("支持批量拖入 CSV 或 Excel，自动锁定最高效率行并整理参数。")
    
    with gr.Row():
        file_input = gr.File(label="📥 上传测试文件 (可多选/全选)", file_count="multiple")
    
    run_btn = gr.Button("🚀 提取并按效率排序", variant="primary")
    
    status_output = gr.Textbox(label="运行状态", interactive=False)
    
    with gr.Row():
        table_output = gr.Dataframe(label="📊 提取结果预览")
        file_output = gr.File(label="📤 下载汇总 Excel 表")

    run_btn.click(
        fn=analyze_files, 
        inputs=file_input, 
        outputs=[table_output, file_output, status_output]
    )

if __name__ == "__main__":
    # Railway 部署环境变量
    port = int(os.getenv("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
