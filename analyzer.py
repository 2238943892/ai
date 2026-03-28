import pandas as pd
import gradio as gr
import os
import io
from pathlib import Path

# 确保存放汇总表的目录存在
REPORT_FOLDER = "reports"
os.makedirs(REPORT_FOLDER, exist_ok=True)

def process_single_file(filepath):
    """
    专门针对钙钛矿 J-V 原始 CSV/Excel 的解析函数
    逻辑：全表扫描带有 "Etac" 关键字的行，提取其下一行数据
    """
    filename = Path(filepath).name
    ext = Path(filepath).suffix.lower()
    
    file_results = []
    
    # 尝试多种编码读取文件，解决乱码问题
    for enc in ['utf-8', 'gbk', 'gb2312', 'latin-1', 'cp1252']:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                lines = f.readlines()
            
            # 逐行扫描寻找汇总表头
            for i, line in enumerate(lines):
                # 寻找包含核心参数的表头行
                if "Etac(%)" in line and "Jsc" in line:
                    header = line.strip().split(',')
                    if i + 1 < len(lines):
                        data_row = lines[i+1].strip().split(',')
                        
                        # 建立列名映射（防止不同软件列顺序不同）
                        col_map = {}
                        for idx, col in enumerate(header):
                            c = col.strip()
                            if "Etac" in c: col_map['Etac'] = idx
                            if "Jsc" in c: col_map['Jsc'] = idx
                            if "Voc" in c: col_map['Voc'] = idx
                            if "Fill Factor" in c or "FF" in c: col_map['FF'] = idx
                        
                        # 确保找到了效率列并提取数据
                        if 'Etac' in col_map and len(data_row) > max(col_map.values()):
                            try:
                                etac_val = float(data_row[col_map['Etac']])
                                file_results.append({
                                    "文件名": filename,
                                    "Etac(%)": etac_val,
                                    "Jsc(mA/cm²)": data_row[col_map['Jsc']] if 'Jsc' in col_map else "N/A",
                                    "Voc(V)": data_row[col_map['Voc']] if 'Voc' in col_map else "N/A",
                                    "Fill Factor(%)": data_row[col_map['FF']] if 'FF' in col_map else "N/A"
                                })
                            except: continue
            
            if file_results:
                # 从该文件的所有 Repeat 中找出 Etac 最大的那一个
                return max(file_results, key=lambda x: x["Etac(%)"])
            
            # 如果没找到汇总行，尝试普通表格读取方式（兜底逻辑）
            if not file_results:
                df = pd.read_csv(filepath, encoding=enc, on_bad_lines='skip')
                # 此处省略普通读取逻辑，可根据需要补充
        except:
            continue
    return None

def main_handler(file_objs):
    if not file_objs:
        return None, None, "⚠️ 请先上传文件。"
    
    summary_list = []
    for f in file_objs:
        res = process_single_file(f.name)
        if res:
            summary_list.append(res)
    
    if not summary_list:
        return None, None, "❌ 匹配失败：未能从文件中提取到 Etac 数据。请确保文件是原始测试报告。"

    # 1. 整理结果并按 Etac(%) 从高到低排序
    summary_df = pd.DataFrame(summary_list)
    summary_df["Etac(%)"] = pd.to_numeric(summary_df["Etac(%)"])
    summary_df = summary_df.sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    
    # 2. 保存 Excel 文件
    out_path = os.path.abspath(os.path.join(REPORT_FOLDER, "钙钛矿参数排序汇总.xlsx"))
    summary_df.to_excel(out_path, index=False)
    
    return summary_df, out_path, f"✅ 成功！已从 {len(summary_list)} 个文件中提取到最优数据。"

# 网页界面：极简稳定版
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("### ☀️ 钙钛矿器件参数极速排序 (J-V 原始文件版)")
    gr.Markdown("直接把测试导出的所有 CSV/Excel 丢进来。我会自动跳过原始曲线数据，只提取汇总表里效率最高的一行。")
    
    with gr.Row():
        files = gr.File(label="📥 批量上传文件", file_count="multiple")
    
    btn = gr.Button("🚀 提取最高效率并排序", variant="primary")
    
    status_msg = gr.Textbox(label="状态提示", interactive=False)
    
    with gr.Row():
        table = gr.Dataframe(label="📊 效率排名表 (Top Result per File)")
        download = gr.File(label="📤 下载汇总 Excel")
    
    btn.click(fn=main_handler, inputs=files, outputs=[table, download, status_msg])

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", 7860)))
