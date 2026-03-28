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
    逻辑：搜索含有 "Etac" 关键字的行，提取其下一行数据
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
                if "Etac(%)" in line and "Jsc" in line:
                    header = line.strip().split(',')
                    if i + 1 < len(lines):
                        data_row = lines[i+1].strip().split(',')
                        
                        # 建立列名映射
                        col_map = {}
                        for idx, col in enumerate(header):
                            c = col.strip()
                            if "Etac" in c: col_map['Etac'] = idx
                            if "Jsc" in c: col_map['Jsc'] = idx
                            if "Voc" in c: col_map['Voc'] = idx
                            if "Fill Factor" in c or "FF" in c: col_map['FF'] = idx
                        
                        # 提取数据，调整字典键的顺序：把 Etac(%) 放在最后
                        if 'Etac' in col_map and len(data_row) > max(col_map.values()):
                            try:
                                etac_val = float(data_row[col_map['Etac']])
                                file_results.append({
                                    "文件名": filename,
                                    "Jsc(mA/cm²)": data_row[col_map['Jsc']] if 'Jsc' in col_map else "N/A",
                                    "Voc(V)": data_row[col_map['Voc']] if 'Voc' in col_map else "N/A",
                                    "Fill Factor(%)": data_row[col_map['FF']] if 'FF' in col_map else "N/A",
                                    "Etac(%)": etac_val
                                })
                            except: continue
            
            if file_results:
                # 从该文件的所有测试轮次中找出 Etac 最大的
                return max(file_results, key=lambda x: x["Etac(%)"])
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
        return None, None, "❌ 未能提取到数据，请确保上传的是原始测试 CSV 文件。"

    # 1. 整理结果并按 Etac(%) 从大到小排序
    summary_df = pd.DataFrame(summary_list)
    # 确保排序逻辑正确（即便它在最后一列）
    summary_df = summary_df.sort_values(by="Etac(%)", ascending=False).reset_index(drop=True)
    
    # 2. 保存 Excel 文件
    out_path = os.path.abspath(os.path.join(REPORT_FOLDER, "钙钛矿器件参数提取结果.xlsx"))
    summary_df.to_excel(out_path, index=False)
    
    return summary_df, out_path, f"✅ 搞定！已从 {len(summary_list)} 个文件中提取到最优效率数据。"

# Gradio 界面
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("### ☀️ 钙钛矿器件参数全自动排序 (效率末尾版)")
    gr.Markdown("全选拖入 CSV 文件，自动锁定最高效率行。表格最后一列即为最终效率 $Etac(\%)$。")
    
    with gr.Row():
        files = gr.File(label="📥 批量上传文件", file_count="multiple")
    
    btn = gr.Button("🚀 提取最高效率并排序", variant="primary")
    
    status_msg = gr.Textbox(label="状态提示", interactive=False)
    
    with gr.Row():
        table = gr.Dataframe(label="📊 器件性能汇总表 (按效率降序)")
        download = gr.File(label="📤 下载汇总 Excel")
    
    btn.click(fn=main_handler, inputs=files, outputs=[table, download, status_msg])

if __name__ == "__main__":
    port = int(os.getenv("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
