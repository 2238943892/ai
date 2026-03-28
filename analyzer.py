import pandas as pd
import gradio as gr
import tempfile
import os
from pathlib import Path
import matplotlib.pyplot as plt

# ====== 列名适配 ======
ETAC_KEYS = ["Etac(%)", "Etac", "PCE", "Eff", "Efficiency"]
JSC_KEYS = ["Jsc(mA/cm²)", "Jsc(mA/cm2)", "Jsc"]
VOC_KEYS = ["Voc(V)", "Voc"]
FF_KEYS  = ["Fill Factor(%)", "FF", "FillFactor"]

def find_column(columns, candidates):
    for c in columns:
        for key in candidates:
            if key.lower() in c.lower():
                return c
    return None

# ====== 自动读取CSV / Excel ======
def read_file_auto(filepath):
    ext = Path(filepath).suffix.lower()

    if ext == ".csv":
        for enc in ["utf-8", "gbk", "gb2312", "latin-1"]:
            try:
                temp = pd.read_csv(filepath, encoding=enc, header=None)
                for i, row in temp.iterrows():
                    row_str = " ".join(row.astype(str))
                    if any(k in row_str for k in ["Etac", "PCE", "Eff"]):
                        return pd.read_csv(filepath, encoding=enc, skiprows=i)
            except:
                continue
    else:
        try:
            return pd.read_excel(filepath)
        except:
            return None

    return None

# ====== 处理单文件 ======
def process_file(filepath):
    df = read_file_auto(filepath)
    if df is None or df.empty:
        return None

    df.columns = df.columns.astype(str)

    etac_col = find_column(df.columns, ETAC_KEYS)
    jsc_col  = find_column(df.columns, JSC_KEYS)
    voc_col  = find_column(df.columns, VOC_KEYS)
    ff_col   = find_column(df.columns, FF_KEYS)

    if etac_col is None:
        return None

    df[etac_col] = pd.to_numeric(df[etac_col], errors='coerce')
    df = df.dropna(subset=[etac_col])

    if df.empty:
        return None

    max_row = df.loc[df[etac_col].idxmax()]

    return {
        "File": Path(filepath).name,
        "Jsc": max_row.get(jsc_col, None),
        "Voc": max_row.get(voc_col, None),
        "Fill Factor": max_row.get(ff_col, None),
        "Etac": max_row[etac_col]
    }

# ====== 自动分析总结 ======
def generate_summary(df):
    if df.empty:
        return "没有有效数据"

    best = df.iloc[0]
    avg = df["Etac"].mean()

    text = f"""
📊 数据分析总结：

- 🥇 最优样品：{best['File']}
- ⚡ 最高 Etac：{best['Etac']:.3f}
- 📉 平均 Etac：{avg:.3f}

👉 结论：
该批次样品整体性能{'较高' if avg > 10 else '一般'}，最优器件表现突出。
建议重点分析该样品的制备条件。
"""
    return text

# ====== 主逻辑 ======
def analyze(files):
    results = []

    for file in files:
        res = process_file(file.name)
        if res:
            results.append(res)

    if not results:
        return None, None, None, "没有有效数据"

    df = pd.DataFrame(results)

    # 排序 + Top10
    df = df.sort_values(by="Etac", ascending=False).head(10)

    # ====== 画图 ======
    plt.figure()
    plt.bar(df["File"], df["Etac"])
    plt.xticks(rotation=45, ha='right')

    img_path = os.path.join(tempfile.gettempdir(), "plot.png")
    plt.tight_layout()
    plt.savefig(img_path)
    plt.close()

    # ====== 保存CSV ======
    csv_path = os.path.join(tempfile.gettempdir(), "result.csv")
    df.to_csv(csv_path, index=False)

    # ====== 分析总结 ======
    summary = generate_summary(df)

    return df, img_path, csv_path, summary

# ====== UI ======
with gr.Blocks() as demo:
    gr.Markdown("# 🦞 龙虾AI Pro版 - 光伏数据分析工具")

    file_input = gr.File(file_count="multiple", label="上传 CSV / Excel 文件")
    btn = gr.Button("开始分析")

    output_table = gr.Dataframe(label="🏆 Top10 结果")
    output_plot = gr.Image(label="📊 Etac 分布图")
    output_file = gr.File(label="📥 下载结果")
    output_text = gr.Textbox(label="🤖 AI分析结论")

    btn.click(
        fn=analyze,
        inputs=file_input,
        outputs=[output_table, output_plot, output_file, output_text]
    )

demo.launch()
