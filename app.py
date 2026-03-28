import gradio as gr
import os

def test():
    return "OK"

with gr.Blocks() as demo:
    gr.Markdown("# 测试页面")

    btn = gr.Button("点我")
    out = gr.Textbox()

    btn.click(fn=test, outputs=out)

demo.launch(
    server_name="0.0.0.0",
    server_port=int(os.environ.get("PORT", 7860))
)
