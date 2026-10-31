"""
run_gui.py
----------
Web GUI Launcher for the Cancer Research TME Surprisal Analysis & 
Exact Sign Sequence (ES) Mining Platform.

Usage:
  python run_gui.py          # Launches the interactive browser-based Web GUI
  python run_gui.py --port 8080 # Custom port (default: 7860)
"""

import sys
import argparse
import gradio as gr
from web_gui import create_web_application, HEAD_SCRIPT, CUSTOM_CSS

def main():
    parser = argparse.ArgumentParser(description="Cancer Research TME Algorithm Web GUI")
    parser.add_argument('--port', type=int, default=7860, help="Port to run the web server on (Default: 7860)")
    parser.add_argument('--host', type=str, default="127.0.0.1", help="Host address (Default: 127.0.0.1)")
    parser.add_argument('--share', action='store_true', help="Generate a public shareable Gradio link")
    args = parser.parse_args()

    print(f"====================================================================")
    print(f" Cancer Research TME Algorithm Workbench (Web GUI)")
    print(f" Server URL: http://{args.host}:{args.port}/?__theme=light")
    print(f" Git Source: https://github.com/natiBirhauz/cancer-research-final-project.git")
    print(f"====================================================================")

    app = create_web_application()
    app.launch(
        server_name=args.host,
        server_port=args.port,
        share=args.share,
        inbrowser=True,
        css=CUSTOM_CSS,
        head=HEAD_SCRIPT,
        theme=gr.themes.Default(primary_hue="neutral", secondary_hue="neutral", neutral_hue="neutral")
    )

if __name__ == '__main__':
    main()
