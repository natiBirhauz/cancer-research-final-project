"""
run_gui.py
----------
Web GUI Launcher for the Surprisal Analysis (SA) – Minimal Exact Sign
Signatures (ES) Discovery Platform.

Usage:
  python run_gui.py             # Launches the interactive browser-based Web GUI
  python run_gui.py --port 8080 # Custom port (default: 7860)
"""

import argparse
import gradio as gr
from web_gui import create_web_application, HEAD_SCRIPT, CUSTOM_CSS


def main():
    parser = argparse.ArgumentParser(description="Surprisal Analysis (SA) – Minimal ES Signatures Discovery Platform")
    parser.add_argument('--port', type=int, default=7860, help="Port to run the web server on (Default: 7860)")
    parser.add_argument('--host', type=str, default="127.0.0.1", help="Host address (Default: 127.0.0.1)")
    parser.add_argument('--share', action='store_true', help="Generate a public shareable Gradio link")
    args = parser.parse_args()

    print("====================================================================")
    print(" Surprisal Analysis (SA) - Minimal ES Signatures Discovery Platform")
    print(f" Server URL: http://{args.host}:{args.port}/?__theme=light")
    print("====================================================================")

    app = create_web_application()
    app.launch(
        server_name=args.host,
        server_port=args.port,
        share=args.share,
        inbrowser=True,
        css=CUSTOM_CSS,
        head=HEAD_SCRIPT,
        theme=gr.themes.Default(primary_hue="blue", secondary_hue="slate", neutral_hue="slate")
    )


if __name__ == '__main__':
    main()
