"""
run_gui.py
----------
Unified Entry Point to launch the Cancer Research TME Algorithm Graphical Interface.

Usage:
  python run_gui.py          # Launches the modern interactive Web GUI (Default, browser-based)
  python run_gui.py --web    # Explicitly launches the Web GUI (Gradio)
  python run_gui.py --desktop# Launches the native Windows Desktop GUI (Tkinter)
"""

import sys
import argparse

def main():
    parser = argparse.ArgumentParser(description="Cancer Research Algorithm GUI Launcher")
    parser.add_argument('--desktop', action='store_true', help="Launch native Tkinter desktop interface")
    parser.add_argument('--web', action='store_true', help="Launch Gradio web interface (Default)")
    parser.add_argument('--port', type=int, default=7860, help="Port for web interface (Default: 7860)")
    args = parser.parse_args()

    if args.desktop:
        print("Launching Native Desktop Interface (Tkinter)...")
        from app_gui import main as run_desktop
        run_desktop()
    else:
        print("Launching Interactive Web Interface (Gradio)...")
        from web_gui import create_app
        demo = create_app()
        demo.launch(server_name="127.0.0.1", server_port=args.port, inbrowser=True)

if __name__ == '__main__':
    main()
