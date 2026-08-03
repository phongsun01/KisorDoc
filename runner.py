"""
runner.py
─────────
Khởi động KisorDoc: Gradio UI (port 7864) + FastAPI (port 8000) song song.

Chạy:
    python runner.py

Dừng:
    Ctrl+C
"""

from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser


def _find_free_port(start: int, end: int = None) -> int:
    end = end or start + 50
    for port in range(start, end):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(("127.0.0.1", port))
            s.close()
            return port
        except OSError:
            continue
    raise RuntimeError(f"Không tìm thấy port trống trong dải {start}–{end}")


def start_fastapi(port: int):
    import uvicorn
    uvicorn.run(
        "api:app",
        host="0.0.0.0",
        port=port,
        log_level="warning",
        access_log=False,
    )


def start_gradio(port: int):
    from app import create_ui
    demo = create_ui()
    demo.queue()
    demo.launch(
        server_name="0.0.0.0",
        server_port=port,
        share=False,
        show_error=True,
        quiet=True,
        inbrowser=False,
    )


def main():
    gradio_port = _find_free_port(int(os.environ.get("GRADIO_PORT", 7864)))
    api_port    = _find_free_port(int(os.environ.get("API_PORT",    8000)))

    print("========================================")
    print("         KisorDoc dang khoi dong        ")
    print("----------------------------------------")
    print(f"  Gradio UI : http://127.0.0.1:{gradio_port}")
    print(f"  FastAPI   : http://127.0.0.1:{api_port}")
    print(f"  API Docs  : http://127.0.0.1:{api_port}/docs")
    print("========================================")

    # Start FastAPI trong thread daemon
    t_api = threading.Thread(
        target=start_fastapi, args=(api_port,), daemon=True, name="fastapi"
    )
    t_api.start()

    # Mở trình duyệt sau 2 giây (đủ thời gian Gradio khởi động)
    threading.Thread(
        target=lambda: (time.sleep(2), webbrowser.open(f"http://127.0.0.1:{gradio_port}")),
        daemon=True,
    ).start()

    # Start Gradio trong main thread (blocking)
    try:
        start_gradio(gradio_port)
    except KeyboardInterrupt:
        print("\nĐang dừng KisorDoc...")
        sys.exit(0)


if __name__ == "__main__":
    main()
