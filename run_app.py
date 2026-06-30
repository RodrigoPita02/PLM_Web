import os
import subprocess
import sys
import threading
import time
import socket
import webbrowser
import tkinter as tk
from tkinter import ttk

# 🔒 impedir múltiplas execuções
if os.environ.get("APP_ALREADY_RUNNING") == "1":
    sys.exit()

os.environ["APP_ALREADY_RUNNING"] = "1"


def run(root):
    print("🔄 A iniciar Streamlit...")

    proc = subprocess.Popen([
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "app.py",
        "--server.headless=true",
        "--server.port=8501",
        "--server.runOnSave=false"
    ])

    print("⏳ A aguardar servidor (pode demorar alguns minutos)...")

    while True:
        # ❗ se o processo morrer → erro
        if proc.poll() is not None:
            print("❌ Streamlit falhou ao arrancar")
            return

        # ✅ tenta conectar à porta
        try:
            with socket.create_connection(("localhost", 8501), timeout=2):
                break
        except:
            time.sleep(1)

    print("✅ Servidor pronto!")
    webbrowser.open("http://localhost:8501")
    root.destroy()


def main():
    root = tk.Tk()
    root.title("Visteon - SAP Validator")
    root.geometry("420x200")
    root.configure(bg="white")
    root.resizable(False, False)

    # 🟠 Header
    header = tk.Frame(root, bg="#F37021", height=50)
    header.pack(fill="x")

    tk.Label(
        header,
        text="Visteon SAP Validator",
        bg="#F37021",
        fg="white",
        font=("Segoe UI", 12, "bold")
    ).pack(pady=10)

    # Texto
    label = tk.Label(
        root,
        text="A iniciar aplicação (pode demorar alguns minutos)...",
        bg="white",
        fg="#333",
        font=("Segoe UI", 10)
    )
    label.pack(pady=15)

    # ✅ Barra loading correta
    style = ttk.Style()
    style.theme_use('default')

    style.configure(
        "Orange.Horizontal.TProgressbar",
        troughcolor='#EEEEEE',
        background='#F37021',
        thickness=6
    )

    progress = ttk.Progressbar(
        root,
        style="Orange.Horizontal.TProgressbar",
        mode="indeterminate"
    )
    progress.pack(fill="x", padx=40)
    progress.start()

    # thread (não bloqueia UI)
    threading.Thread(target=run, args=(root,), daemon=True).start()

    root.mainloop()


# ✅ obrigatório
if __name__ == "__main__":
    main()
