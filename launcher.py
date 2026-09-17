import os
import sys
import threading
import webbrowser

from app import app

BUNDLE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

app.root_path = os.getcwd()
app.template_folder = os.path.join(BUNDLE_DIR, "templates")
app.static_folder = os.path.join(BUNDLE_DIR, "static")

threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
app.run(port=5000)
