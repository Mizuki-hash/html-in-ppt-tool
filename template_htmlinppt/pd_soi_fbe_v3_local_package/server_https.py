from pathlib import Path
import http.server
import json
import ssl
import sys

PORT = 8765
HOST = "127.0.0.1"
BASE_DIR = Path(__file__).resolve().parent
CERT_FILE = BASE_DIR / "localhost_selfsigned.crt"
KEY_FILE = BASE_DIR / "localhost_selfsigned.key"
CONFIG_FILE = BASE_DIR / "pd_soi_fbe_visual_config.json"

COLOR_KEYS = {
    "pageBg", "bgA", "bgB", "glowBlue", "glowAmber", "cardBg",
    "cardBorder", "boxBg", "boxBg2", "line", "text", "muted",
    "canvasBg", "chartBg", "accent", "accentSoft", "carrier",
    "carrierDark", "heat", "green", "substrate", "substrateStroke",
    "boxA", "boxB", "siliconA", "siliconB", "electrode",
    "electrodeStroke", "oxide", "gate", "chartGrid", "chartAxis",
    "bulkLine", "noteBg", "noteBorder",
}
NUMBER_LIMITS = {
    "fontScale": (0.82, 1.25),
    "lineScale": (0.65, 1.70),
    "particleScale": (0.35, 1.90),
    "leftWeight": (0.86, 1.28),
    "leftTopWeight": (0.55, 1.65),
    "rightTopWeight": (0.45, 1.55),
    "deviceOrder": (1, 3),
    "metricsOrder": (1, 3),
    "controlsOrder": (1, 3),
    "chartOrder": (1, 2),
    "historyOrder": (1, 2),
    "cardRadius": (4, 28),
    "boxRadius": (4, 24),
    "canvasRadius": (0, 18),
    "shadow": (0, 1),
    "glow": (0, 1),
}
BOOL_KEYS = {
    "showHeader", "showSubtitle", "showChips", "showFooter", "showDevice",
    "showMetrics", "showControls", "showChart", "showHistory", "showExplain",
}


def is_hex_color(value):
    if not isinstance(value, str) or len(value) != 7 or value[0] != "#":
        return False
    return all(ch in "0123456789abcdefABCDEF" for ch in value[1:])


def clean_config(data):
    clean = {}
    for key in COLOR_KEYS:
        value = data.get(key)
        if is_hex_color(value):
            clean[key] = value.lower()
    for key, (min_value, max_value) in NUMBER_LIMITS.items():
        value = data.get(key)
        if isinstance(value, (int, float)):
            clean[key] = max(min_value, min(max_value, float(value)))
    for key in BOOL_KEYS:
        value = data.get(key)
        if isinstance(value, bool):
            clean[key] = value
    return clean


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def do_POST(self):
        route = self.path.split("?", 1)[0].strip("/")
        if route != "save_pd_soi_fbe_visual_config":
            self.send_error(404, "Unknown endpoint")
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        if length <= 0 or length > 32768:
            self.send_error(400, "Invalid config size")
            return
        try:
            data = json.loads(self.rfile.read(length).decode("utf-8"))
            clean = clean_config(data)
            CONFIG_FILE.write_text(
                json.dumps(clean, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as e:
            self.send_error(400, f"Cannot save config: {e}")
            return
        payload = json.dumps({"ok": True}, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        super().end_headers()

    def log_message(self, fmt, *args):
        print("[%s] %s" % (self.log_date_time_string(), fmt % args))


try:
    httpd = http.server.ThreadingHTTPServer((HOST, PORT), NoCacheHandler)
except OSError as e:
    print(f"Cannot start: {HOST}:{PORT} may already be in use.")
    print("Close the previous server window and try again.")
    print(e)
    input("Press Enter to exit...")
    sys.exit(1)

context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
context.load_cert_chain(certfile=str(CERT_FILE), keyfile=str(KEY_FILE))
httpd.socket = context.wrap_socket(httpd.socket, server_side=True)

print("Local HTTPS server is running.")
print("Single-page URL:")
print(f"https://{HOST}:{PORT}/index.html")
print()
print("Merged collections use URLs under:")
print(f"https://{HOST}:{PORT}/pages/")
print()
print("Keep this window open while presenting. Press Ctrl+C to stop.")
try:
    httpd.serve_forever()
except KeyboardInterrupt:
    print("\nStopped.")
