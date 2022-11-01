from urllib.parse import quote_plus, unquote_plus

import pythoncom
from flask import Flask, redirect, render_template, request

from foundation import InitialWipeOut, Printer, ScannedImage

app = Flask(__name__)

FLASK_IP = "127.0.0.1"
FLASK_PORT = "1234"

SESSION_DATA = {}
SCAN_ERROR = ""


@app.route("/")
def home():
    return render_template("home.html", images=SESSION_DATA, SCAN_ERROR=SCAN_ERROR)


@app.route("/scan")
def scan():
    global SCAN_ERROR
    try:
        # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
        # To avoid errors of COM, use the below line
        pythoncom.CoInitialize()
        # Scan the object and return the absolute path of the scanned image
        scanned_image_path = Printer().acquire_image_wia()
        scanned_image_path = f"http://{FLASK_IP}:{FLASK_PORT}/{scanned_image_path}"
        SESSION_DATA[scanned_image_path] = {
            "crop_url": f"/crop?scanned_image_path={quote_plus(scanned_image_path)}"
        }
        defaultCrop(scanned_image_path)
        SCAN_ERROR = ""
    except Exception as e:
        SCAN_ERROR = (
            "1. Please turn on printer and connect it to your computer;\n"
            "2. Do not turn off computer in the middle of the scan"
        )
    return redirect("/")


@app.route("/crop")
def crop():
    scanned_image_path = unquote_plus(request.args["scanned_image_path"])
    print(f"crop(): {scanned_image_path}")
    return render_template("crop.html", scanned_image_path=scanned_image_path)


@app.route("/postCropInfo", methods=["POST"])
def postCropInfo():
    resp = request.json
    print(resp)
    cropAndSave(resp)
    return redirect("/")


def defaultCrop(scanned_image_path):
    resp = {
        "scanned_image_path": scanned_image_path,
        "x": 10,
        "y": 1855,
        "width": 685,
        "height": 295,
    }
    cropAndSave(resp)


def cropAndSave(resp):
    scanned_image_path = resp["scanned_image_path"]
    SESSION_DATA[scanned_image_path].update(resp)
    image = ScannedImage(scanned_image_path)
    cropped_image_path = image.crop(resp)
    SESSION_DATA[scanned_image_path]["cropped_image_path"] = cropped_image_path
    pdf_path = image.saveAsPDF()
    SESSION_DATA[scanned_image_path]["pdf_path"] = pdf_path


if __name__ == "__main__":
    InitialWipeOut()
    app.run(host=FLASK_IP, port=FLASK_PORT)
