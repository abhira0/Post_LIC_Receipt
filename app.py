from urllib.parse import quote_plus, unquote_plus

import pythoncom
from flask import (
    Flask,
    send_from_directory,
    redirect,
    render_template,
    request,
    url_for,
)

from foundation import Printer, ScannedImage

app = Flask(__name__)

SESSION_DATA = {}
SCAN_ERROR = ""


@app.route("/")
def home():
    return render_template("home.html", images=SESSION_DATA)


@app.route("/scan")
def scan():
    # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
    # To avoid errors of COM, use the below line
    pythoncom.CoInitialize()

    # Scan the object and return the absolute path of the scanned image
    scanned_image_path = Printer().acquire_image_wia()

    return redirect(f"/crop?scanned_image_path={quote_plus(scanned_image_path)}")
    global SCAN_ERROR
    try:
        # https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
        # To avoid errors of COM, use the below line
        pythoncom.CoInitialize()
        # Scan the object and return the absolute path of the scanned image
        scanned_image_path = Printer().acquire_image_wia()
        SCAN_ERROR = ""
    except Exception as e:
        SCAN_ERROR = (
            "1. Please turn on printer and connect it to your computer;\n"
            "2. Do not turn off computer in the middle of the scan"
        )


@app.route("/crop")
def crop():
    scanned_image_path = unquote_plus(request.args["scanned_image_path"])
    print(f"crop(): {scanned_image_path}")
    return render_template("crop.html", scanned_image_path=scanned_image_path)


@app.route("/postCropInfo", methods=["POST"])
def postCropInfo():
    resp = request.json
    print(resp)
    scanned_image_path = resp["scanned_image_path"]
    SESSION_DATA[scanned_image_path] = resp
    image = ScannedImage(scanned_image_path)
    cropped_image_path = image.crop(resp)
    SESSION_DATA[scanned_image_path]["cropped_image_path"] = cropped_image_path
    pdf_path = image.saveAsPDF()
    SESSION_DATA[scanned_image_path]["pdf_path"] = pdf_path
    return redirect("/")


@app.route("/download_pdf")
def tos():
    filepath = "/static/2022-10-25_16-59-59/a4.pdf"
    return f"""<embed src="{filepath}">"""


if __name__ == "__main__":
    app.run(host="0.0.0.0", port="1234")
