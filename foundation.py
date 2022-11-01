import logging
import os
import shutil
import sys
from datetime import datetime
from io import BytesIO

import requests
import win32com.client
from PIL import Image


class Logger:
    def __init__(self):
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.DEBUG)

        handler = logging.StreamHandler(sys.stdout)
        handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter("%(asctime)s : %(levelname)s : %(message)s")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    def debug(self, *args):
        self.logger.debug(*args)

    def info(self, *args):
        self.logger.info(*args)

    def warning(self, *args):
        self.logger.warning(*args)

    def error(self, *args):
        self.logger.error(*args)

    def critical(self, *args):
        self.logger.critical(*args)


logger = Logger()


class Printer:
    WIA_COM = "WIA.CommonDialog"
    WIA_COMMAND_TAKE_PICTURE = "{AF933CAC-ACAD-11D2-A093-00C04F72DC3C}"
    WIA_IMG_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"

    def acquire_image_wia(self):
        """
        What I do?\n
            1. Connect to printer
            2. Scan the object
            3. Save the scanned image as png

        Ref\n
            https://stackoverflow.com/questions/15631046/i-want-to-connect-my-program-to-image-scanner"""

        if not os.path.exists("static/tmp"):
            os.mkdir("static/tmp")
        wia = win32com.client.Dispatch(self.WIA_COM)  # wia is a CommonDialog object
        dev = wia.ShowSelectDevice()
        try:
            for command in dev.Commands:
                if command.CommandID == self.WIA_COMMAND_TAKE_PICTURE:
                    dev.ExecuteCommand(self.WIA_COMMAND_TAKE_PICTURE)
        except:
            msg = "Device Error: Switch on the printer and connect to the computer"
            logger.error(msg)

        logger.info("Scanning...")
        for i, item in enumerate(dev.Items, start=1):
            if i == dev.Items.Count:
                image = item.Transfer(self.WIA_IMG_FORMAT_PNG)
                break
        logger.info("Scan Completed")

        try:
            timestamp = str(datetime.now()).replace(":", "-")[:19].replace(" ", "_")
            os.mkdir(f"static/tmp/{timestamp}")
            fname = f"static/tmp/{timestamp}/receipt.png"
            image.SaveFile(fname)
            logger.info(f"Scanned image got saved as {fname}")
            return fname
        except Exception as e:
            logging.error("Could not save the scanned image as PNG", exc_info=True)


class ScannedImage:
    def __init__(self, scanned_image_path):
        self.scanned_image_path = scanned_image_path
        self.cropped_image = None

    def crop(self, dimension):
        response = requests.get(self.scanned_image_path)
        img = Image.open(BytesIO(response.content))
        left = dimension["x"]
        top = dimension["y"]
        right = left + dimension["width"]
        bottom = top + dimension["height"]
        self.cropped_image = img.crop((left, top, right, bottom))

        op_path = self.scanned_image_path.split("/")[-2]
        op_path = f"static/tmp/{op_path}/cropped.png"
        self.cropped_image.save(op_path)
        return op_path

    def saveAsPDF(self):
        addr = self.cropped_image
        if addr.width > 540:
            ratio = 540 / addr.width
            addr = addr.resize((int(addr.width * ratio), int(addr.height * ratio)))
        if addr.height > 256:
            ratio = 256 / addr.height
            addr = addr.resize((int(addr.width * ratio), int(addr.height * ratio)))

        logger.info(f"{addr.width}, {addr.height}")

        post = Image.open("./static/base.jpg")
        # post.crop((699,285,1259,561)).show()
        post.paste(addr, (699 + 10, 285 + 10))
        post = post.transpose(Image.ROTATE_90)
        a4 = Image.new("RGB", (1240, 1754), (255, 255, 255))
        a4.paste(post, (310, 425))
        a4 = a4.rotate(180)

        op_path = self.scanned_image_path.split("/")[-2]
        op_path = f"static/tmp/{op_path}/a4.pdf"
        a4.save(op_path)
        return op_path


class InitialWipeOut:
    def __init__(self) -> None:
        path = "static/tmp"
        logger.info(f"Removing the directory {path}")
        if os.path.exists(path):
            shutil.rmtree(path)
