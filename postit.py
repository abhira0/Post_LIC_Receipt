# This is to scan using cannon printer automatically and save it as image file
import win32com.client, time, os
import pygame, sys
from PIL import Image
import tkinter as tk
from tkinter.ttk import *
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import messagebox
import inspect
from reportlab.pdfgen.canvas import Canvasx
from reportlab.lib.units import inch, cm

WIA_COM = "WIA.CommonDialog"
WIA_DEVICE_UNSPECIFIED = 0
WIA_DEVICE_CAMERA = 2
WIA_INTENT_UNSPECIFIED = 0
WIA_BIAS_MIN_SIZE = 65536
WIA_BIAS_MAX_QUALITY = 65536
WIA_IMG_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
WIA_COMMAND_TAKE_PICTURE = "{AF933CAC-ACAD-11D2-A093-00C04F72DC3C}"
scanned = 0
cropped = 0
PATHS = []


def printAtConsole(header, info, solution=""):
    global console
    switch = {
        "fex": "FUNCTION EXIT",
        "FEX": "FUNCTION EXIT",
        "fen": "FUNCTION ENTRY",
        "FEN": "FUNCTION ENTRY",
        "tas": "TASK",
        "TAS": "TASK",
        "err": "ERROR",
        "ERR": "ERROR",
        "war": "WARNING",
        "WAR": "WARNING",
    }
    information = "${}$ {}".format(switch[header], info)
    if solution != "":
        information = "{}\n\t$TRY$ {}".format(information, solution)
    print(information)
    Label(console.scrollable_frame, text=information).grid(sticky="w")


def getImage(cropbtn):
    printAtConsole(
        "fen", "On: {} ;By: {}".format(inspect.stack()[0][3], inspect.stack()[1][3])
    )
    os.chdir(os.getcwd())
    acquire_image_wia()
    global scanned
    if scanned == 1:
        cropbtn["state"] = "normal"
    else:
        messagebox.showerror("Error", "Scan Again")
    printAtConsole(
        "fex", "On: {} ;By: {}".format(inspect.stack()[0][3], inspect.stack()[1][3])
    )


def acquire_image_wia():
    printAtConsole(
        "fen", "On: {} ;By: {}".format(inspect.stack()[0][3], inspect.stack()[1][3])
    )
    wia = win32com.client.Dispatch(WIA_COM)  # wia is a CommonDialog object
    dev = wia.ShowSelectDevice()
    try:
        for command in dev.Commands:
            if command.CommandID == WIA_COMMAND_TAKE_PICTURE:
                dev.ExecuteCommand(WIA_COMMAND_TAKE_PICTURE)
    except:
        printAtConsole(
            "err",
            "Device Error",
            solution="Switch on the printer and connect to the computer]",
        )
    i = 1
    printAtConsole("tas", "Scanning...")
    for item in dev.Items:
        if i == dev.Items.Count:
            image = item.Transfer(WIA_IMG_FORMAT_PNG)
            break
        i = i + 1
    printAtConsole("tas", "Scan Completed")
    try:
        fname = "receipt.png"
        if os.path.exists(fname):
            os.remove(fname)
            print("$TASK$ Older scanned image got removed")
        image.SaveFile(fname)
        printAtConsole("tas", "Scanned image got saved as {}".format(fname))
        global scanned
        scanned = 1
    except:
        scanned = 0
    printAtConsole(
        "fex", "On: {} ;By: {}".format(inspect.stack()[0][3], inspect.stack()[1][3])
    )


def displayImage(screen, px, topleft, prior):
    # ensure that the rect always has positive width, height
    x, y = topleft
    width = pygame.mouse.get_pos()[0] - topleft[0]
    height = pygame.mouse.get_pos()[1] - topleft[1]
    if width < 0:
        x += width
        width = abs(width)
    if height < 0:
        y += height
        height = abs(height)

    # eliminate redundant drawing cycles (when mouse isn't moving)
    current = x, y, width, height
    if not (width and height):
        return current
    if current == prior:
        return current

    # draw transparent box and blit it onto canvas
    screen.blit(px, px.get_rect())
    im = pygame.Surface((width, height))
    im.fill((128, 128, 128))
    pygame.draw.rect(im, (32, 32, 32), im.get_rect(), 1)
    im.set_alpha(128)
    screen.blit(im, (x, y))
    pygame.display.flip()

    # return current box extents
    return (x, y, width, height)


def setup(path):
    px = pygame.image.load(path)
    screen = pygame.display.set_mode(px.get_rect()[2:])
    screen.blit(px, px.get_rect())
    pygame.display.flip()
    return screen, px


def mainLoop(screen, px):
    topleft = bottomright = prior = None
    n = 0
    while n != 1:
        for event in pygame.event.get():
            if event.type == pygame.MOUSEBUTTONUP:
                if not topleft:
                    topleft = event.pos
                else:
                    bottomright = event.pos
                    n = 1
        if topleft:
            prior = displayImage(screen, px, topleft, prior)
    return topleft + bottomright


def doTheJob(do):
    printAtConsole(
        "fen", "On: {} ;By: {}".format(inspect.stack()[0][3], inspect.stack()[1][3])
    )
    printAtConsole("tas", "Initializing pygame window")
    pygame.init()
    input_loc = "receipt.png"
    img = Image.open(input_loc)
    input_loc2 = "image.png"
    img.resize((int(img.width * 0.20), int(img.height * 0.20))).save(input_loc2)
    printAtConsole("tas", "Saving after resizing the copy of the scanned image")
    output_loc = "out.png"
    screen, px = setup(input_loc2)
    left, upper, right, lower = mainLoop(screen, px)
    # ensure output rect always has positive width, height
    if right < left:
        left, right = right, left
    if lower < upper:
        lower, upper = upper, lower
    printAtConsole("tas", "Cropping the image with the given size")
    im = Image.open(input_loc)
    im = im.crop((left / 0.2, upper / 0.2, right / 0.2, lower / 0.2))
    pygame.display.quit()
    im.save(output_loc)
    global cropped
    cropped = 1
    do["state"] = "normal"
    im.show()
    printAtConsole(
        "fex", "On: {} ;By: {}".format(inspect.stack()[0][3], inspect.stack()[1][3])
    )


def getPdf():
    address = Image.open("out.png")


class C4U:
    # sizes (w*h) are in mm
    versionX = "1.0"
    paper_sizes = {"A4": (210, 297), "A5": (148, 210)}
    grid_offsets = {
        "xscreen": 455,
        "yscreen": 450,
        "x1": 20,
        "y1": 10,
        "x2": 2,
        "y2": 2,
    }
    photo_sizes = {
        "Indian Passport Size": (35, 45),
        "Indian Stamp Size": (20, 25),
        "Indian SSLC Size": (25, 25),
        "Indian Normal Stamp size": (25, 30),
        "Indian Pan Card Size": (25, 35),
        "Indian Passport Form Size": (35, 35),
    }
    photo_borders = {"passport": (12, 12)}


class ScrollableFrame(Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        pre = Frame(self)
        pre.pack(side="bottom")
        scrollbar = Scrollbar(self, orient="vertical", command=canvas.yview)
        xs = Scrollbar(pre, orient="horizontal", command=canvas.xview)
        self.scrollable_frame = Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=xs.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        xs.pack(fill="x")


def addAddress():
    printAtConsole("tas", "Removing unnecessary files")
    input_loc = "receipt.png"
    input_loc2 = "image.png"
    os.remove(input_loc2)
    os.remove(input_loc)
    post = Image.open("./img/post.jpg")
    # post.crop((699,285,1259,561)).show()
    addr = Image.open("out.png")
    if addr.width > 540:
        ratio = 540 / addr.width
        addr = addr.resize((int(addr.width * ratio), int(addr.height * ratio)))
    if addr.height > 256:
        ratio = 256 / addr.height
        addr = addr.resize((int(addr.width * ratio), int(addr.height * ratio)))

    print(addr.width, addr.height)
    newpost = post.copy()
    newpost.paste(addr, (699 + 10, 285 + 10))
    newpost = newpost.transpose(Image.ROTATE_90)
    a4 = Image.new("RGB", (1240, 1754), (255, 255, 255))
    a4.paste(newpost, (310, 425))
    a4 = a4.rotate(180)
    a4.save("post.pdf")
    os.remove("out.png")


window = tk.Tk()
window.title("All Size Photo Builder v{}".format(C4U.versionX))
window.geometry(
    "{}x{}".format(C4U.grid_offsets["xscreen"], C4U.grid_offsets["yscreen"])
)
window.resizable(0, 0)
# ------------------------------------------------------------------
# Main Frames
# ------------------------------------------------------------------
scan_crop = Frame(window)
scan_crop.grid(
    column=0,
    row=2,
    columnspan=2,
    pady=C4U.grid_offsets["y1"],
    padx=C4U.grid_offsets["x1"],
)
Separator(window).grid(column=0, row=3, sticky="ew")
console = ScrollableFrame(window)
console.grid(
    column=0,
    row=4,
    columnspan=2,
    pady=C4U.grid_offsets["y1"],
    padx=C4U.grid_offsets["x1"],
)
Separator(window).grid(column=0, row=5)
# ------------------------------------------------------------------
# Items
# ------------------------------------------------------------------
# ----------------
# Scan and crop and do
# ----------------
imgbtn = Button(scan_crop, text="Scan", command=lambda: getImage(cropbtn))
imgbtn.grid(
    ipady=5,
    ipadx=10,
    column=0,
    row=0,
    pady=C4U.grid_offsets["y2"],
    padx=C4U.grid_offsets["x2"],
)
cropbtn = Button(scan_crop, state="disabled", text="Crop", command=lambda: doTheJob(do))
cropbtn.grid(
    ipady=5,
    ipadx=10,
    column=0,
    row=1,
    pady=C4U.grid_offsets["y2"],
    padx=C4U.grid_offsets["x2"],
)
do = Button(scan_crop, state="disabled", text="Get Print", command=lambda: addAddress())
do.grid(
    ipady=5,
    ipadx=10,
    column=0,
    row=2,
    pady=C4U.grid_offsets["y2"],
    padx=C4U.grid_offsets["x2"],
)
# ----------------
# Console
# ----------------
# # Added at runtime
# ------------------------------------------------------------------
# Loop
# ------------------------------------------------------------------
window.mainloop()

# wdFormatPDF = 17
#
# in_file = os.getcwd()+"\\newpost.pdf"
# out_file = os.getcwd()+"\\address.docx"
# print(in_file)
# word = win32com.client.Dispatch('Word.Application')
# doc = word.Documents.Open(in_file)
# with doc.PageSetup:
#     PageHeight = InchesToPoints(8)
#     PageWidth = InchesToPoints(4)
# # doc.SaveAs(out_file, FileFormat=wdFormatPDF)
# doc.Close()
# word.Quit()
#
