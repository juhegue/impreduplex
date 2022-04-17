# -*- coding: utf-8 -*-

from io import BytesIO
import os
import sys
import glob
import platform
import subprocess
import tempfile
from PIL import Image
from pdf2image import convert_from_path
from fpdf import FPDF

if platform.system() == 'Windows':
    import win32print
    import win32event
    import win32process
    import win32con
    from win32com.shell.shell import ShellExecuteEx
    from win32com.shell import shellcon

__version__ = '0.0.1'
appName = 'impreduplex'


def win_duplex(nom_imp, duplex=3):
    print(f'Duplex({duplex}):{nom_imp}')
    printdefaults = {'DesiredAccess': win32print.PRINTER_ACCESS_USE}
    handle = win32print.OpenPrinter(nom_imp, printdefaults)
    level = 2
    attributes = win32print.GetPrinter(handle, level)
    antes = attributes['pDevMode'].Duplex
    attributes['pDevMode'].Duplex = duplex  # 1=no flip, 2=flip up, 3=flip over
    try:
        win32print.SetPrinter(handle, level, attributes, 0)
    except:
        pass

    return antes


def win_imprime(docu, impresora, duplex):
    duplex_ant = None
    if impresora == 'ver':
        procInfo = ShellExecuteEx(nShow=win32con.SW_HIDE,
                                  fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                  lpVerb='open',
                                  lpFile=docu,
                                  )
    elif impresora == 'defecto':
        if duplex:
            impresora = win32print.GetDefaultPrinter()
            duplex_ant = win_duplex(impresora)

        procInfo = ShellExecuteEx(nShow=win32con.SW_HIDE,
                                  fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                  lpVerb='printto',
                                  lpFile=docu,
                                  )
    else:
        if duplex:
            duplex_ant = win_duplex(impresora)

        procInfo = ShellExecuteEx(nShow=win32con.SW_HIDE,
                                  fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                  lpVerb='printto',
                                  lpFile=docu,
                                  lpParameters='"%s"' % impresora
                                  )
    procHandle = procInfo['hProcess']
    obj = win32event.WaitForSingleObject(procHandle, win32event.INFINITE)
    rc = win32process.GetExitCodeProcess(procHandle)
    if duplex_ant:
        win_duplex(impresora, duplex_ant)


def escala_imagen(image, width, height):
    w, h = image.size
    w = height * w / h
    if w > width:
        height = height * width / w
    else:
        width = w
    return int(width), int(height)


class MiFPDF(FPDF):

    def image(self, image, x, y, w, h):
        tmp_file = os.path.join(tempfile.gettempdir(), next(tempfile._get_candidate_names())) + '.png'
        image.save(tmp_file)
        super().image(tmp_file, x=x, y=y, w=w, h=h, type='png', link='')
        os.remove(tmp_file)

    def footer(self):
        self.set_y(-6)
        self.set_font('Arial', 'I', 6)
        self.cell(0, 10, f'Página {self.page_no()}' + '/{nb}', 0, 0, 'C')


def paste_imagen(pdf, imagen, posx, posy, width, height):
    image = Image.open(imagen)
    w, h = escala_imagen(image, width, height)
    inc_y = (height - h) / 2
    inc_x = (width - w) / 2
    image = image.convert('RGB')
    pdf.image(image, posx + inc_x, posy + inc_y, w, h)


def albaranes(pdf, albaranes_img, albaran, img_pag_ancho, img_pag_alto):
    w, h = (pdf.fh, pdf.fw) if pdf.def_orientation == 'L' else (pdf.fw, pdf.fh)

    pdf.add_page()
    posy = 0
    for alto in range(0, img_pag_alto):
        posx = 0
        for ancho in range(0, img_pag_ancho):
            if albaran < len(albaranes_img):
                img = albaranes_img[albaran]
                paste_imagen(pdf, img, posx, posy, w / img_pag_ancho, h / img_pag_alto)
                posx += w / img_pag_ancho
                albaran += 1
        posy += h / img_pag_alto
    return albaran


def crea_pdf(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto):
    w, h = Image.open(facturas_img[0]).size
    orientation = 'L' if w > h else 'P'
    pdf = MiFPDF(orientation=orientation)
    pdf.set_margins(0, 0, 0)
    pdf.alias_nb_pages()

    albaran = 0
    for factura in facturas_img:
        pdf.add_page()
        w, h = (pdf.fh, pdf.fw) if pdf.def_orientation == 'L' else (pdf.fw, pdf.fh)
        paste_imagen(pdf, factura, 0, 0, w, h)
        if albaran < len(albaranes_img):
            albaran = albaranes(pdf, albaranes_img, albaran, img_pag_ancho, img_pag_alto)

    while albaran < len(albaranes_img):
        albaran = albaranes(pdf, albaranes_img, albaran, img_pag_ancho, img_pag_alto)

    # siempre pares
    if pdf.page_no() % 2 != 0:
        pdf.add_page()

    pdf.output(doc_destino)


def main():
    args = sys.argv[1:]
    if len(args) < 5:
        print('Error.')
        sys.exit(-1)

    file_factu = args[0]
    path_alba = args[1]
    doc_destino = args[2]
    img_pag_ancho = int(args[3])
    img_pag_alto = int(args[4])
    # ver=Visualiza y para windows: defecto=Impresora defecto o nombre impresora
    impresora = args[5] if len(args) > 5 else None
    duplex = True if len(args) > 6 else False

    facturas_img = list()
    imagenes = convert_from_path(file_factu)
    for imagen in imagenes:
        imagefile = BytesIO()
        imagen.save(imagefile, format='PNG')
        facturas_img.append(imagefile)

    albaranes_img = list()
    for albaran in sorted(glob.glob(path_alba)):
        with open(albaran, 'rb') as f:
            data = f.read()

        if data[1:4] == b'PDF':
            imagenes = convert_from_path(albaran)
            for imagen in imagenes:
                imagefile = BytesIO()
                imagen.save(imagefile, format='PNG')
                albaranes_img.append(imagefile)
        else:
            albaranes_img.append(BytesIO(data))

    crea_pdf(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto)

    if impresora:
        if platform.system() == 'Darwin':       # mac
            subprocess.call(('open', doc_destino))
        elif platform.system() == 'Windows':
            win_imprime(doc_destino, impresora, duplex)
        else:                                   # linux
            subprocess.call(('xdg-open', doc_destino))


if __name__ == '__main__':
    main()

