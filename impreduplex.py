# -*- coding: utf-8 -*-

from io import BytesIO
import sys
import os
import glob
import platform
import subprocess
import base64
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Cm
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
if platform.system() == 'Windows':
    import win32api
    import win32com.client
    # import win32event
    # import win32process
    # import win32con
    # from win32com.shell.shell import ShellExecuteEx
    # from win32com.shell import shellcon

__version__ = '0.0.1'
appName = 'impreduplex'


img_relleno = """
iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAACXBI
WXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH3wYHFSo7+aaN1AAAADNJREFUOMtj/P//PwMlgImBQkCx
ASzInJ07dxKlyd3dfRB5YdSAYWEA44Blpt+/f1PHBQDWcw8IAHFJDwAAAABJRU5ErkJggg==
"""
img_relleno = base64.b64decode(img_relleno.replace('\n', ''))


# def win_imprime_espera(docu, impresora):
#     procInfo = ShellExecuteEx(nShow=win32con.SW_HIDE,
#                               fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
#                               lpVerb='printto',
#                               lpFile=docu,
#                               lpParameters='"%s"' % impresora
#                               )
#     procHandle = procInfo['hProcess']
#     obj = win32event.WaitForSingleObject(procHandle, win32event.INFINITE)
#     rc = win32process.GetExitCodeProcess(procHandle)


def win_imprime(docu, impresora):
    win32api.ShellExecute(
        0,
        'printto',
        docu,
        '"%s"' % impresora,
        '.',
        0
    )


def get_paginas(documento):
    # TODO:: Para corregir el error
    # win32com\client\CLSIDToClass.pyc: KeyError: '{00020970-0000-0000-C000-000000000046}'
    if win32com.client.gencache.is_readonly:
        win32com.client.gencache.is_readonly = False
        win32com.client.gencache.Rebuild()  # create gen_py folder if needed

    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    w = word.Documents.Open(documento)
    w.Repaginate()
    paginas = w.ComputeStatistics(2)
    w.Close()
    word.Quit()
    return paginas


def resize_imagen(imagen, width, height):
    try:
        im = Image.open(imagen)
        w, h = im.size
        w = height * w / h
        if w > width:
            height = height * width / w
        else:
            width = w
    except:
        pass

    return width, height


def albaranes(document, albaranes_img, albaran, img_pag_ancho, img_pag_alto, twidth, theight):
    twidth -= .5
    theight -= img_pag_alto     # para que no haga saltos de página
    table = document.add_table(rows=img_pag_alto, cols=img_pag_ancho)
    table.allow_autofit = False
    table.style = None
    for alto in range(0, img_pag_alto):
        for ancho in range(0, img_pag_ancho):
            cell = table.rows[alto].cells[ancho]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.style = None
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = paragraph.add_run()
            if albaran < len(albaranes_img):
                width = twidth / img_pag_ancho
                height = theight / img_pag_alto
                imagen = albaranes_img[albaran]
                w, h = resize_imagen(imagen, width, height)
                run.add_picture(BytesIO(img_relleno), width=Cm(width), height=Cm((height-h)/2))
                run.add_picture(imagen, width=Cm(w), height=Cm(h))
                run.add_picture(BytesIO(img_relleno), width=Cm(width), height=Cm((height-h))/2)
                albaran = albaran + 1
    return albaran


def crea_docx(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto):
    document = Document('impreduplex.docx')
    width = document.sections[0].page_width.cm
    height = document.sections[0].page_height.cm
    top = document.sections[0].top_margin.cm
    bottom = document.sections[0].bottom_margin.cm
    left = document.sections[0].left_margin.cm
    right = document.sections[0].right_margin.cm
    twidth = width - left - right
    theight = height - top - bottom

    for paragraph in document.paragraphs:
        p = paragraph._element
        p.getparent().remove(p)
        p._p = p._element = None

    paginas = albaran = 0
    for factura in facturas_img:
        paginas += 1
        document.add_picture(factura, width=Cm(twidth), height=Cm(theight))
        if albaran < len(albaranes_img):
            albaran = albaranes(document, albaranes_img, albaran, img_pag_ancho, img_pag_alto, twidth, theight)
            paginas += 1

    while albaran < len(albaranes_img):
        albaran = albaranes(document, albaranes_img, albaran, img_pag_ancho, img_pag_alto, twidth, theight)
        paginas += 1

    # Las páginas deben ser siempre pares
    if platform.system() == 'Windows':
        # en windows se cuentan las páginas con el Word
        document.save(doc_destino)
        paginas = get_paginas(doc_destino)
        print(f'Páginas creadas: {paginas}')
        if paginas % 2 != 0:
            document = Document(doc_destino)
            document.add_page_break()
            document.save(doc_destino)
            paginas = get_paginas(doc_destino)
            print(f'Página añadida, total: {paginas}')
    elif paginas % 2 != 0:
        document.add_page_break()
        # document.add_paragraph('.')
        document.save(doc_destino)


def main():
    args = sys.argv[1:]
    if len(args) < 5:
        print('Error.')
        sys.exit(-1)

    file_factu = args[0]
    path_alba = args[1]
    # TODO:: para windows path+fichero ya que se calculan las páginas
    doc_destino = args[2]
    img_pag_ancho = int(args[3])
    img_pag_alto = int(args[4])
    # ver=Visualiza y para windows: defecto=Impresora defecto o el nombre impresora
    impresora = args[5] if len(args) > 5 else None

    facturas_img = list()
    imagenes = convert_from_path(file_factu)
    for imagen in imagenes:
        imagefile = BytesIO()
        imagen.save(imagefile, format='PNG')
        facturas_img.append(imagefile)

    albaranes_img = list()
    for albaran in glob.glob(path_alba):
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

    crea_docx(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto)

    if impresora:
        if platform.system() == 'Darwin':       # mac
            subprocess.call(('open', doc_destino))
        elif platform.system() == 'Windows':
            if impresora == 'ver':
                os.startfile(doc_destino)
            elif impresora == 'defecto':
                os.startfile(doc_destino, 'print')
            elif impresora:
                # win_imprime_ex(doc_destino, impresora)
                win_imprime(doc_destino, impresora)
        else:                                   # linux
            subprocess.call(('xdg-open', doc_destino))


if __name__ == '__main__':
    main()

