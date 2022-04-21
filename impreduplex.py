# -*- coding: utf-8 -*-

import os
from io import BytesIO
import math
import sys
import glob
import platform
import subprocess
from PIL import Image
from pdf2image import convert_from_path
if platform.system() == 'Windows':
    import win32print
    import win32event
    import win32process
    import win32con
    from win32com.shell.shell import ShellExecuteEx
    from win32com.shell import shellcon

__version__ = '0.0.1'
appName = 'impreduplex'
author = 'juhegue'
date = 'jue 14 abr 2022'

path = sys.executable if hasattr(sys, 'frozen') else sys.argv[0]
path = os.path.split(path)[0]

FORMATO = ''
DPI = 200
GHOSTSCRIPT_PATH = os.path.join(path, 'gswin64c.exe')
GSPRINT_PATH = os.path.join(path, 'gsprint.exe')


def win_duplex(nom_imp, duplex):
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
    else:
        if impresora == 'defecto':
            impresora = win32print.GetDefaultPrinter()

        print(f'Impresora: {impresora}')

        if duplex:
            print('Asignado duplex: 3')
            duplex_ant = win_duplex(impresora, 3)

        param = f'-ghostscript "{GHOSTSCRIPT_PATH}" -dPDFFitPage -dFitPage -{FORMATO} -color -q -r{DPI} -printer "{impresora}" "{docu}"'

        print(f'{GSPRINT_PATH} {param}')

        procInfo = ShellExecuteEx(nShow=win32con.SW_HIDE,
                                  fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                  lpVerb='open',
                                  lpFile=GSPRINT_PATH,
                                  lpParameters=param
                                  )

    procHandle = procInfo['hProcess']
    obj = win32event.WaitForSingleObject(procHandle, win32event.INFINITE)
    rc = win32process.GetExitCodeProcess(procHandle)

    if duplex_ant:
        print(f'Restaurando duplex: {duplex_ant}')
        win_duplex(impresora, duplex_ant)


def entero(numero):
    return int(math.modf(numero)[1])


def escala_imagen(image, width, height):
    w, h = image.size
    w = height * w / h
    if w > width:
        height = height * width / w
    else:
        width = w
    return entero(width), entero(height)


def paste_imagen(pagina, imagen, posx, posy, width, height):
    image = Image.open(imagen)
    w, h = escala_imagen(image, width, height)
    inc_y = (height - h) / 2
    inc_x = (width - w) / 2
    image = image.convert('RGB').resize((w, h))
    pagina.paste(image, box=(entero(posx + inc_x), entero(posy + inc_y)))


def albaranes_pdf(pagina, width, height, albaranes_img, albaran, img_pag_ancho, img_pag_alto):
    posy = 0
    for alto in range(0, img_pag_alto):
        posx = 0
        for ancho in range(0, img_pag_ancho):
            if albaran < len(albaranes_img):
                img = albaranes_img[albaran]
                paste_imagen(pagina, img, posx, posy, width / img_pag_ancho, height / img_pag_alto)
                posx += width / img_pag_ancho
                albaran += 1
        posy += height / img_pag_alto
    return albaran


def crea_pdf(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto):
    """
    A4 es de 210 x 297 mm
    1 Pulgada es 25,4 mm
    A4 es 8,27 x 11,69 pulgadas.
    """
    global FORMATO

    w, h = Image.open(facturas_img[0]).size
    if w > h:
        height, width = entero(8.27 * DPI), entero(11.69 * DPI)  # A4 Landscape
        FORMATO = 'landscape'
    else:
        width, height = entero(8.27 * DPI), entero(11.69 * DPI)  # A4 Portrait
        FORMATO = 'portrait'

    paginas = list()
    albaran = 0
    for factura in facturas_img:
        pagina = Image.new('RGB', (width, height), 'white')
        paste_imagen(pagina, factura, 0, 0, width, height)
        paginas.append(pagina)
        if albaran < len(albaranes_img):
            pagina = Image.new('RGB', (width, height), 'white')
            albaran = albaranes_pdf(pagina, width, height, albaranes_img, albaran, img_pag_ancho, img_pag_alto)
            paginas.append(pagina)

    while albaran < len(albaranes_img):
        pagina = Image.new('RGB', (width, height), 'white')
        albaran = albaranes_pdf(pagina, width, height, albaranes_img, albaran, img_pag_ancho, img_pag_alto)
        paginas.append(pagina)

    # siempre pares
    if len(paginas) % 2 != 0:
        pagina = Image.new('RGB', (width, height), 'white')
        paginas.append(pagina)

    paginas[0].save(doc_destino, save_all=True, append_images=paginas[1:])


def main():
    args = sys.argv[1:]
    if len(args) < 5:
        print('Error parámetros.')
        sys.exit(-1)

    file_factu = args[0]
    path_alba = args[1]
    doc_destino = args[2]
    img_pag_ancho = int(args[3])
    img_pag_alto = int(args[4])
    # ver=Visualiza y para windows: defecto=Impresora defecto o nombre impresora
    impresora = args[5] if len(args) > 5 else None
    tiempo_acrobat = int(args[6]) if len(args) > 6 else 0   # NO SE USA
    duplex = True if len(args) > 7 else False

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
