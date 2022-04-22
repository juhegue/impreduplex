# -*- coding: utf-8 -*-

import os
from io import BytesIO
import math
import sys
import glob
import platform
import subprocess
from PIL import Image, ImageWin
from pdf2image import convert_from_path
if platform.system() == 'Windows':
    import win32con
    import win32print
    import win32ui

__version__ = '0.0.1'
appName = 'impreduplex'
author = 'juhegue'
date = 'jue 14 abr 2022'

FORMATO = ''
DPI = 200


def win_set_atributo_impresora(nombre_impresora, atributo, valor):
    printdefaults = {'DesiredAccess': win32print.PRINTER_ACCESS_USE}
    handle = win32print.OpenPrinter(nombre_impresora, printdefaults)
    level = 2
    attributes = win32print.GetPrinter(handle, level)
    anterior = getattr(attributes['pDevMode'], atributo)
    setattr(attributes['pDevMode'], atributo, valor)
    try:
        win32print.SetPrinter(handle, level, attributes, 0)
    except:
        pass
    win32print.ClosePrinter(handle)
    return anterior


def win_imprime(paginas, impresora, duplex):
    if impresora == 'defecto':
        impresora = win32print.GetDefaultPrinter()

    if duplex:
        print(f'Asigna duplex: {duplex}')
        duplex_ant = win_set_atributo_impresora(impresora, 'Duplex', duplex)

    orientacion = 1 if FORMATO == 'portrait' else 2
    orientacion = win_set_atributo_impresora(impresora, 'Orientation', orientacion)

    print(f'Imprimir ({FORMATO}): {impresora}')
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(impresora)
    horzres = hdc.GetDeviceCaps(win32con.HORZRES)
    vertres = hdc.GetDeviceCaps(win32con.VERTRES)
    try:
        hdc.StartDoc('Factura')
        for n, img in enumerate(paginas):
            print(f'Página: {n + 1}')
            img_width, img_height = img.size

            ratio = horzres / vertres
            max_height = img_height
            max_width = (int)(max_height * ratio)

            # ajusta imagen al tamaño de la página
            hdc.SetMapMode(win32con.MM_ISOTROPIC)
            hdc.SetViewportExt((horzres, vertres))
            hdc.SetWindowExt((max_width, max_height))

            # desplazamiento para centrar
            offset_x = (int)((max_width - img_width) / 2)
            offset_y = (int)((max_height - img_height) / 2)
            hdc.SetWindowOrg((-offset_x, -offset_y))

            hdc.StartPage()
            dib = ImageWin.Dib(img)
            dib.draw(hdc.GetHandleOutput(), (0, 0, img_width, img_height))
            hdc.EndPage()

        hdc.EndDoc()
    except Exception as e:
        print(e)

    hdc.DeleteDC()

    win_set_atributo_impresora(impresora, 'Orientation', orientacion)

    if duplex:
        print(f'Restaura duplex: {duplex_ant}')
        win_set_atributo_impresora(impresora, 'Duplex', duplex_ant)


def win_ver(docu):
    arg = ['cmd', '/C', f'{docu}']
    info = subprocess.STARTUPINFO()
    info.dwFlags = subprocess.STARTF_USESHOWWINDOW
    info.wShowWindow = win32con.SW_HIDE
    subprocess.run(arg, startupinfo=info)


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
    return paginas


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
    duplex = int(args[6]) if len(args) > 6 else None

    print('Cargando: ', end='')
    print(f'{os.path.basename(file_factu)}', end='', flush=True)
    facturas_img = list()
    imagenes = convert_from_path(file_factu)
    for imagen in imagenes:
        imagefile = BytesIO()
        imagen.save(imagefile, format='PNG')
        facturas_img.append(imagefile)

    albaranes_img = list()
    for albaran in sorted(glob.glob(path_alba)):
        print(f', {os.path.basename(albaran)}', end='', flush=True)
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

    print('\nCreando pdf.')
    paginas = crea_pdf(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto)

    if impresora:
        if platform.system() == 'Darwin':       # mac
            subprocess.call(('open', doc_destino))
        elif platform.system() == 'Windows':    # windows
            win_ver(doc_destino) if impresora == 'ver' else win_imprime(paginas, impresora, duplex)
        else:                                   # linux
            subprocess.call(('xdg-open', doc_destino))


if __name__ == '__main__':
    main()
