crea_pdf(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto)


def crea_pdf(facturas_img, albaranes_img, doc_destino, img_pag_ancho, img_pag_alto):
    relleno = Image.open(BytesIO(img_relleno)).convert('RGB')
    lista = list()
    albaran = 0
    for n, factura in enumerate(facturas_img):
        img = Image.open(factura)
        twidth, theight = img.size
        lista.append(img.convert('RGB'))
        pagina = Image.new('RGB', (twidth, theight), 'white')
        hpos = 0
        for alto in range(0, img_pag_alto):
            wpos = 0
            for ancho in range(0, img_pag_ancho):
                img = Image.open(albaranes_img[albaran])
                width = int(twidth / img_pag_ancho)
                height = int(theight / img_pag_alto)

                w, h = resize_imagen(img, width, height)
                # img_size = relleno.resize((int(w), int((height - h) / 2)))
                # pagina.paste(img_size, (wpos, hpos))
                # hpos += int((height - h) / 2)


                img_size = img.resize((int(w), int(h)))
                pagina.paste(img_size, (wpos, hpos))
                hpos += h

                # img_size = relleno.resize((int(w), int((height - h) / 2)))
                # pagina.paste(img_size, (wpos, hpos))
                # hpos += int((height - h) / 2)


                pagina.paste(relleno, (wpos, hpos))
                wpos += int(twidth / img_pag_ancho)
                albaran += 1
            #hpos += theight / img_pag_alto
        lista.append(pagina)

    lista[0].save(doc_destino, save_all=True, append_images=lista[1:])



https://stackoverflow.com/questions/47435973/print-pdf-file-in-duplex-mode-via-python