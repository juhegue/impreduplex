@echo off
set /p respuesta=Tipo de archivo Docx/PDF/Ambos?(d/p/a):
IF %respuesta%==d goto docx
IF %respuesta%==p goto pdf
IF %respuesta%==a goto ambos
goto fin
:docx
@echo on
dist\impreduplex.exe imagenes\Factura.pdf imagenes\D*.* c:\tmp\doc.docx 2 4 ver
goto fin
:pdf
@echo on
dist\impreduplex.exe imagenes\Factura.pdf imagenes\D*.* c:\tmp\doc.pdf 2 4 ver
:ambos
@echo on
dist\impreduplex.exe imagenes\Factura.pdf imagenes\D*.* c:\tmp\doc.docx 2 4 c:\tmp\doc.pdf
:fin
