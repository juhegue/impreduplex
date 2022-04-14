@echo off
set /p respuesta=Tipo de archivo Docx o PDF?(d/p):
IF %respuesta%==d goto docx
IF %respuesta%==p goto pdf
goto fin
:docx
@echo on
dist\impreduplex.exe imagenes\Factura.pdf imagenes\D*.* c:\tmp\doc.docx 2 4 ver
goto fin
:pdf
@echo on
dist\impreduplex.exe imagenes\Factura.pdf imagenes\D*.* c:\tmp\doc.pdf 2 4 ver
:fin
