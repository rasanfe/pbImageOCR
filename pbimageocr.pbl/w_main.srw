forward
global type w_main from window
end type
type cbx_clipboard from checkbox within w_main
end type
type cb_ocr from commandbutton within w_main
end type
type cb_pdf_to_png from commandbutton within w_main
end type
type cb_pdf_to_text from commandbutton within w_main
end type
type wb_1 from webbrowser within w_main
end type
type p_1 from picture within w_main
end type
type st_infocopyright from statictext within w_main
end type
type st_myversion from statictext within w_main
end type
type st_platform from statictext within w_main
end type
type sle_file from singlelineedit within w_main
end type
type st_file from statictext within w_main
end type
type cb_openfile from commandbutton within w_main
end type
type r_2 from rectangle within w_main
end type
end forward

global type w_main from window
integer width = 4631
integer height = 3368
boolean titlebar = true
string title = "PowerBuilder ImageOCR"
boolean controlmenu = true
boolean minbox = true
string icon = "AppIcon!"
boolean center = true
cbx_clipboard cbx_clipboard
cb_ocr cb_ocr
cb_pdf_to_png cb_pdf_to_png
cb_pdf_to_text cb_pdf_to_text
wb_1 wb_1
p_1 p_1
st_infocopyright st_infocopyright
st_myversion st_myversion
st_platform st_platform
sle_file sle_file
st_file st_file
cb_openfile cb_openfile
r_2 r_2
end type
global w_main w_main

type prototypes
Function boolean QueryPerformanceFrequency ( &
	Ref Double lpFrequency &
	) Library "kernel32.dll"

Function boolean QueryPerformanceCounter ( &
	Ref Double lpPerformanceCount &
	) Library "kernel32.dll"

end prototypes

type variables
Double idbl_frequency = 0



end variables

forward prototypes
public subroutine wf_version (statictext ast_version, statictext ast_patform)
end prototypes

public subroutine wf_version (statictext ast_version, statictext ast_patform);String ls_version, ls_platform
environment env
integer rtn

rtn = GetEnvironment(env)

IF rtn <> 1 THEN 
	ls_version = string(year(today()))
	ls_platform="32"
ELSE
	ls_version = "20"+ string(env.pbmajorrevision)+ "." + string(env.pbbuildnumber)
	ls_platform=string(env.ProcessBitness)
END IF

ls_platform += " Bits"

ast_version.text=ls_version
ast_patform.text=ls_platform

end subroutine

on w_main.create
this.cbx_clipboard=create cbx_clipboard
this.cb_ocr=create cb_ocr
this.cb_pdf_to_png=create cb_pdf_to_png
this.cb_pdf_to_text=create cb_pdf_to_text
this.wb_1=create wb_1
this.p_1=create p_1
this.st_infocopyright=create st_infocopyright
this.st_myversion=create st_myversion
this.st_platform=create st_platform
this.sle_file=create sle_file
this.st_file=create st_file
this.cb_openfile=create cb_openfile
this.r_2=create r_2
this.Control[]={this.cbx_clipboard,&
this.cb_ocr,&
this.cb_pdf_to_png,&
this.cb_pdf_to_text,&
this.wb_1,&
this.p_1,&
this.st_infocopyright,&
this.st_myversion,&
this.st_platform,&
this.sle_file,&
this.st_file,&
this.cb_openfile,&
this.r_2}
end on

on w_main.destroy
destroy(this.cbx_clipboard)
destroy(this.cb_ocr)
destroy(this.cb_pdf_to_png)
destroy(this.cb_pdf_to_text)
destroy(this.wb_1)
destroy(this.p_1)
destroy(this.st_infocopyright)
destroy(this.st_myversion)
destroy(this.st_platform)
destroy(this.sle_file)
destroy(this.st_file)
destroy(this.cb_openfile)
destroy(this.r_2)
end on

event open;wf_version(st_myversion, st_platform)

end event

type cbx_clipboard from checkbox within w_main
integer x = 2587
integer y = 344
integer width = 343
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 553648127
string text = "Clipboard"
end type

type cb_ocr from commandbutton within w_main
integer x = 2167
integer y = 332
integer width = 411
integer height = 92
integer taborder = 110
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "OCR"
end type

event clicked;nvo_imageocr ln_ocr
nvo_imagefromclipboard ln_clipboard
String ls_image, ls_error, ls_text
String ls_clipboard="N"

if cbx_clipboard.Checked then ls_clipboard = "S"

OpenWithParm(w_Capture, ls_clipboard)

ln_ocr = CREATE nvo_imageocr

IF ls_clipboard = "N" THEN
	ls_image = Message.StringParm
ELSE
	ln_clipboard =  CREATE nvo_imagefromclipboard
	
	ls_image = ln_clipboard.of_GetClipboardImage()
		
	//Checks the result
	If ln_clipboard.il_ErrorType < 0 Then
		ls_error = ln_clipboard.is_ErrorText + " Exception: "+ln_clipboard.of_GetLastError()
		MessageBox("Failed", ls_error )
		Return
	End If
END IF	

ls_text = ln_ocr.of_ConvertImageToString(ls_image)

//Checks the result
If ln_ocr.il_ErrorType < 0 Then
	ls_error = ln_ocr.is_ErrorText + " Exception: "+ln_ocr.of_GetLastError()
	MessageBox("Failed", ls_error )
	Return
End If

MessageBox("Result", ls_text)
Clipboard(ls_text)


end event

type cb_pdf_to_png from commandbutton within w_main
integer x = 3415
integer y = 332
integer width = 411
integer height = 92
integer taborder = 100
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Pdf to Png"
end type

event clicked;String ls_pdf, ls_error, ls_image, ls_format
Integer li_FormatLen
nvo_imagefrompdf nvo_ifpdf

ls_pdf = sle_file.text


nvo_ifpdf = CREATE nvo_imagefrompdf

ls_image = nvo_ifpdf.of_PdfToPng(ls_pdf)

//Checks the result
If nvo_ifpdf.il_ErrorType < 0 Then
	ls_error = nvo_ifpdf.is_ErrorText + " Exception: "+nvo_ifpdf.of_GetLastError()
	MessageBox("Failed", ls_error )
	Return
End If

sle_file.text = ls_image 
wb_1.Navigate(sle_file.text)
end event

type cb_pdf_to_text from commandbutton within w_main
integer x = 2976
integer y = 332
integer width = 411
integer height = 92
integer taborder = 90
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Pdf to Txt"
end type

event clicked;String ls_image, ls_txt, ls_error, ls_format, ls_pdf
Integer li_FormatLen
nvo_imageocr ln_ocr
nvo_imagefrompdf nvo_ifpdf
ln_ocr = CREATE nvo_imageocr

ls_image = sle_file.text

li_FormatLen = Len(ls_image) - LastPos(ls_image, ".") + 1

ls_format = lower(mid(ls_image, LastPos(ls_image, "."),  li_FormatLen))

ls_txt = replace(ls_image, pos(ls_image, ls_format), li_FormatLen, ".txt")
	
IF FileExists(ls_txt) THEN FileDelete(ls_txt)

IF lower(ls_format)=".pdf" THEN 
	nvo_ifpdf = CREATE nvo_imagefrompdf
	ls_pdf = ls_image
	ls_image = nvo_ifpdf.of_PdfToPng(ls_pdf)
END IF	

ln_ocr.of_ConvertImageToTXT(ls_image, ls_txt)

//Checks the result
If ln_ocr.il_ErrorType < 0 Then
	ls_error = ln_ocr.is_ErrorText + " Exception: "+ln_ocr.of_GetLastError()
	MessageBox("Failed", ls_error )
	Return
End If

sle_file.text = ls_txt 
wb_1.Navigate(sle_file.text)
end event

type wb_1 from webbrowser within w_main
integer x = 64
integer y = 472
integer width = 4489
integer height = 2712
end type

type p_1 from picture within w_main
integer x = 5
integer y = 4
integer width = 1253
integer height = 248
boolean originalsize = true
string picturename = "logo.jpg"
boolean focusrectangle = false
end type

type st_infocopyright from statictext within w_main
integer x = 3072
integer y = 3216
integer width = 1289
integer height = 56
integer textsize = -7
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 8421504
long backcolor = 553648127
string text = "Copyright © Ramón San Félix Ramón  rsrsystem.soft@gmail.com"
boolean focusrectangle = false
end type

type st_myversion from statictext within w_main
integer x = 4059
integer y = 60
integer width = 489
integer height = 84
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Versión"
alignment alignment = right!
boolean focusrectangle = false
end type

type st_platform from statictext within w_main
integer x = 4059
integer y = 148
integer width = 489
integer height = 84
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Bits"
alignment alignment = right!
boolean focusrectangle = false
end type

type sle_file from singlelineedit within w_main
integer x = 251
integer y = 332
integer width = 1673
integer height = 92
integer taborder = 60
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
boolean displayonly = true
borderstyle borderstyle = stylelowered!
end type

type st_file from statictext within w_main
integer x = 73
integer y = 344
integer width = 169
integer height = 76
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
string text = "File:"
boolean focusrectangle = false
end type

type cb_openfile from commandbutton within w_main
integer x = 1934
integer y = 332
integer width = 174
integer height = 92
integer taborder = 70
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "..."
end type

event clicked;integer li_rtn
string ls_path, ls_ruta
string  ls_current

ls_ruta= gs_dir
ls_current=GetCurrentDirectory ( )
li_rtn = GetFileOpenName ("Abrir",  sle_file.text, ls_path, "", "All Files (*.*),*.*") 
ChangeDirectory ( ls_current )

if li_rtn < 1 then 	sle_file.text = ""

wb_1.Navigate(sle_file.text)

end event

type r_2 from rectangle within w_main
long linecolor = 33554432
linestyle linestyle = transparent!
integer linethickness = 4
long fillcolor = 33521664
integer width = 4599
integer height = 260
end type

