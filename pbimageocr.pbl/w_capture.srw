forward
global type w_capture from window
end type
end forward

global type w_capture from window
integer width = 1243
integer height = 440
windowtype windowtype = response!
long backcolor = 65535
string icon = "AppIcon!"
boolean toolbarvisible = false
boolean center = true
integer transparency = 75
end type
global w_capture w_capture

type prototypes
FUNCTION	uLong SetWindowLong ( long hWindow, integer nIndex, long dwNewLong ) LIBRARY "USER32.dll" Alias for  "SetWindowLongW"
FUNCTION	uLong  GetWindowLong ( uLong hWindow, integer nIndex )   Library "USER32.dll" Alias for  "GetWindowLongW"
//Function Boolean SetWindowPos(long hwnd, long hmode, integer ix, integer iy, integer cx, integer cy, ulong flags) library "user32.dll"
FUNCTION	ulong	 SetWindowPos ( ulong hwnd, ulong hwndAfter, ulong xPos, ulong yPos, ulong cX, ulong cY, long wFlage ) 	library "user32.dll" alias for "SetWindowPos"
FUNCTION	boolean	GetCursorPos (ref sr_point POINT)  LIBRARY "User32.dll"
end prototypes

type variables
n_bitmap in_bitmap
n_resizable_response in_rsz
String is_ClipBoard = "N"

// Para establecer el orden Z de la ventana.
CONSTANT ulong HWND_TOP = 0
CONSTANT ulong HWND_BOTTOM = 1
CONSTANT ulong HWND_TOPMOST = -1
CONSTANT ulong HWND_NOTOPMOST = -2

// Parámetros del falg de SetWindowPos
CONSTANT ulong SWP_NOSIZE = 1
CONSTANT ulong SWP_NOMOVE = 2
CONSTANT ulong SWP_NOZORDER = 4
CONSTANT ulong SWP_NOREDRAW = 8
CONSTANT ulong SWP_NOACTIVATE = 16
CONSTANT ulong SWP_FRAMECHANGED = 32
CONSTANT ulong SWP_SHOWWINDOW = 64
CONSTANT ulong SWP_HIDEWINDOW = 128
CONSTANT ulong SWP_NOCOPYBITS = 256
CONSTANT ulong SWP_NOOWNERZORDER = 512
CONSTANT ulong SWP_NOSENDCHANGING = 1024
CONSTANT ulong SWP_DRAWFRAME = SWP_FRAMECHANGED
CONSTANT ulong SWP_NOREPOSITION = SWP_NOOWNERZORDER
CONSTANT ulong SWP_DEFERERASE = 8192
CONSTANT ulong SWP_ASYNCWINDOWPOS = 16384
end variables

on w_capture.create
end on

on w_capture.destroy
end on

event open;long    ll_hWnd 


is_Clipboard = Message.StringParm

// make resizable
in_rsz.of_SetResizable(this)

//https://conpb.blogspot.com/
// Recogemos el handle de la ventana y la ponemos en primer plano.
ll_hWnd  = Handle(This) 
SetWindowPos(ll_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE + SWP_SHOWWINDOW)

end event

event mousemove;sr_point lsr_point

if flags = 1 then
	GetCursorPos(REF lsr_point)
	this.x = PixelsToUnits(lsr_point.il_x, XPixelsToUnits!) - (this.Width / 2)
	this.y = PixelsToUnits(lsr_point.il_y, YPixelsToUnits!) - (this.Height / 2)
end if
end event

event doubleclicked;String ls_fname
Blob lblb_bitmap

SetPointer(HourGlass!)

this.transparency = 100

IF is_ClipBoard = "N" THEN
	lblb_bitmap = in_bitmap.of_WindowCapture(this, False)
	ls_fname =  gs_dir + "Response.bmp"
	in_bitmap.of_WriteBlob(ls_fname, lblb_bitmap)
	CloseWithReturn(this, ls_fname)
ELSE
	lblb_bitmap = in_bitmap.of_WindowCapture(this, True)
	Close(this)
END IF	

end event

