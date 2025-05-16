forward
global type n_bitmap from nonvisualobject
end type
type bitmapfileheader from structure within n_bitmap
end type
type bitmapinfoheader from structure within n_bitmap
end type
type bitmapinfo from structure within n_bitmap
end type
end forward

type bitmapfileheader from structure
	integer		bftype
	long		bfsize
	integer		bfreserved1
	integer		bfreserved2
	long		bfoffbits
end type

type bitmapinfoheader from structure
	long		bisize
	long		biwidth
	long		biheight
	integer		biplanes
	integer		bibitcount
	long		bicompression
	long		bisizeimage
	long		bixpelspermeter
	long		biypelspermeter
	long		biclrused
	long		biclrimportant
end type

type bitmapinfo from structure
	bitmapinfoheader		bmiheader
	unsignedlong		bmicolors[]
end type

global type n_bitmap from nonvisualobject autoinstantiate
end type

type prototypes
Function ulong GetLastError( ) Library "kernel32.dll"

Function ulong FormatMessage( &
	ulong dwFlags, &
	ulong lpSource, &
	ulong dwMessageId, &
	ulong dwLanguageId, &
	Ref string lpBuffer, &
	ulong nSize, &
	ulong Arguments &
	) Library "kernel32.dll" Alias For "FormatMessageW"

Function long GetTempPath ( &
	long nBufferLength, &
	Ref string lpBuffer &
	) Library "kernel32.dll" Alias For "GetTempPathW"

Function longptr GetDesktopWindow ( &
	) Library "user32.dll"

Function longptr GetDC ( &
	longptr hWnd &
	) Library "user32.dll"

Function longptr CreateCompatibleDC ( &
	longptr hdc &
	) Library "gdi32.dll"

Function longptr CreateCompatibleBitmap ( &
	longptr hdc, &
	long nWidth, &
	long nHeight &
	) Library "gdi32.dll"

Function longptr SelectObject ( &
	longptr hdc, &
	longptr hgdiobj &
	) Library "gdi32.dll"

Function boolean BitBlt ( &
	longptr hdcDest, &
	long nXDest, &
	long nYDest, &
	long nWidth, &
	long nHeight, &
	longptr hdcSrc, &
	long nXSrc, &
	long nYSrc, &
	long dwRop &
	) Library "gdi32.dll"

Function long GetDIBits ( &
	longptr hdc, &
	longptr hbmp, &
	ulong uStartScan, &
	ulong cScanLines, &
	Ref blob lpvBits, &
	Ref bitmapinfo lpbi, &
	ulong uUsage &
	) Library "gdi32.dll"

Function long GetDIBits ( &
	longptr hdc, &
	longptr hbmp, &
	ulong uStartScan, &
	ulong cScanLines, &
	longptr lpvBits, &
	ref bitmapinfo lpbi, &
	ulong uUsage &
	) Library "gdi32.dll"

Subroutine CopyBitmapFileHeader ( &
	Ref blob Destination, &
	bitmapfileheader Source, &
	long Length &
	) Library "kernel32.dll" Alias For "RtlMoveMemory" progma_pack(1)

Subroutine CopyBitmapInfo ( &
	Ref blob Destination, &
	bitmapinfo Source, &
	long Length &
	) Library "kernel32.dll" Alias For "RtlMoveMemory" progma_pack(1)

Function boolean DeleteDC ( &
	longptr hdc &
	) Library "gdi32.dll"

Function long ReleaseDC ( &
	longptr hWnd, &
	longptr hDC &
	) Library "user32.dll"

Function boolean OpenClipboard ( &
	longptr hWndNewOwner &
	) Library "user32.dll"

Function boolean EmptyClipboard ( &
	) Library "user32.dll"

Function boolean CloseClipboard ( &
	) Library "user32.dll"

Function longptr SetClipboardData ( &
	ulong uFormat, &
	longptr hMem &
	) Library "user32.dll"

Function long GetSystemMetrics ( &
	long nIndex &
	) Library "user32.dll"

Function longptr CreateFile ( &
	string lpFileName, &
	ulong dwDesiredAccess, &
	ulong dwShareMode, &
	ulong lpSecurityAttributes, &
	ulong dwCreationDisposition, &
	ulong dwFlagsAndAttributes, &
	ulong hTemplateFile &
	) Library "kernel32.dll" Alias For "CreateFileW"

Function boolean WriteFile ( &
	longptr hFile, &
	blob lpBuffer, &
	ulong nNumberOfBytesToWrite, &
	Ref ulong lpNumberOfBytesWritten, &
	longptr lpOverlapped &
	) Library "kernel32.dll"

Function boolean CloseHandle ( &
	longptr hObject &
	) Library "kernel32.dll"

end prototypes

type variables
Constant uint CF_BITMAP = 2
Constant uint DIB_PAL_COLORS = 1
Constant uint DIB_RGB_COLORS = 0
Constant integer BITMAPTYPE = 19778

// raster operation codes
Constant Long BLACKNESS			= 66
Constant Long CAPTUREBLT		= 1073741824
Constant Long DSTINVERT			= 5570569
Constant Long MERGECOPY			= 12583114
Constant Long MERGEPAINT		= 12255782
Constant Long NOMIRRORBITMAP	= 2147483648
Constant Long NOTSRCCOPY		= 3342344
Constant Long NOTSRCERASE		= 1114278
Constant Long PATCOPY			= 15728673
Constant Long PATINVERT			= 5898313
Constant Long PATPAINT			= 16452105
Constant Long SRCAND				= 8913094
Constant Long SRCCOPY			= 13369376
Constant Long SRCERASE			= 4457256
Constant Long SRCINVERT			= 6684742
Constant Long SRCPAINT			= 15597702
Constant Long WHITENESS			= 16711778

end variables

forward prototypes
public function string of_gettemppath ()
public function unsignedlong of_writeblob (string as_filename, blob ablb_bitmap)
public function blob of_controlcapture (dragobject a_object, boolean ab_clipboard)
public function blob of_screencapture (boolean ab_clipboard)
public function string of_getlasterror ()
public function decimal of_dpifactor ()
public function blob of_capture (longptr al_hwnd, unsignedlong al_xpos, unsignedlong al_ypos, unsignedlong al_width, unsignedlong al_height, boolean ab_clipboard)
public function blob of_windowcapture (window aw_window, boolean ab_clipboard)
public function blob of_sheetcapture (mdiclient am_mdi, window aw_sheet, boolean ab_clipboard)
end prototypes

public function string of_gettemppath ();// returns temp directory name

String ls_path
Long ll_rtn, ll_buflen

ll_buflen = 250
ls_path = Space(ll_buflen)

ll_rtn = GetTempPath(ll_buflen, ls_path)

Return ls_path

end function

public function unsignedlong of_writeblob (string as_filename, blob ablb_bitmap);Constant Long INVALID_HANDLE_VALUE = -1
Constant ULong GENERIC_WRITE		= 1073741824
Constant ULong FILE_SHARE_WRITE	= 2
Constant ULong CREATE_ALWAYS		= 2

Longptr ll_file
ULong lul_length, lul_written
Boolean lb_rtn

// open file for write
ll_file = CreateFile(as_filename, GENERIC_WRITE, &
					FILE_SHARE_WRITE, 0, CREATE_ALWAYS, 0, 0)
If ll_file = INVALID_HANDLE_VALUE Then
	Return -999
End If

// write file to disk
lul_length = Len(ablb_bitmap)
lb_rtn = WriteFile(ll_file, ablb_bitmap, &
					lul_length, lul_written, 0)

// close the file
CloseHandle(ll_file)

Return 0

end function

public function blob of_controlcapture (dragobject a_object, boolean ab_clipboard);// capture bitmap of a control

PowerObject lpo_parent
Long ll_hWnd, ll_XPos, ll_YPos, ll_Width, ll_Height

// loop thru parents until a window is found
lpo_parent = a_object.GetParent()
Do While lpo_parent.TypeOf() <> Window! and IsValid (lpo_parent)
	lpo_parent = lpo_parent.GetParent()
Loop

// get handle to window
ll_hWnd = Handle(lpo_parent)

// convert x, y, Width and Height from PBU to Pixels
ll_XPos   = UnitsToPixels(a_object.X, XUnitsToPixels!)
ll_YPos   = UnitsToPixels(a_object.Y, YUnitsToPixels!)
ll_Width  = UnitsToPixels(a_object.Width, XUnitsToPixels!)
ll_Height = UnitsToPixels(a_object.Height, YUnitsToPixels!)

Return this.of_Capture(ll_hWnd, ll_XPos, ll_YPos, &
								ll_Width, ll_Height, ab_clipboard)

end function

public function blob of_screencapture (boolean ab_clipboard);// capture bitmap of entire screen

Environment le_env
Long ll_hWnd, ll_Width, ll_Height

// get handle to windows background
ll_hWnd = GetDesktopWindow()

// get screen size
GetEnvironment(le_env)
ll_Width  = le_env.ScreenWidth  * of_dpiFactor()
ll_Height = le_env.ScreenHeight * of_dpiFactor()

Return this.of_Capture(ll_hWnd, 0, 0, &
								ll_Width, ll_Height, ab_clipboard)

end function

public function string of_getlasterror ();// This function returns the most recent API error message

Constant ULong FORMAT_MESSAGE_FROM_SYSTEM = 4096
Constant ULong LANG_NEUTRAL = 0
ULong lul_rtn, lul_error
String ls_msgtext

lul_error = GetLastError()

ls_msgtext = Space(200)

lul_rtn = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, &
				lul_error, LANG_NEUTRAL, ls_msgtext, 200, 0)

Return ls_msgtext

end function

public function decimal of_dpifactor ();//// Determine the DPI Scaling Factor
//
//Decimal ldec_dpiFactor
//String ls_regkey, ls_monitor[]
//ULong lul_DpiValue
//
//// get list of monitor keys
//ls_regkey = "HKEY_CURRENT_USER\Control Panel\Desktop\PerMonitorSettings"
//RegistryKeys(ls_regkey, ls_monitor)
//If UpperBound(ls_monitor) = 0 Then Return 1
//
//// get DpiValue from first one
//ls_regkey += "\" + ls_monitor[1]
//RegistryGet(ls_regkey, "DpiValue", ReguLong!, lul_DpiValue)
//
//// determine DPI Scaling Factor
//choose case lul_DpiValue
//	case 0
//		ldec_dpiFactor = 1.25
//	case 1
//		ldec_dpiFactor = 1.50
//	case 2
//		ldec_dpiFactor = 1.75
//	case 3
//		ldec_dpiFactor = 2.00
//	case else
//		ldec_dpiFactor = 1.00
//end choose
//
//Return ldec_dpiFactor

return 1

end function

public function blob of_capture (longptr al_hwnd, unsignedlong al_xpos, unsignedlong al_ypos, unsignedlong al_width, unsignedlong al_height, boolean ab_clipboard);// capture bitmap and return as blob

BitmapInfo lstr_Info
BitmapFileHeader lstr_Header
Blob lblb_header, lblb_info, lblb_bitmap
Longptr ll_hdc, ll_hdcMem, ll_hBitmap
Integer li_pixels
Boolean lb_result

// Get the device context of window and allocate memory
ll_hdc = GetDC(al_hWnd)
ll_hdcMem = CreateCompatibleDC(ll_hdc)
ll_hBitmap = CreateCompatibleBitmap(ll_hdc, al_width, al_height)

If ll_hBitmap <> 0 Then
	// Select an object into the specified device context
	SelectObject(ll_hdcMem, ll_hBitmap)
	// Copy the bitmap from the source to the destination
   lb_result = BitBlt(ll_hdcMem, 0, 0, al_width, al_height, &
									ll_hdc, al_xpos, al_ypos, SRCCOPY)
	// try to store the bitmap into a blob so we can save it
	lstr_Info.bmiHeader.biSize = 40
	// Get the bitmapinfo
	If GetDIBits(ll_hdcMem, ll_hBitmap, 0, al_height, &
							0, lstr_Info, DIB_RGB_COLORS) > 0 Then
		// allocate a blob to hold the image pixels
		li_pixels = lstr_Info.bmiHeader.biBitCount
		lstr_Info.bmiColors[li_pixels] = 0
		lblb_bitmap = Blob(Space(lstr_Info.bmiHeader.biSizeImage/2))
		// get the actual bits
		GetDIBits(ll_hdcMem, ll_hBitmap, 0, al_height, &
							lblb_bitmap, lstr_Info, DIB_RGB_COLORS) 
		// create a bitmap header
		lstr_Header.bfType = BITMAPTYPE
		lstr_Header.bfSize = lstr_Info.bmiHeader.biSizeImage
		lstr_Header.bfOffBits = 54 + (li_pixels * 4)
		// copy the header structure to a blob
		lblb_header = Blob(Space(14/2))
		CopyBitmapFileHeader(lblb_header, lstr_Header, 14)
		// copy the info structure to a blob
		lblb_Info = Blob(Space((40  + li_pixels * 4)/2))
		CopyBitmapInfo(lblb_Info, lstr_Info, 40 + li_pixels * 4)
		// add all together and we have a window bitmap in a blob
		lblb_bitmap = lblb_header + lblb_info + lblb_bitmap
	End If

	// paste bitmap to clipboard
	If ab_clipboard Then
		If OpenClipboard(al_hWnd) Then
			EmptyClipboard()
			SetClipboardData(CF_BITMAP, ll_hBitmap)
			CloseClipboard()
		Else
			MessageBox("OpenClipboard Failed", of_GetLastError())
		End If
	End If
End If

// Clean up handles
DeleteDC(ll_hdcMem)
ReleaseDC(al_hWnd, ll_hdc)

Return lblb_bitmap

end function

public function blob of_windowcapture (window aw_window, boolean ab_clipboard);// capture bitmap of a window

Long ll_hWnd, ll_XPos, ll_YPos, ll_Width, ll_Height

// get handle to windows background
ll_hWnd = GetDesktopWindow()

// convert x, y, Width and Height from PBU to Pixels
ll_XPos   = UnitsToPixels(aw_window.X, XUnitsToPixels!) * of_dpiFactor()
ll_YPos   = UnitsToPixels(aw_window.Y, YUnitsToPixels!) * of_dpiFactor()
ll_Width  = UnitsToPixels(aw_window.Width, XUnitsToPixels!)  * of_dpiFactor()
ll_Height = UnitsToPixels(aw_window.Height, YUnitsToPixels!) * of_dpiFactor()

Return this.of_Capture(ll_hWnd, ll_XPos, ll_YPos, &
								ll_Width, ll_Height, ab_clipboard)

end function

public function blob of_sheetcapture (mdiclient am_mdi, window aw_sheet, boolean ab_clipboard);// capture bitmap of a sheet window

Window lw_parent
Long ll_hWnd, ll_pbunits, ll_XPos, ll_YPos, ll_Width, ll_Height

// get reference to frame window
lw_parent = am_mdi.GetParent()

// get handle to window
ll_hWnd = Handle(lw_parent)

// convert x, y, Width and Height from PBU to Pixels
ll_XPos   = UnitsToPixels(am_mdi.X + aw_sheet.X + 8, XUnitsToPixels!)
ll_YPos   = UnitsToPixels(am_mdi.Y + aw_sheet.Y + 8, YUnitsToPixels!)
ll_Width  = UnitsToPixels(aw_sheet.Width, XUnitsToPixels!)
ll_Height = UnitsToPixels(aw_sheet.Height, YUnitsToPixels!)

Return this.of_Capture(ll_hWnd, ll_XPos, ll_YPos, &
								ll_Width, ll_Height, ab_clipboard)

end function

on n_bitmap.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_bitmap.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

