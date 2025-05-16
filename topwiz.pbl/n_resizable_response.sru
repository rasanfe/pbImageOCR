forward
global type n_resizable_response from nonvisualobject
end type
end forward

global type n_resizable_response from nonvisualobject autoinstantiate
end type

type prototypes
Function boolean DrawMenuBar ( &
	long hWnd &
	) Library "user32.dll"

Function long GetSystemMenu ( &
	long hWnd, &
	boolean bRevert &
	) Library "user32.dll"

Function long GetSystemMetrics ( &
	long nIndex &
	) Library "user32.dll"

Function long GetWindowLong ( &
	long hWnd, &
	long nIndex &
	) Library "user32.dll" Alias For "GetWindowLongW"

Function boolean InsertMenu ( &
	long hMenu, &
	uint uPosition, &
	uint uFlags, &
	uint uIDNewItem, &
	string lpNewItem &
	) Library "user32.dll" Alias For "InsertMenuW"

Function long SetWindowLong ( &
	long hWnd, &
	long nIndex, &
	long dwNewLong &
	) Library "user32.dll" Alias For "SetWindowLongW"

end prototypes

type variables
Private:

Constant Long GWL_STYLE = -16
Constant Long WS_THICKFRAME = 262144
Constant Long WS_MINIMIZEBOX = 65536
Constant Long MF_BYCOMMAND = 0
Constant Long MF_BYPOSITION = 1024
Constant Long MF_STRING = 0
Constant Long SC_RESTORE = 61728
Constant Long SC_MAXIMIZE = 61488
Constant Long SC_MINIMIZE = 61472
Constant Long SC_SIZE = 61440
Constant Long SM_CXSIZEFRAME = 32
Constant Long SM_CYSIZEFRAME = 33
Constant Long SM_CYCAPTION = 4

Public:

Boolean IsResizable
Integer iOrigXPixels
Integer iOrigYPixels

end variables

forward prototypes
public subroutine of_setresizable (window aw_response)
end prototypes

public subroutine of_setresizable (window aw_response);// -----------------------------------------------------------------------------
// SCRIPT:     n_resizeable_response.of_SetResizable
//
// PURPOSE:    This function makes the passed response window resizable. The
//					function is called from the window open event passing 'this'.
//					The window should also implement the GetMinMaxInfo event so
//					that it can set original size to be the minimum size.
//
//					The code is based on examples/discussion between Bruce Armstrong
//					and Dan Cooperstock on the SAP PowerBuilder forum:
//					http://scn.sap.com/thread/3266956
//
// ARGUMENTS:  aw_response	- The response window
//
// DATE        PROG/ID		DESCRIPTION OF CHANGE / REASON
// ----------  --------		-----------------------------------------------------
// 01/11/2015	RolandS		Initial Coding
// -----------------------------------------------------------------------------

Long ll_hWnd, ll_hMenu
Long ll_style, ll_newstyle
Integer li_XSize, li_YSize, li_Title
Integer li_NewWidth, li_NewHeight

aw_response.SetRedraw(False)
 
ll_hWnd = Handle(aw_response)

// Get size of the window borders
li_XSize = PixelsToUnits(GetSystemMetrics(SM_CXSIZEFRAME), XPixelsToUnits!)
li_YSize = PixelsToUnits(GetSystemMetrics(SM_CYSIZEFRAME), YPixelsToUnits!)
li_Title = PixelsToUnits(GetSystemMetrics(SM_CYCAPTION),   YPixelsToUnits!) 

// calculate Width/Height to allow the window body to remain the same size
li_NewWidth  = aw_response.WorkSpaceWidth()  + (li_XSize * 2)
li_NewHeight = aw_response.WorkSpaceHeight() + (li_YSize * 2) + li_Title 

// Save original size of window in pixels for minimum size event
iOrigXPixels  = UnitsToPixels(li_NewWidth,  XUnitsToPixels!)
iOrigYPixels  = UnitsToPixels(li_NewHeight, YUnitsToPixels!)

// Get the current window style
ll_style = GetWindowLong(ll_hWnd, GWL_STYLE)

If aw_response.ControlMenu Then
	// You have to include MINIMIZEBOX for the controls to show
	ll_newstyle = ll_style + WS_THICKFRAME + WS_MINIMIZEBOX
Else
	ll_newstyle = ll_style + WS_THICKFRAME
End If

If ll_style <> 0 Then
	If SetWindowLong(ll_hWnd, GWL_STYLE, ll_newstyle) <> 0 Then
		If aw_response.ControlMenu Then
			// Add items to the system menu
			ll_hMenu = GetSystemMenu(ll_hWnd, False)
			If ll_hMenu > 0 Then
				InsertMenu(ll_hMenu, 1, MF_BYPOSITION + MF_STRING, &
						SC_MAXIMIZE, "Maximize")
				InsertMenu(ll_hMenu, 1, MF_BYPOSITION + MF_STRING, &
						SC_RESTORE, "Restore")
				InsertMenu(ll_hMenu, 1, MF_BYPOSITION + MF_STRING, &
						SC_SIZE, "Size")
				DrawMenuBar(ll_hWnd)
			End If
		End If
		// Return to original size
		aw_response.Resize(li_NewWidth, li_NewHeight)
		IsResizable = True
	End If
End If

aw_response.SetRedraw(True)

end subroutine

on n_resizable_response.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_resizable_response.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

