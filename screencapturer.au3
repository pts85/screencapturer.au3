; Script captures selected area of screen and saves it to imagefile and optionally save coordinates to txt-file
; Script is a modified version of a script found in autoit-forum, where it was released under PUBLIC DOMAIN by user Melba23
; Original: https://www.autoitscript.com/forum/topic/117114-capture-mouse-selection/?do=findComment&comment=816497
;
; Modifications: support for: closing child window, showing and saving of coordinates and size, changing of filename
; TODO(maybe sometime): Some error handling

#include <GuiConstantsEx.au3>
#include <WindowsConstants.au3>
#Include <ScreenCapture.au3>
#Include <Misc.au3>
#include <EditConstants.au3>

Global $iX1, $iY1, $iX2, $iY2, $aPos, $sMsg, $sBMP_Path

; Create GUI
Global $hMain_GUI = GUICreate("Select Rectangle", 230, 130)
Global $hRect_Button   = GUICtrlCreateButton("Mark Area",  10, 10, 60, 30)
Global $hClose_Button = GUICtrlCreateButton("Close",    160, 10, 60, 30)

Global $hEditbox1 = GUICtrlCreateEdit ( "", 80, 1 , 70 , 60 , $ES_READONLY , -1 )
GUICtrlSetData($hEditbox1, "       X" & @CRLF &"        |" & @CRLF & "------- | -------- Y" & @CRLF & "        |" & @CRLF & "        |", 1)
Global $hCoords_Editbox = GUICtrlCreateEdit ( "", 10, 63 , 210 , 35 , $ES_READONLY , -1 )
Global $hPath_Editbox = GUICtrlCreateEdit ( @DesktopDir, 10, 100 , 110 , 25 , $ES_AUTOHSCROLL , -1 )
GUICtrlSetTip($hPath_Editbox, "Directory for screencapturefiles")
Global $hFilename_Editbox = GUICtrlCreateEdit ( "filename.bmp", 120, 100 , 100 , 25 , $ES_AUTOHSCROLL , -1 )
GUICtrlSetTip($hFilename_Editbox, "filename.extension (BMP, GIF, JPEG, PNG, TIF)")
Global $hTXT_Checkbox = GUICtrlCreateCheckbox("Save coords", 153, 41, 80, 20)
GUICtrlSetTip($hTXT_Checkbox, "Check to save coordinates to .TXT file")

GUISetState()

Global $SizeX, $SizeY, $capturefilename, $savecoordstoo
Global $hBitmap_GUI, $aMsg

While 1
   $aMsg = GUIGetMsg(1)
   Switch $aMsg[1]

   Case $hMain_GUI
	  Switch $aMsg[0]
		 Case $GUI_EVENT_CLOSE, $hClose_Button
            Exit
		 Case $hRect_Button
			$capturefilepath = GUICtrlRead($hPath_Editbox, $GUI_READ_EXTENDED)
			$capturefilename = GUICtrlRead($hFilename_Editbox, $GUI_READ_EXTENDED)
            GUISetState(@SW_HIDE, $hMain_GUI)
            Mark_Rect()
            ; Capture selected area
            $sBMP_Path = $capturefilepath & "\" & $capturefilename
			$SizeX = $iX2 - $iX1
			$SizeY = $iY2 - $IY1
			Local $coords_str = ("X1=" & $iX1 & " X2=" & $iX2 & " Size=" & $SizeX & @CRLF & "Y1=" & $iY1 &  " Y2=" & $iY2 & " Size=" & $SizeY)
			GUICtrlSetData($hCoords_Editbox,"");
			GUICtrlSetData($hCoords_Editbox, $coords_str, 1)
			If $savecoordstoo == 1 Then
				  $hTXTfile = FileOpen ($capturefilepath & "\" & $capturefilename & ".txt", 2 ) ; 2=erase contents
				  FileWriteLine($hTXTfile, $coords_str);
				  FileClose($hTXTfile)
			EndIf
            _ScreenCapture_Capture($sBMP_Path, $iX1, $iY1, $iX2, $iY2, False)
            GUISetState(@SW_SHOW, $hMain_GUI)
            ; Display image
            $hBitmap_GUI = GUICreate("File: " & $capturefilename, $iX2 - $iX1 + 1, $iY2 - $iY1 + 1, 100, 100)
            $hPic = GUICtrlCreatePic($capturefilepath & "\" & $capturefilename, 0, 0, $iX2 - $iX1 + 1, $iY2 - $iY1 + 1)
            GUISetState()
		 Case $hTXT_Checkbox
			If GuiCtrlRead($hTXT_Checkbox) = $GUI_CHECKED Then
			   $savecoordstoo = 1;
			Else
			   $savecoordstoo = 0;
			EndIf
	  EndSwitch ;==> aMsg[0]

   Case $hBitmap_GUI
	  Switch $aMsg[0]
	   Case $GUI_EVENT_CLOSE
			   GuiDelete($hBitmap_GUI)
			   GUICtrlSetData($hCoords_Editbox,"");
	  EndSwitch  ;==> aMsg[0]
   EndSwitch ;==> aMsg[1]

WEnd

; -------------

Func Mark_Rect()

    Local $aMouse_Pos, $hMask, $hMaster_Mask, $iTemp
    Local $UserDLL = DllOpen("user32.dll")

    Global $hRectangle_GUI = GUICreate("", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
    _GUICreateInvRect($hRectangle_GUI, 0, 0, 1, 1)
    GUISetBkColor(0)
    WinSetTrans($hRectangle_GUI, "", 50)
    GUISetState(@SW_SHOW, $hRectangle_GUI)
    GUISetCursor(3, 1, $hRectangle_GUI)

    ; Wait until mouse button pressed
    While Not _IsPressed("01", $UserDLL)
        Sleep(10)
    WEnd

    ; Get first mouse position
    $aMouse_Pos = MouseGetPos()
    $iX1 = $aMouse_Pos[0]
    $iY1 = $aMouse_Pos[1]

    ; Draw rectangle while mouse button pressed
    While _IsPressed("01", $UserDLL)

        $aMouse_Pos = MouseGetPos()

        ; Set in correct order if required
        If $aMouse_Pos[0] < $iX1 Then
            $iX_Pos = $aMouse_Pos[0]
            $iWidth = $iX1 - $aMouse_Pos[0]
        Else
            $iX_Pos = $iX1
            $iWidth = $aMouse_Pos[0] - $iX1
        EndIf
        If $aMouse_Pos[1] < $iY1 Then
            $iY_Pos = $aMouse_Pos[1]
            $iHeight = $iY1 - $aMouse_Pos[1]
        Else
            $iY_Pos = $iY1
            $iHeight = $aMouse_Pos[1] - $iY1
        EndIf

        _GUICreateInvRect($hRectangle_GUI, $iX_Pos, $iY_Pos, $iWidth, $iHeight)

        Sleep(10)

    WEnd

    ; Get second mouse position
    $iX2 = $aMouse_Pos[0]
    $iY2 = $aMouse_Pos[1]

    ; Set in correct order if required
    If $iX2 < $iX1 Then
        $iTemp = $iX1
        $iX1 = $iX2
        $iX2 = $iTemp
    EndIf
    If $iY2 < $iY1 Then
        $iTemp = $iY1
        $iY1 = $iY2
        $iY2 = $iTemp
    EndIf

    GUIDelete($hRectangle_GUI)
    DllClose($UserDLL)

EndFunc   ;==>Mark_Rect

Func _GUICreateInvRect($hWnd, $iX, $iY, $iW, $iH)

    $hMask_1 = _WinAPI_CreateRectRgn(0, 0, @DesktopWidth, $iY)
    $hMask_2 = _WinAPI_CreateRectRgn(0, 0, $iX, @DesktopHeight)
    $hMask_3 = _WinAPI_CreateRectRgn($iX + $iW, 0, @DesktopWidth, @DesktopHeight)
    $hMask_4 = _WinAPI_CreateRectRgn(0, $iY + $iH, @DesktopWidth, @DesktopHeight)

    _WinAPI_CombineRgn($hMask_1, $hMask_1, $hMask_2, 2)
    _WinAPI_CombineRgn($hMask_1, $hMask_1, $hMask_3, 2)
    _WinAPI_CombineRgn($hMask_1, $hMask_1, $hMask_4, 2)

    _WinAPI_DeleteObject($hMask_2)
    _WinAPI_DeleteObject($hMask_3)
    _WinAPI_DeleteObject($hMask_4)

    _WinAPI_SetWindowRgn($hWnd, $hMask_1, 1)

EndFunc
