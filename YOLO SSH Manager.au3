#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/so /sf /sv /soi /mi
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Bitvise.au3>
#include <File.au3>
#include <GUIlistview.au3>
#include <String.au3>
Opt("GUIOnEventMode", 1)
;~ Opt("MustDeclareVars", 1)
Opt("TrayIconHide", 1)
Global $BitviseSSH = IniRead(@ScriptDir & "\Options.ini", "Options", "BitviseDir", "C:\Program Files\Bitvise SSH Client")
Global $Main, $Listview, $Delaytime, $ConnectAll, $ConnectN, $hCheckbox, $hFeedback
Global $contextmenu, $cLogin, $cLogout, $iHide, $HideBitvise, $UseStartPort, $iStartPort
Global $SSHList, $AutoConnectAll, $Title, $hCheckbox2, $hStartPort
Global $StartPort = IniRead(@ScriptDir & "\Options.ini", "Options", "StartPort", 1080)

$Title = "YOLO SSH Manager v1.2"

$Main = GUICreate($Title, 610, 418, -1, -1, -1, 0x00000010)
GUISetOnEvent(-3, "_Exit")
GUISetOnEvent(-13, "Drag")
GUISetBkColor(0x0080FF)
$Listview = GUICtrlCreateListView("IP|User|Password|Port|Status|PiD", 3, 3, 603, 342, BitOR(0x0008, 0x0001))
GUICtrlSendMsg(-1, 0x1000 + 30, 0, 105)
GUICtrlSendMsg(-1, 0x1000 + 30, 1, 105)
GUICtrlSendMsg(-1, 0x1000 + 30, 2, 105)
GUICtrlSendMsg(-1, 0x1000 + 30, 3, 100)
GUICtrlSendMsg(-1, 0x1000 + 30, 4, 90)
GUICtrlSendMsg(-1, 0x1000 + 30, 5, 90)
GUICtrlSetState(-1, 8)
GUICtrlCreateLabel("Delay Time (ms)", 17, 363, 100, 16)
GUICtrlSetColor(-1, 0xFFFFFF)
GUICtrlSetFont(-1, 8.5, 800, 0, "Tahoma")
$Delaytime = GUICtrlCreateInput(1000, 129, 359, 50, 21, 8192)
GUICtrlSetFont(-1, 9, 800, 0, "Tahoma")
$ConnectAll = GUICtrlCreateButton("Connect All", 393, 359, 196, 47)
GUICtrlSetFont(-1, 8.5, 800, 0)
GUICtrlCreateLabel("Connect Number", 18, 390, 105, 16)
GUICtrlSetFont(-1, 8.5, 800, 0, "Tahoma")
GUICtrlSetColor(-1, 0xFFFFFF)
$ConnectN = GUICtrlCreateInput(13, 129, 388, 51, 21, 8192)
GUICtrlSetFont(-1, 9, 800, 0, "Tahoma")
$hCheckbox = GUICtrlCreateCheckbox("Hide Bitvise", 228, 360, 93, 19)
GUICtrlSetOnEvent(-1, "_HideBitvise")
GUICtrlSetFont(-1, 8.5, 800, 0)
GUICtrlSetState(-1, 0)
$contextmenu = GUICtrlCreateContextMenu($Listview)
$cLogin = GUICtrlCreateMenuItem("Login", $contextmenu)
$cLogout = GUICtrlCreateMenuItem("Logout", $contextmenu)

$hCheckbox2 = GUICtrlCreateCheckbox("Startport", 228, 391, 70, 15)
GUICtrlSetOnEvent(-1, "_StartPort")
GUICtrlSetFont(-1, 8.5, 800, 0, "Tahoma")
$hStartPort = GUICtrlCreateInput($StartPort, 311, 388, 50, 20, 8192)
GUICtrlSetFont(-1, 8.5, 800, 0, "Tahoma", 5)
GUICtrlSetLimit(-1, 8080, 1080)
GUISetState()

If $CmdLine[0] = 1 Then
	$SSHList = $CmdLine[1]
	$AutoConnectAll = False
ElseIf $CmdLine[0] = 2 Then
	If $CmdLine[2] = "True" Then
		$SSHList = $CmdLine[1]
		$AutoConnectAll = True
	Else
		$SSHList = ""
		$AutoConnectAll = False
	EndIf
Else
	$SSHList = ""
	$AutoConnectAll = False
EndIf

$HideBitvise = IniRead(@ScriptDir & "\Options.ini", "Options", "HideBitvise", 1)
If Number($HideBitvise) = 1 Then
	GUICtrlSetState($hCheckbox, 1)
	$iHide = True
Else
	GUICtrlSetState($hCheckbox, 4)
	$iHide = False
EndIf

$UseStartPort = IniRead(@ScriptDir & "\Options.ini", "Options", "StartPortEnable", 1)
If Number($UseStartPort) = 1 Then
	GUICtrlSetState($hCheckbox2, 1)
	$iStartPort = True
Else
	GUICtrlSetState($hCheckbox, 4)
	$iStartPort = False
EndIf

_ListView()
_CheckRunning()
GUIRegisterMsg(0x004E, "WM_NOTIFY")
If $AutoConnectAll = True Then
	GUICtrlSetOnEvent($cLogin, "_cLogin")
	GUICtrlSetOnEvent($cLogout, "_cLogout")
	_ConnectAll()
Else
	GUICtrlSetOnEvent($cLogin, "_cLogin")
	GUICtrlSetOnEvent($cLogout, "_cLogout")
	GUICtrlSetOnEvent($ConnectAll, "_ConnectAll")
EndIf

Func _ConnectAll()
	Local $SSH, $iSSH, $ConnectNum, $iSleep, $FilePath, $Data
	$ConnectNum = GUICtrlRead($ConnectN) - 1
	$iSleep = GUICtrlRead($Delaytime)
	If $SSHList = "" Then
		$FilePath = IniRead(@ScriptDir & "\Options.ini", "Options", "SSHPath", @ScriptDir & "\SSH list.ini")
	Else
		$FilePath = $SSHList
	EndIf
	If $FilePath = "" Then $FilePath = @ScriptDir & "\SSH list.ini"
	ConsoleWrite($FilePath & @CRLF)
	If FileRead($FilePath) = "" Then
		Return
	EndIf

	For $i = 1 To _FileCountLines($FilePath)
		$SSH = FileReadLine($FilePath, $i)
		$SSH = StringStripWS($SSH, 8)
		$iSSH = _StringSplitSSH($SSH, $StartPort - 1 + $i)
		If IsArray($iSSH) Then
			Local $HexToString = '0000000E54756E6E656C69657220342E35320000000000000016'
			$HexToString &= '00000000000000000000000000000000000000000000000B627364617574682C70616D01010001020000000200'
			$HexToString &= '000000000005787465726D010000FDE900000050000000190000012C07010000000000000000000D3132372E302'
			$HexToString &= 'E302E313A302E30000000000000000000000000000000093132372E302E302E3100000D3D000000000000000000'
			$HexToString &= '00000000000000010000010101010101010101000001010101000001010101000000012C0100000000000000000'
			$HexToString &= '000017F0000010000000431303830000000000000000000007F0000010000000232310000000001010100000000'
			$HexToString &= '0000000000000000000000010100000001010000000000000000000000000000000000000200'
			Local $DefaultProfile = _HexToString($HexToString)
			RegWrite("HKEY_CURRENT_USER\Software\Bitvise" & $iSSH[3] & "\BvSshClient", 'DefaultProfile', 'REG_BINARY', $DefaultProfile)
			Local $x = RegRead("HKEY_CURRENT_USER\Software\Bitvise" & $iSSH[3] & "\BvSshClient", "DefaultProfile")
			Local $xy = StringRegExpReplace($x, '(7F00000100000004(?:.*?)000000000000000000007F)', '7F00000100000004' & StringTrimLeft(StringToBinary($iSSH[3]), 2) & '000000000000000000007F', 1)
			RegWrite("HKEY_CURRENT_USER\Software\Bitvise" & $iSSH[3] & "\BvSshClient", "DefaultProfile", "REG_BINARY", $xy)
		Else
			MsgBox(48, $Title, "Wrong SSH Format." & @CRLF & "IP | USER | PASS")
			Return
		EndIf
	Next

	WinSetTitle($Main, "", $Title & " - Total SSH: " & _FileCountLines($FilePath) & ' - Start Connect...')
	For $i = 1 To _FileCountLines($FilePath)
		_GUICtrlListView_AddSubItem($Listview, $i - 1, "Starting", 4, 1)
		$SSH = FileReadLine($FilePath, $i)
		$SSH = StringStripWS($SSH, 8)
		$iSSH = _StringSplitSSH($SSH, $StartPort - 1 + $i)
		If IsArray($iSSH) Then
			_Login($iSSH[0], $iSSH[1], $iSSH[2], $iSSH[3], $i - 1)
			Sleep($iSleep / 2)
		Else
			MsgBox(48, $Title, "Wrong SSH Format." & @CRLF & "IP | USER | PASS")
			Return
		EndIf
	Next

	For $i = 1 To _FileCountLines($FilePath)
		$SSH = FileReadLine($FilePath, $i)
		$SSH = StringStripWS($SSH, 8)
		$iSSH = _StringSplitSSH($SSH, $StartPort - 1 + $i)
		If IsArray($iSSH) Then
			If _Bitvise_Wait_For_Connection($iSSH[3], 3000) = 0 Then
				_GUICtrlListView_AddSubItem($Listview, $i - 1, "Died", 4)
			Else
				_GUICtrlListView_AddSubItem($Listview, $i - 1, "Ready to use", 4)
			EndIf
		Else
			MsgBox(48, $Title, "Wrong SSH Format." & @CRLF & "IP | USER | PASS")
			Return
		EndIf
	Next
	WinSetTitle($Main, "", $Title & " - Total SSH: " & _FileCountLines($FilePath) & ' - Ready!')
EndFunc   ;==>_ConnectAll

Func _Login($iHost, $iUser, $iPass, $iPort, $i)
	Local $Cmdlines
	$Cmdlines = $BitviseSSH & "\BvSsh.exe -host=" & $iHost & " -port=22"
	$Cmdlines &= " -user=" & $iUser & " -password=" & $iPass & " -loginOnStartup -title=iPort:" & $iPort & " -baseRegistry=HKEY_CURRENT_USER\Software\Bitvise" & $iPort
	If $iHide = True Then
		$Cmdlines &= " -menu=small -hide=popups,trayLog,trayPopups,trayIcon"
	EndIf

	Local $PiD = Run($Cmdlines)
;~ 	ConsoleWrite($Cmdlines & @CRLF)
	_GUICtrlListView_AddSubItem($Listview, $i, $PiD, 5)
EndFunc   ;==>_Login

Func _Logout($i, $PiD)
	_GUICtrlListView_AddSubItem($Listview, $i, "Logout", 4, 1)
	RunWait($BitviseSSH & "\BvSshCtrl.exe " & $PiD & " logout", "", @SW_HIDE)
	_GUICtrlListView_AddSubItem($Listview, $i, "", 4)
	_GUICtrlListView_AddSubItem($Listview, $i, "", 5)
EndFunc   ;==>_Logout

Func _StringSplitSSH($String, $port)
	Local $SSH[4]
	If StringInStr($String, '|') Then
		Local $Split
		$Split = StringSplit($String, '|', 1)
		If $Split[0] >= 4 Then
			$SSH[0] = $Split[1]
			$SSH[1] = $Split[2]
			$SSH[2] = $Split[3]
			If $iStartPort = False Then
				$SSH[3] = $Split[4]
			Else
				$SSH[3] = $port
			EndIf
		EndIf
		Return $SSH
	EndIf
EndFunc   ;==>_StringSplitSSH

Func _ListView()
	Local $FilePath
	If $SSHList = "" Then
		$FilePath = IniRead(@ScriptDir & "\Options.ini", "Options", "SSHPath", @ScriptDir & "\SSH list.ini")
	Else
		$FilePath = $SSHList
	EndIf
	If $FilePath = "" Then $FilePath = @ScriptDir & "\SSH list.ini"
	_GUICtrlListView_DeleteAllItems($Listview)
	LoadListView($FilePath, $Listview)
	WinSetTitle($Main, "", $Title & " - Total SSH: " & _FileCountLines($FilePath) & ' - Ready!')
EndFunc   ;==>_ListView

Func LoadListView($FilePath, $Listview)
	Local $SSH, $iSSH
	For $i = 1 To _FileCountLines($FilePath)
		$SSH = FileReadLine($FilePath, $i)
		$SSH = StringStripWS($SSH, 8)
		$iSSH = _StringSplitSSH($SSH, $StartPort - 1 + $i)
		If IsArray($iSSH) Then
			GUICtrlCreateListViewItem($iSSH[0] & '|' & $iSSH[1] & '|' & $iSSH[2] & '|' & $iSSH[3], $Listview)
		Else
			MsgBox(48, $Title, "Wrong SSH Format." & @CRLF & "IP | USER | PASS")
			Return
		EndIf
	Next
EndFunc   ;==>LoadListView

Func _cLogin()
	Local $aItem, $sText, $iSelect
	$iSelect = _GUICtrlListView_GetSelectedIndices($Listview, True)
	If $iSelect[0] > 0 Then
		For $i = 1 To $iSelect[0]
			$aItem = _GUICtrlListView_GetItemTextArray($Listview, $iSelect[$i])
			_GUICtrlListView_AddSubItem($Listview, $iSelect[$i], "Starting", 4, 1)
			Local $HexToString = '0000000E54756E6E656C69657220342E35320000000000000016'
			$HexToString &= '00000000000000000000000000000000000000000000000B627364617574682C70616D01010001020000000200'
			$HexToString &= '000000000005787465726D010000FDE900000050000000190000012C07010000000000000000000D3132372E302'
			$HexToString &= 'E302E313A302E30000000000000000000000000000000093132372E302E302E3100000D3D000000000000000000'
			$HexToString &= '00000000000000010000010101010101010101000001010101000001010101000000012C0100000000000000000'
			$HexToString &= '000017F0000010000000431303830000000000000000000007F0000010000000232310000000001010100000000'
			$HexToString &= '0000000000000000000000010100000001010000000000000000000000000000000000000200'
			Local $DefaultProfile = _HexToString($HexToString)
			RegWrite("HKEY_CURRENT_USER\Software\Bitvise" & $aItem[4] & "\BvSshClient", 'DefaultProfile', 'REG_BINARY', $DefaultProfile)
			Local $x = RegRead("HKEY_CURRENT_USER\Software\Bitvise" & $aItem[4] & "\BvSshClient", "DefaultProfile")
			Local $xy = StringRegExpReplace($x, '(7F00000100000004(?:.*?)000000000000000000007F)', '7F00000100000004' & StringTrimLeft(StringToBinary($aItem[4]), 2) & '000000000000000000007F', 1)
			RegWrite("HKEY_CURRENT_USER\Software\Bitvise" & $aItem[4] & "\BvSshClient", "DefaultProfile", "REG_BINARY", $xy)
			_Login($aItem[1], $aItem[2], $aItem[3], $aItem[4], $iSelect[$i])
		Next
		For $i = 1 To $iSelect[0]
			If _Bitvise_Wait_For_Connection($aItem[4]) = 0 Then
				_GUICtrlListView_AddSubItem($Listview, $iSelect[$i], "Died", 4)
			Else
				_GUICtrlListView_AddSubItem($Listview, $iSelect[$i], "Ready to use", 4)
			EndIf
		Next
	EndIf
EndFunc   ;==>_cLogin

Func _cLogout()
	Local $aItem, $sText, $iSelect
	$iSelect = _GUICtrlListView_GetSelectedIndices($Listview, True)
	If $iSelect[0] > 0 Then
		For $i = 1 To $iSelect[0]
			$aItem = _GUICtrlListView_GetItemTextArray($Listview, $iSelect[$i])
			If $aItem[6] <> "" Then _Logout($iSelect[$i], $aItem[6])
		Next
	EndIf
EndFunc   ;==>_cLogout

Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndListView, $tInfo
	$hWndListView = $Listview
	If Not IsHWnd($Listview) Then $hWndListView = GUICtrlGetHandle($Listview)
	$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
	$iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $hWndListView
			Switch $iCode
				Case 0 - 3 ; Sent by a list-view control when the user double-clicks an item with the left mouse button
					_cLogin()
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
			EndSwitch
	EndSwitch
	Return
EndFunc   ;==>WM_NOTIFY

Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), 1) = 1
EndFunc   ;==>_IsChecked

Func _ExitBvS()
	While ProcessExists("BvSsh.exe")
		ProcessClose("BvSsh.exe")
	WEnd
EndFunc   ;==>_ExitBvS

Func _Exit()
	_CheckRunning()
	Exit
EndFunc   ;==>_Exit

Func _CheckRunning()
	If ProcessExists("BvSsh.exe") Then
		If $AutoConnectAll = True Then
			_ExitBvS()
		Else
			Local $yes = MsgBox(4, "SSH Manager", "BitviseSSH is running." & @CRLF & "Do you want to exit all BitviseSSH?", 0, $Main)
			If $yes = 6 Then
				_ExitBvS()
			Else
				MsgBox(48, "SSH Manager", "Some BitviseSSH is running." & @CRLF & "Maybe error when login SSH.", 0, $Main)
			EndIf
		EndIf
	EndIf
EndFunc   ;==>_CheckRunning

Func Drag()
	IniWrite(@ScriptDir & "\Options.ini", "Options", "SSHPath", @GUI_DragFile)
	_ListView()
EndFunc   ;==>Drag

Func _HideBitvise()
	If _IsChecked($hCheckbox) Then
		$iHide = True
		IniWrite(@ScriptDir & "\Options.ini", "Options", "HideBitvise", 1)
	Else
		$iHide = False
		IniWrite(@ScriptDir & "\Options.ini", "Options", "HideBitvise", 0)
	EndIf
EndFunc   ;==>_HideBitvise

Func _StartPort()
	If _IsChecked($hCheckbox2) Then
		$iStartPort = True
		IniWrite(@ScriptDir & "\Options.ini", "Options", "StartPortEnable", 1)
		IniWrite(@ScriptDir & "\Options.ini", "Options", "StartPort", GUICtrlRead($hStartPort))
		GUICtrlSetState($hStartPort, 64)
		_ListView()
	Else
		$iStartPort = False
		IniWrite(@ScriptDir & "\Options.ini", "Options", "StartPortEnable", 0)
		GUICtrlSetState($hStartPort, 128)
		_ListView()
	EndIf
EndFunc   ;==>_StartPort

While 1
	Sleep(20)
WEnd
