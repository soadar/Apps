#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.0
 Author:         Soadar

 Script Function:
	Perfil backup.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <File.au3>
#include <Date.au3>
#include <Array.au3>
#include <GUIConstants.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <ScreenCapture.au3>

$hGUI = GUICreate ("Perfil Backup", 240,320)
GUICtrlCreateLabel( "Backup personalizado", 10, 15, 200, 20)
GUICtrlSetFont(-1, 12, 100, 0, "Arial")
$checkDrives = GUICtrlCreateCheckbox("Backup Map Drives", 10, 50, 185, 25)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")
$checkChrome = GUICtrlCreateCheckbox("Backup Chrome/Brave", 10, 80, 180, 25)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")
$checkFF = GUICtrlCreateCheckbox("Backup Firefox", 10, 110, 120, 25)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")
$checkPrinters = GUICtrlCreateCheckbox("Backup Printers", 10, 140, 120, 25)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")
$checkNetwork = GUICtrlCreateCheckbox("Backup Network info", 10, 170, 185, 25)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")
$checkPrograms = GUICtrlCreateCheckbox("Backup Programs Info", 10, 200, 185, 25)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")
$checkAll = GUICtrlCreateCheckbox("Seleccionar todo", 10, 230, 185, 25)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")
$btnOk = GUICtrlCreateButton ( "Guardar", 20, 265, 80, 40)
GUICtrlSetFont(-1, 11, 100, 0, "Arial")

$btnCaptura = GUICtrlCreateButton ( "Capturar Imagenes", 140, 265, 80, 40, $BS_MULTILINE)

;GUISetState()
GUISetState(@SW_SHOW, $hGUI)
global $bkpFolder

$user = @username

While 1
   Switch GUIGetMsg()
      Case $GUI_EVENT_CLOSE
		 Exit
	  Case $checkAll
		 If GUICtrlRead($checkAll) = $GUI_CHECKED Then
			_checkBoxState($GUI_CHECKED)
		Else
			_checkBoxState($GUI_UNCHECKED)
		EndIf
	  Case $btnOk
		if not DirCreate(@ScriptDir & '\Backup_' & @username) Then 
			MsgBox($MB_SYSTEMMODAL, "", 'No se pudo generar la carpeta de backup' & @CRLF & 'Los archivos se crearan en la ruta del ejecutable')
			$bkpFolder = @ScriptDir & '\'
		Else
			$bkpFolder = @ScriptDir & '\Backup_' & @username & '\'
		EndIf
		if GUICtrlRead($checkDrives) = $GUI_CHECKED Then
			_MapDrives()
		EndIf
		if GUICtrlRead($checkChrome) = $GUI_CHECKED Then
			_Google()
			_Brave()
		EndIf
		if GUICtrlRead($checkFF) = $GUI_CHECKED Then
			_Firefox()
		EndIf
		if GUICtrlRead($checkPrinters) = $GUI_CHECKED Then
			_Printers()
		EndIf
		if GUICtrlRead($checkNetwork) = $GUI_CHECKED Then
			_NetworkInfo()
		EndIf
		if GUICtrlRead($checkPrograms) = $GUI_CHECKED Then
			_SoftwareInfo()
		EndIf
	  Case $btnCaptura
			_Capture()
	EndSwitch
Wend

Func _MapDrives()
	local $hFile = FileOpen($bkpFolder & 'MapDrives.bat', $FO_OVERWRITE)
	If $hFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, "Atención", "Error al generar MapDrives.bat en " & $bkpFolder )
	Else
		FileWriteLine($hFile, "@echo off")
		local $network_drives = DriveGetDrive('NETWORK')
		If @error Then
			FileWriteLine($hFile, "echo No se encontraron unidades de red" & @CRLF & "pause>nul")
		Else
			For $1 = 1 To UBound($network_drives) -1
				$network_path = DriveMapGet($network_drives[$1])
				If @error Then ContinueLoop
				$command = 'net use ' & $network_drives[$1] & ' "' & $network_path & '"' & ' /persistent:yes'
				FileWriteLine($hFile, $command)
			Next
		EndIf
		FileClose($hFile)
	EndIf
EndFunc

Func _NetworkInfo()
	local $ip1 = @IPAddress1, $ip2 = @IPAddress2, $ip3 = @IPAddress3, $ip4 = @IPAddress4
	local $cont = 1, $ip = @IPAddress1, $linea = '--------------------------------------------------'
	Local $hFile = FileOpen($bkpFolder & "Network_Info.log", 1)
	$string = ''
	$string &= $linea & @CRLF & _Now() & @CRLF
	while $ip <> '0.0.0.0'
		$MAC_Address = _GetMac($ip)
		If ($ip & $cont <> "") and ($ip & $cont <> "0.0.0.0") Then
			$string &= "IP " & $cont & ": " & $ip & " - MAC " & $cont & ": " & $MAC_Address & @CRLF
		endif
		$cont = $cont + 1
		$ip = Eval('ip' & $cont)
	WEnd
	FileWrite($hFile, $string & $linea & @CRLF)
EndFunc

Func _GetMac($_MACsIP)
    Local $_MAC,$_MACSize
    Local $_MACi,$_MACs,$_MACr,$_MACiIP
    $_MAC = DllStructCreate("byte[6]")
    $_MACSize = DllStructCreate("int")
    DllStructSetData($_MACSize,1,6)
    $_MACr = DllCall ("Ws2_32.dll", "int", "inet_addr", "str", $_MACsIP)
    $_MACiIP = $_MACr[0]
    $_MACr = DllCall ("iphlpapi.dll", "int", "SendARP", "int", $_MACiIP, "int", 0, "ptr", DllStructGetPtr($_MAC), "ptr", DllStructGetPtr($_MACSize))
    $_MACs  = ""
    For $_MACi = 0 To 5
    If $_MACi Then $_MACs = $_MACs & ":"
        $_MACs = $_MACs & Hex(DllStructGetData($_MAC,1,$_MACi+1),2)
    Next
    DllClose($_MAC)
    DllClose($_MACSize)
    Return $_MACs		
EndFunc



Func _Firefox()
	local $profiles = _ListProfile()
	local $aLines, $flagChange = 0
	;;;;;;;;;;;;;;;;;;;;;;
	$folderFF = $bkpFolder
	if DirCreate($folderFF & 'Firefox') Then 
		$folderFF &= 'Firefox\'
	EndIf
	;;;;;;;;;;;;;;;;;;;;;;
	For $j = 1 To $profiles[0]
		$ruteProfiles = @AppDataDir & '\Mozilla\Firefox\Profiles\' & $profiles[$j] & '\places.sqlite'
		if FileExists($ruteProfiles) Then
			$ultimoCambio = FileGetTime($ruteProfiles)
			$ultimoCambio = "-" & $ultimoCambio[2] & "-" & $ultimoCambio[1] & "-" & $ultimoCambio[0]
			$rutaFull = $profiles[$j] & $ultimoCambio
			if DirCreate($folderFF & $rutaFull) Then $folderFF &= $rutaFull 
			
			FileCopy($ruteProfiles, $folderFF, 1)
			$folderFF = $bkpFolder & 'Firefox\'
		EndIf
	Next
EndFunc

Func _ListProfile()
        Local $aFileList = _FileListToArray(@AppDataDir & '\Mozilla\Firefox\Profiles', "*")
        If @error = 1 Then
                MsgBox($MB_SYSTEMMODAL, "Atención", "Error en la ruta de los perfiles de FF.")
                Exit
        EndIf
        If @error = 4 Then
                MsgBox($MB_SYSTEMMODAL, "Atención", "No se encontraron perfiles activos de FF")
                Exit
        EndIf
		;_ArrayDelete($aFileList, 0)
		return $aFileList
EndFunc

Func _Printers()
	local $cont = 1, $flag = 0
	local $file = $bkpFolder & 'Impresoras.bat'
	local $hFile = FileOpen($file, $FO_OVERWRITE)
	If $hFile = -1 Then
		MsgBox($MB_SYSTEMMODAL, "Atención", "Error al generar Impresoras.bat en " & $bkpFolder )
	Else
		FileWriteLine($hFile, "@echo off")
		local $printerkey = RegEnumKey("HKEY_CURRENT_USER\Printers\Connections", $cont)
		While StringLen($printerkey) > 0
			$flag = 1
			$printerkey = StringReplace($printerkey, ",", "\")
			$printerkey = StringReplace($printerkey, "sfs-1", "sprt0001")
			$printerkey = StringReplace($printerkey, "sprt0002", "sprt0001")
			FileWriteLine($hFile, "explorer.exe " & $printerkey)
			$cont = $cont + 1
			$printerkey = RegEnumKey("HKEY_CURRENT_USER\Printers\Connections", $cont)
		WEnd
		if $flag = 0 Then
			FileWriteLine($hFile, "echo No se encontraron impresoras" & @CRLF & "pause>nul")
		EndIf
		FileClose($hFile)
	EndIf
EndFunc

Func _SoftwareInfo()
	$file = FileOpen($bkpFolder & 'Programas.log', $FO_OVERWRITE)
	If $file = -1 Then
		FileWrite($file, "Error obteniendo datos")
	Else
		Local $count = 1
		Local Const $regkey = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		$key = RegEnumKey ($regkey, $count)
		if $key <> '' Then
			While 1
				$key = RegEnumKey ($regkey, $count)
				If @error <> 0 then ExitLoop
				$line = RegRead ($regkey & '\' & $key, 'Displayname')
				$line = StringReplace ($line, ' (remove only)', '')

				If $line <> '' And StringInStr($line, "C++") <= 0 And StringInStr($line, "Update") <= 0 And StringInStr($line, "Actualizaci") <= 0 And StringInStr($line, "Proof") <= 0 And StringInStr($line, "Service Pack") <= 0 Then
					If Not IsDeclared('avArray') Then Dim $avArray[1]
					ReDim $avArray[UBound($avArray) + 1]
					$avArray[UBound($avArray) - 1] = $line          
				EndIf
				$count = $count + 1
			WEnd
			If Not IsDeclared('avArray') Or Not IsArray($avArray) Then
				FileWrite($file, "Error obteniendo datos")
			EndIf
		EndIf
		_FileWriteFromArray($file, $avArray, 1)
	EndIf
EndFunc

Func _Google()
	$folderChr = $bkpFolder
	if DirCreate($folderChr & 'Chrome') Then 
		$folderChr &= 'Chrome\'
	EndIf

	$googlePath = @LocalAppDataDir & '\google\chrome\User Data\Default\'
	copyFilesNaveg($googlePath, $folderChr)
EndFunc

Func _Brave()
	$folderBrave = $bkpFolder
	if DirCreate($folderBrave & 'Brave') Then 
		$folderBrave &= 'Brave\'
	EndIf
	$bravePath = @LocalAppDataDir & '\BraveSoftware\Brave-Browser\User Data\Default\'
	copyFilesNaveg($bravePath, $folderBrave)
EndFunc

Func _Capture()
	MsgBox($MB_SYSTEMMODAL, "Atención", 'No utilizar la maquina hasta que finalicen las capturas.')
	$folderImg = $bkpFolder
	if DirCreate($folderImg & 'Capturas') Then 
		$folderImg &= 'Capturas\'
	EndIf
	
	$devices = "control /name Microsoft.DeviceManager"
	$red = "ncpa.cpl"
	$impresoras = "control printers"
	$thisPC = "explorer file:\\"
	;$capturaRute = "explorer " & $folderImg
	
	_CaptureRun($folderImg & "Devices.jpg", $devices, "dispositivos")
	_CaptureShell($folderImg & "Red.jpg", $red, "Conexiones")
	_CaptureRun($folderImg & "Impresoras.jpg", $impresoras, "Dispositivos")
	_CaptureRun($folderImg & "Equipo.jpg", $thisPC, "equipo")

	WinMinimizeAll()
	Sleep(500)
	_ScreenCapture_Capture($folderImg & "Desktop.jpg")
	WinMinimizeAllUndo()
	WinActivate("[RegexpTitle:(?i)(.*perfil backup*)]")
	;run($capturaRute)
	;Sleep(400)
	MsgBox($MB_SYSTEMMODAL, "Atención", 'Finalizo la captura de imágenes.')
EndFunc

Func _CaptureShell($file, $comando, $handler)
	ShellExecute($comando)
	if WinWaitActive("[RegexpTitle:(?i)(.*" & $handler & "*)]","",10) == 0 Then exit
	Sleep(1000)
	$hWnd = WinGetHandle("[RegexpTitle:(?i)(.*" & $handler & "*)]")
	WinSetState ($hWnd, "", @SW_MAXIMIZE)
	_ScreenCapture_CaptureWnd($file, $hWnd)
	WinClose($hWnd)
EndFunc

Func _CaptureRun($file, $comando, $handler)
	run($comando)
	if WinWaitActive("[RegexpTitle:(?i)(.*" & $handler & "*)]","",10) == 0 Then exit
	Sleep(1000)
	$hWnd = WinGetHandle("[RegexpTitle:(?i)(.*" & $handler & "*)]")
	WinSetState ($hWnd, "", @SW_MAXIMIZE)
	_ScreenCapture_CaptureWnd($file, $hWnd)
	WinClose($hWnd)
EndFunc

Func _checkBoxState($estado)
	GUICtrlSetState($checkDrives, $estado)
	GUICtrlSetState($checkChrome, $estado)
	GUICtrlSetState($checkFF, $estado)
	GUICtrlSetState($checkPrinters, $estado)
	GUICtrlSetState($checkNetwork, $estado)
	GUICtrlSetState($checkPrograms, $estado)
EndFunc

Func copyFilesNaveg($path, $folder)
	FileCopy($path & 'Bookmarks', $folder, 1)
	FileCopy($path & 'Favicons', $folder, 1)
	FileCopy($path & 'History', $folder, 1)
	FileCopy($path & 'Preferences', $folder, 1)
	FileCopy($path & 'Web Data', $folder, 1)
	FileCopy($path & 'Login Data', $folder, 1)
EndFunc