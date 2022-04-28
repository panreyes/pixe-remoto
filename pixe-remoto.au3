;============================================================================================================
; Date:             2022-04-28, 11:23
; Description:		Instant support executable. Launches UltraVNC and reports some data back to an URL.
; Author(s):        Pablo Navarro (@panreyes)
;============================================================================================================

#NoTrayIcon
Opt('MustDeclareVars', 1) ;Safer, always.

#include <ButtonConstants.au3>
#include <InetConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPISys.au3>
#include <WinAPIDiag.au3>
#include <WinAPIEx.au3>

; API URLs
Global $ip_resolver_api = "" 				;Should report back an IP or Domain, and maybe the port.
Global $helpdesk_service_api = ""  			;Where to send data related to the remote session. Not required.
Global $url_check_https = ""				;Should just report back OK.

; Texts
Const $txt_default_helpdesk_name = "Asistencia remota"
Const $txt_connection_prompt = "¿Deseas iniciar una asistencia remota?"
Const $txt_admin_prompt = "Se recomienda ejecutar este programa como un usuario Administrador para facilitar las operaciones del técnico." & @CRLF & @CRLF & "¿Deseas ejecutar el programa en modo administrador?"
Const $txt_starting_remote_assistance = "Iniciando asistencia remota..."
Const $txt_ending_remote_assistante = "Finalizando asistencia remota..."
Const $txt_reconnecting = "Reconectando..."
Const $txt_rejected_connection = "Conexión rechazada. El programa terminará."
Const $txt_prompt_connection_string = "Por favor, introduce la IP que te indique el técnico o pulsa Cancelar para finalizar este programa."
Const $txt_error_connecting = "No se pudo realizar una conexión con el servidor."
Const $txt_http_not_permitted = "No se pudo establecer una conexión https segura. El programa terminará."
Const $txt_reconnect = "Reconectar"
Const $txt_disconnect = "Desconectar"

; A few globals
Const $version = "1.0.22.4"
Const $debug = False
Global $helpdesk_name = $txt_default_helpdesk_name
Global $connection_strings_available[20][3] ;1: Technician name, 2: Connection string
Global $connection_strings_number
Global $connection_string
Global $connection_name
Global $service_mode = True
Global $connection_accepted = False
Global $datetime_start
Global $datetime_end
Global $hTimer
Global $do_report = True
Global $http_method = ""
Global $permit_http_without_ssl = True

;Main
check_system()
close_other_instances()
find_out_ip()
OnAutoItExitRegister("close_and_exit")

If (MsgBox(4, $helpdesk_name, $txt_connection_prompt) = 6) Then
	$connection_accepted = True
	$hTimer = TimerInit()
	$datetime_start = timestamp()
	If (report_helpdesk_service("start")) Then
		copy_files()
		launch_uvnc_and_gui()
	Else
		MsgBox(16, $helpdesk_name, $txt_rejected_connection)
	EndIf
EndIf
close_and_exit()

;-------------------

Func report_helpdesk_service($message)
	Local $hwid, $timerdiff, $query, $response, $domain

	If (Not $do_report OR $helpdesk_service_api == "") Then
		Return True
	EndIf

	;Get an unique hardware ID
	$hwid = _WinAPI_UniqueHardwareID()

	;Find out how much seconds has passed since we started the support session
	$timerdiff = Floor(TimerDiff($hTimer) / 1000)
	$datetime_end = timestamp()

	If (@LogonDomain <> @ComputerName) Then
		$domain = @LogonDomain
	Else
		$domain = ""
	EndIf

	$query = $helpdesk_service_api & _
			"&version=" & _URIEncode($version) & _
			"&helpdesk_name=" & _URIEncode($helpdesk_name) & _
			"&connection_name=" & _URIEncode($connection_name) & _
			"&lan_ip=" & _URIEncode(@IPAddress1) & _
			"&connection_string=" & _URIEncode($connection_string) & _
			"&message=" & _URIEncode($message) & _
			"&hwid=" & _URIEncode($hwid) & _
			"&dt_start=" & _URIEncode($datetime_start) & _
			"&dt_end=" & _URIEncode($datetime_end) & _
			"&duration=" & _URIEncode($timerdiff) & _
			"&pcname=" & _URIEncode(@ComputerName) & _
			"&username=" & _URIEncode(@UserName) & _
			"&language=" & _URIEncode(@MUILang) & _
			"&domain=" & _URIEncode($domain) & _
			"&workgroup=" & _URIEncode(get_workgroup()) & _
			"&random=" & Random(0, 10000)

	DebugMsgBox($query)

	;Reportamos el servicio
	$response = InetReadText($query)

	DebugMsgBox($response)

	If (StringLeft($response, 3) = "OK|") Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>report_helpdesk_service

Func check_system()
	;Check if it is the program is being run in a RDP session
	If (_WinAPI_GetSystemMetrics($SM_REMOTESESSION) = 1) Then
		$service_mode = False
	EndIf

	;Check if it is being run as admin, and ask the user if they want to elevate it
	If (Not IsAdmin()) Then
		If (MsgBox(4, $helpdesk_name, $txt_admin_prompt) = 6) Then
			_GetAdminRight()
		Else
			$service_mode = False
		EndIf
	EndIf

	;Check if we support https and API urls
	check_https()
	$ip_resolver_api = StringReplace($ip_resolver_api, "http://", $http_method&"://")
	$helpdesk_service_api = StringReplace($helpdesk_service_api, "http://", $http_method&"://")
EndFunc   ;==>check_system

Func find_out_ip()
	Local $response

	$helpdesk_name = StringLeft(@ScriptName, StringLen(@ScriptName) - 4)

	;Debug
	;$helpdesk_name = ""

	;Clean up number related to the file being downloaded multiple times
	For $i = 1 To 10
		$helpdesk_name = StringReplace($helpdesk_name, " (" & $i & ")", "")
		$helpdesk_name = StringReplace($helpdesk_name, "(" & $i & ")", "") ;Hay veces que no hay espacio...
	Next

	If (Not $connection_string) Then
		If ($helpdesk_name <> "remoto.exe" And $helpdesk_name <> "pixe-remoto.exe" and $ip_resolver_api <> "") Then
			$response = InetReadText($ip_resolver_api & "&n=" & $helpdesk_name)
			$response = StringReplace($response, @CR, "")
			If (StringInStr($response, @LF)) Then
				;Multiple connection strings are available. Process them:
				Local $temp = StringSplit($response, @LF)
				For $i = 1 To $temp[0]
					Local $temp2 = StringSplit($temp[$i], "|")
					$connection_strings_number = $i
					$connection_strings_available[$i][1] = $temp2[1]
					$connection_strings_available[$i][2] = $temp2[2]
				Next

				;TODO: Menu to select between multiple connection strings

				;By now we will first connect to the first available vncviewer
				$connection_name = $connection_strings_available[1][1]
				$connection_string = $connection_strings_available[1][2]
			Else
				$connection_name = ""
				$connection_string = $response
				$connection_strings_available[1][1] = ""
				$connection_strings_available[1][2] = $connection_string
				$connection_strings_number = 1
			EndIf
		EndIf

		DebugMsgBox($ip_resolver_api & "&n=" & $helpdesk_name)
		DebugMsgBox($connection_string)

		If (Not $connection_string) Then
			$do_report = False
			$connection_string = InputBox($helpdesk_name, $txt_prompt_connection_string)
			If ($connection_string = "") Then
				MsgBox(64, "Error", $txt_error_connecting)
				close_and_exit()
			Else
				If (Not StringInStr($connection_string, ":")) Then
					$connection_string = $connection_string & "::5510"
				EndIf

				$connection_strings_available[1][1] = ""
				$connection_strings_available[1][2] = $connection_string
				$connection_strings_number = 1
				$connection_name = ""
			EndIf
		EndIf
	EndIf
EndFunc   ;==>find_out_ip

Func copy_files()
	If (@OSVersion = "WIN_XP") Then
		FileInstall("include\remoto\winvnc-xp.exe", @TempDir & "\winvnc.exe", 1)
	Else
		FileInstall("include\remoto\winvnc.exe", @TempDir & "\winvnc.exe", 1)
		FileInstall("include\remoto\vnchooks.dll", @TempDir & "\vnchooks.dll", 1)
	EndIf
	FileInstall("include\remoto\SecureVNCPlugin.dsm", @TempDir & "\SecureVNCPlugin.dsm", 1)
	FileInstall("include\remoto\ultravnc.ini", @TempDir & "\ultravnc.ini", 1)
	FileInstall("include\remoto\options.vnc", @TempDir & "\options.vnc", 1)
	FileInstall("include\remoto\close.gif", @TempDir & "\close.gif", 1)
	FileInstall("include\remoto\refresh.gif", @TempDir & "\refresh.gif", 1)
EndFunc   ;==>copy_files

Func launch_uvnc_and_gui()
	Local $btn_reconnect[10]
	Local $btn_cancel
	Local $nMsg

	While 1
		status($txt_starting_remote_assistance)

		;Clean up existing VNCs, just in case
		close_vnc()

		;Debug
		;IniWrite(@TempDir & "\ultravnc.ini", "admin", "Path", @TempDir)
		;IniWrite(@TempDir & "\ultravnc.ini", "admin", "DebugMode", 2)
		;IniWrite(@TempDir & "\ultravnc.ini", "admin", "DebugLevel", 12)

		;No debug
		IniWrite(@TempDir & "\ultravnc.ini", "admin", "DebugMode", 0)
		IniWrite(@TempDir & "\ultravnc.ini", "admin", "DebugLevel", 0)

		;As a service:
		If ($service_mode = True) Then
			IniWrite(@TempDir & "\ultravnc.ini", "admin", "service_commandline", '-autoreconnect -connect ' & $connection_string)
			RunWait(@TempDir & "\winvnc.exe -install", @TempDir)
		Else
			;User mode (no CAD or UAC for you!):
			Run(@TempDir & "\winvnc.exe -autoreconnect -connect " & $connection_string & " -run", @TempDir)
		EndIf
		Sleep(5000)
		GUIDelete()

		Local $action = 0

		;Little window with 2 buttons: Reconnect and Disconnect
		GUICreate($helpdesk_name, 93 + (82 * $connection_strings_number), 115, @DesktopWidth - 180 - 82, @DesktopHeight - 145, $WS_EX_TOOLWINDOW, $WS_EX_TOPMOST)
		GUISetBkColor(0x0078D7)

		For $i = 1 To $connection_strings_number
			GUICtrlCreateLabel($txt_reconnect & @CRLF & $connection_strings_available[$i][1], 5 + (82 * ($i - 1)), 8, 77, 30, $SS_CENTER)
			GUICtrlSetFont(-1, 8, 800, 0, "Verdana")
			GUICtrlSetColor(-1, 0xFFFFFF)
			$btn_reconnect[$i] = GUICtrlCreatePic(@TempDir & "\refresh.gif", 18 + (82 * ($i - 1)), 35, 45, 45)
		Next

		GUICtrlCreateLabel($txt_disconnect, 5 + (82 * $connection_strings_number), 8, 85, 17)
		GUICtrlSetFont(-1, 8, 800, 0, "Verdana")
		GUICtrlSetColor(-1, 0xFFFFFF)
		$btn_cancel = GUICtrlCreatePic(@TempDir & "\close.gif", 18 + (82 * $connection_strings_number), 35, 45, 45)

		GUISetState(@SW_SHOW)

		While 1
			$nMsg = GUIGetMsg()
			If ($nMsg <> 0) Then
				If ($nMsg = $GUI_EVENT_CLOSE Or $nMsg = $btn_cancel) Then
					$action = 1
				Else
					For $i = 1 To $connection_strings_number
						If ($nMsg = $btn_reconnect[$i]) Then
							$connection_name = $connection_strings_available[$i][1]
							$connection_string = $connection_strings_available[$i][2]
							$action = 2
						EndIf
					Next
				EndIf
			EndIf

			If ($action <> 0) Then
				ExitLoop
			EndIf
		WEnd

		GUIDelete()

		If ($action = 1) Then ;finalizamos
			status($txt_ending_remote_assistante)
			ExitLoop
		Else ;reconectamos
			status($txt_reconnecting)
			;close_vnc()
		EndIf
	WEnd
EndFunc   ;==>launch_uvnc_and_gui

Func close_vnc()
	Local $i

	;Fast check. We may want to close a latent VNC session from another run of this program that was not closed correctly.
	If (Not ProcessExists("winvnc.exe")) Then
		Return
	EndIf

	;Close current connection and delete the service if it exists
	If ($service_mode = True) Then
		RunWait(@TempDir & "\winvnc.exe -stopservice", @TempDir)
		Sleep(1000)
		Run("cmd /c sc delete uvnc_service", @TempDir, @SW_HIDE)
		If (ProcessExists("winvnc.exe")) Then
			ProcessClose("winvnc.exe")
			ProcessClose("winvnc.exe")
			ProcessClose("winvnc.exe")
		EndIf
	Else
		ProcessClose("winvnc.exe")
	EndIf

	;Wait up to 5 seconds if winvnc still is running
	For $i = 1 To 5
		If (Not ProcessExists("winvnc.exe")) Then
			ExitLoop
		EndIf
		Sleep(1000)
	Next
EndFunc   ;==>close_vnc

Func close_and_exit()
	If ($connection_accepted) Then
		report_helpdesk_service("end")
		$connection_accepted = False
	EndIf

	close_vnc()

	;Delete temporary files
	FileDelete(@TempDir & "\winvnc.exe")
	FileDelete(@TempDir & "\ultravnc.ini")
	FileDelete(@TempDir & "\vnchooks.dll")
	FileDelete(@TempDir & "\options.vnc")
	FileDelete(@TempDir & "\refresh.gif")
	FileDelete(@TempDir & "\close.gif")
	FileMove(@TempDir & "\winvnc.exe", @TempDir & "\winvnc.tmp", 1) ;Just in case if it did not get deleted

	Exit
EndFunc   ;==>close_and_exit

Func status($message)
	GUIDelete()
	GUICreate($helpdesk_name, 300, 50)
	GUICtrlCreateLabel($message, 15, 13, 280, 100)
	GUISetState()
EndFunc   ;==>status

Func close_other_instances()
	Local $pid_list = ProcessList(@ScriptName)
	If (Not @error) Then
		If ($pid_list[0][0] > 1) Then
			For $i = 1 To $pid_list[0][0]
				If ($pid_list[$i][1] <> @AutoItPID) Then
					ProcessClose($pid_list[$i][1])
				EndIf
			Next
		EndIf
	EndIf
EndFunc   ;==>close_other_instances

Func _GetAdminRight($sCmdLineRaw = "")
	If Not IsAdmin() Then
		If Not $sCmdLineRaw Then $sCmdLineRaw = $CmdLineRaw
		ShellExecute(@AutoItExe, $sCmdLineRaw, "", "runas")
		ProcessClose(@AutoItPID)
		Exit
	EndIf
EndFunc   ;==>_GetAdminRight

Func get_workgroup()
	Local $colItems = wmi_query("CIMV2", "SELECT Workgroup FROM Win32_ComputerSystem")

	If IsObj($colItems) Then
		For $objItem In $colItems
			Return $objItem.Workgroup
		Next
	EndIf

	Return ""
EndFunc   ;==>get_workgroup

Func wmi_query($path, $query)
	Local $wbemFlagReturnImmediately = 0x10
	Local $wbemFlagForwardOnly = 0x20
	Local $objWMIService = ObjGet("winmgmts:\\.\root\" & $path)
	If (Not IsObj($objWMIService)) Then
		Return ""
	EndIf

	Local $colItems = $objWMIService.ExecQuery($query, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If (Not IsObj($colItems)) Then
		Return ""
	EndIf

	Return $colItems
EndFunc   ;==>wmi_query

Func InetReadText($query)
	Return BinaryToString(InetRead($query, $INET_FORCERELOAD))
	;Return HttpPost($query) ;Deprecated. Does not support https
EndFunc   ;==>InetReadText

Func timestamp()
	Return @YEAR & "-" & @MON & "-" & @MDAY & "T" & @HOUR & ":" & @MIN & ":" & @SEC
EndFunc   ;==>timestamp

Func DebugMsgBox($mensaje)
	If ($debug) Then
		MsgBox(0, "debug", $mensaje)
	EndIf
EndFunc   ;==>DebugMsgBox

Func check_https()
	If($url_check_https == "") Then
		$http_method = "http"
		Return
	EndIf

	If ($permit_http_without_ssl) Then
		;These old OS don't support current SSL protocols:
		If (@OSVersion = "WIN_7" Or @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_XP") Then
			$http_method = "http"
			Return
		EndIf
	EndIf

	Local $resultado = InetReadText($url_check_https)
	If ($resultado <> "OK") Then
		$http_method = "http"
	Else
		$http_method = "https"
	EndIf

	If ($http_method == "http" And $permit_http_without_ssl == False) Then
		MsgBox(16, $helpdesk_name, $txt_http_not_permitted)
		close_and_exit()
	EndIf
EndFunc   ;==>check_https

Func _URIEncode($sData)
    Local $aData = StringSplit(BinaryToString(StringToBinary($sData,4),1),"")
    Local $nChar
    $sData=""
    For $i = 1 To $aData[0]
        $nChar = Asc($aData[$i])
        Switch $nChar
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
                $sData &= $aData[$i]
            Case 32
                $sData &= "+"
            Case Else
                $sData &= "%" & Hex($nChar,2)
        EndSwitch
    Next
    Return $sData
EndFunc