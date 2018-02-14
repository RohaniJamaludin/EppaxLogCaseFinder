#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>

Show()

Func catchIt()
   _IEQuit($oIE)
EndFunc

Func Show()

GUICreate("Scan Eppax Log Case Number", 400, 130, @DesktopWidth / 2 - 160, @DesktopHeight / 2 - 45, -1, $WS_EX_ACCEPTFILES)
GUICtrlCreateLabel("User Name", 16, 16, 100, 20)
Local $userName = GUICtrlCreateInput("", 130, 16, 250, 20)
GUICtrlCreateLabel("Password", 16, 40, 100, 20)
Local $password = GUICtrlCreateInput("", 130, 40, 250, 20, $ES_PASSWORD)
GUICtrlCreateLabel("Scan No From", 16, 64, 100, 20)
Local $from = GUICtrlCreateInput("", 130, 64, 100, 20)
GUICtrlCreateLabel("To", 240, 64, 30, 20)
Local $to= GUICtrlCreateInput("", 280, 64, 100, 20)
Local $btnRun = GUICtrlCreateButton("OK", 300, 100, 80, 26)

;GUICtrlSetState($btnStop, $GUI_DISABLE)

GUISetState(@SW_SHOW)

; Create file in same folder as script
$sFileName = @ScriptDir &"\Test.txt"

 ; Open file - deleting any existing content
$hFilehandle = FileOpen($sFileName, $FO_OVERWRITE)


While 1

If @error Then catchIt()

Switch GUIGetMsg()
   Case $GUI_EVENT_CLOSE
	  ExitLoop
      _IEQuit($oIE)

   Case $btnRun

	  If(GUICtrlRead($userName) = "" Or GUICtrlRead($password) = "" Or GUICtrlRead($from) = "" Or GUICtrlRead($to) = "") Then
		 MsgBox(0, "Error", "Please fill in all the field!")
	  Else
		  ; Prove it exists
	  If FileExists($sFileName) Then
	  ;MsgBox($MB_SYSTEMMODAL, "File", "Exists")
	  Else
	  ;MsgBox($MB_SYSTEMMODAL, "File", "Does not exist")
	  EndIf

	  Local $oUser, $oPass, $oSubmit
	  ;Local $sUser = "794886H81300"
	  Local $sUser = $userName
	  ;Local $sPass = "123456"
	  Local $sPass = $password
	  Local $url = "https://www.eppax.gov.my/eppax/login"
	  Local $oIE = _IECreate($url, 1)

	  _IELoadWait($oIE)

	  $oInputs = _IETagNameGetCollection($oIE, "input")
	  for $oInput in $oInputs
		 if $oInput.type = "text" And $oInput.id = "j_username" Then $oUser = $oInput
		 if $oInput.type = "password" And $oInput.id = "j_password" Then $oPass = $oInput
		 ;if $oInput.type = "submit" And $oInput.value = "submit" Then $oSubmit = $oInput
		 if isObj($oUser) And isObj($oPass) then exitloop
	  Next

	  $oButtons = _IETagNameGetCollection($oIE, "button")
	  for $oButton in $oButtons
		 if $oButton.type = "submit" And $oButton.value = "submit" Then $oSubmit = $oButton
		 if isObj($oSubmit) then exitloop
		 Next

	  ;MsgBox(0, "Inner Text", "Username: " & GUICtrlRead($sUser))

	  $oUser.value = GUICtrlRead($sUser)
	  $oPass.value = GUICtrlRead($sPass)
	  _IEAction($oSubmit, "click")

	  _IELoadWait($oIE)

	  ;_IENavigate($oIE, "https://www.eppax.gov.my/eppax/sp/log-kes/53004")

	  ;for $i=61694 to 63000
	  for $i=GUICtrlRead($from) to GUICtrlRead($to)

		 _IENavigate($oIE, "https://www.eppax.gov.my/eppax/sp/log-kes/" & $i)

		 FileWrite($hFilehandle, @CRLF &  @CRLF & @CRLF &  "Running Number : " & $i)

		 $tags = $oIE.document.GetElementsByTagName("span")
   Local $j = 0
   For $tag in $tags
	  $class_value = $tag.className
	  If ($class_value = "form-control-static" ) Then
		 $sInnerText = _IEPropertyGet($tag, "innertext")
		 If $j = 0 Then
			$sInnerText = "Nama Majikan : " &  $sInnerText
		 EndIf

		 If $j = 1 Then
			$sInnerText = "Bilangan Mesyuarat : " &  $sInnerText
		 EndIf

		 If $j = 2 Then
			$sInnerText = "Keputusan Akhir Selepas Mesyuarat : " &  $sInnerText
		 EndIf

        ;MsgBox(0, "Class: ", $sInnerText)
		FileWrite($hFilehandle, @CRLF & $sInnerText)
		$j = $j + 1
	  EndIf
   Next

   FileWrite($hFilehandle, @CRLF )

   $tds = $oIE.document.GetElementsByTagName("td")
   Local $s = 0
   For $td in $tds
	  $sInnerText = _IEPropertyGet($td, "innertext")
		 If $s = 0 Then
			$sInnerText = "No Rujukan : " &  $sInnerText
		 EndIf

		 If $s = 1 Then
			$sInnerText = "Status Tindakan : " &  $sInnerText
		 EndIf

		 If $s = 2 Then
			$sInnerText = "Pegawai : " &  $sInnerText
		 EndIf

		 If $s = 3 Then
			$sInnerText = "Pejabat : " &  $sInnerText
		 EndIf

		 If $s = 4 Then
			$sInnerText = "Tarikh Tindakan : " &  $sInnerText
		 EndIf


        ;MsgBox(0, "Class: ", $sInnerText)
		FileWrite($hFilehandle, @CRLF & $sInnerText)

		 If Mod($s, 4) = 0 And $s <> 0   Then
			FileWrite($hFilehandle, @CRLF )
			$s = 0
		 Else
			$s = $s + 1
		 EndIf


   Next


	  #cs------------------------
		 _IENavigate($oIE, "https://www.eppax.gov.my/eppax/sp/log-kes/" & $i)

		 FileWrite($hFilehandle, @CRLF & "Running Number : " & $i)

		 $oDivs = _IETagNameGetCollection($oIE, "span")

		 for $oDiv  in $oDivs
			$sInnerText = _IEPropertyGet($oDiv, "innertext")

			If Not @error Then
			   If Not($sInnerText = " Language " Or $sInnerText = "" Or $sInnerText = "JPC BUILDERS SDN. BHD." Or $sInnerText = "Dashboard" Or $sInnerText = " Permohonan Baharu" Or  $sInnerText = " Permohonan Rayuan" Or $sInnerText = " Semakan Status Permohonan Pekerja Asing" Or $sInnerText = "Kemaskini Majikan dan Pekerja Asing" Or $sInnerText = " Semakan Maklumat Pekerja Asing" Or $sInnerText = " Kemaskini Maklumat Permohonan Pekerja Asing" Or  $sInnerText = " Cetakan Dokumen-dokumen permohonan" Or $sInnerText = "Kemaskini Maklumat VDR dan Bayaran Levi" Or $sInnerText = " Semakan Status Levi" Or $sInnerText = " Permohonan VDR" Or $sInnerText = " Semakan Status VDR " Or $sInnerText = " Semakan Status Visa" Or $sInnerText = "Permohonan, Pembaharuan dan Pengemaskinian Permit" Or $sInnerText = " Semakan Maklumat FOMEMA" Or $sInnerText = " Mohon Penggantian Pekerja Asing" Or $sInnerText = " Semakan Memo Periksa Keluar (COM)" Or $sInnerText = " Permohonan Pekerja Asing" Or $sInnerText = " Kemaskini Maklumat Pekerja Asing" Or $sInnerText = "Log Kes" Or $sInnerText = "Maklumat Asas" Or $sInnerText = "Hakcipta © 2016 Jabatan Tenaga Kerja Semenanjung Malaysia" Or $sInnerText = "/eppax/" Or $sInnerText = "794886H81300" Or $sInnerText = "M" Or $sInnerText = "ms" Or $sInnerText = "Sila Tunggu..." Or $sInnerText = "Data telah ada" Or $sInnerText = "Hapus" Or $sInnerText = "Permohonan Pekerja Asing" Or $sInnerText = "Laporan" Or $sInnerText = " Laporan Adhoc Permohonan Pekerja Asing Majikan" Or $sInnerText = " Statistik Permohonan PA Sektor & Subsektor") Then

				  ;MsgBox(0, "Inner Text", "The innertext is: " & $sInnerText)
				  ; Write a line
				  ;FileWrite($hFilehandle, $sInnerText)

				  ; Read it
				  ;MsgBox($MB_SYSTEMMODAL, "File Content", FileRead($sFileName))

				  ;Append a line
				  FileWrite($hFilehandle, @CRLF & $sInnerText)
			   EndIf

			Else
			   MsgBox(0, "Error", "There was an error retrieving the innertext. The error number is: " & @error)
			EndIf
		 Next
 #ce-----------------------------------------
		 Sleep(1000 * 1)
   Next

   ; Close the file handle
   FileClose($hFilehandle)

   EndIf




 EndSwitch
 WEnd

EndFunc
