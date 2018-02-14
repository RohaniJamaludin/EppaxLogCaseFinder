#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <Excel.au3>

; Create application object and create a new workbook
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

Local $oWorkbook = _Excel_BookNew($oExcel)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeWrite Example", "Error creating the new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
    _Excel_Close($oExcel)
    Exit
EndIf

_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "NO", "A2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "LOG KES", "B2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "NAMA MAJIKAN", "C2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "BILANGAN MESYUARAT", "D2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "KEPUTUSAN AKHIR MESYUARAT", "E2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "NO RUJUKAN", "F2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "STATUS TINDAKAN", "G2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "PEGAWAI", "H2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "PEJABAT", "I2")
_Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, "TARIKH TINDAKAN", "J2")

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

GUISetState(@SW_SHOW)


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

	  $oUser.value = GUICtrlRead($sUser)
	  $oPass.value = GUICtrlRead($sPass)
	  _IEAction($oSubmit, "click")

	  _IELoadWait($oIE)

	  $counter = 3
	  $numbering = 1

	  for $i=GUICtrlRead($from) to GUICtrlRead($to)

		 _IENavigate($oIE, "https://www.eppax.gov.my/eppax/sp/log-kes/" & $i)

		 $tags = $oIE.document.GetElementsByTagName("span")
		 Local $j = 0
		 $namaMajikan = ""
		 $bilanganMesyuarat = ""
		 $keputusanMesyuarat = ""

		 For $tag in $tags
			$class_value = $tag.className
			If ($class_value = "form-control-static" ) Then
			   $sInnerText = _IEPropertyGet($tag, "innertext")
			   If $j = 0 Then
				  $namaMajikan = $sInnerText
			   EndIf

			   If $j = 1 Then
				  $bilanganMesyuarat = $sInnerText
			   EndIf

			   If $j = 2 Then
				  $keputusanMesyuarat = $sInnerText
			   EndIf

			   $j = $j + 1
			EndIf
		 Next

		 _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $numbering , "A" & $counter)
		 _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet, $i , "B" & $counter)
		 _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $namaMajikan, "C" & $counter)
		 _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $bilanganMesyuarat, "D" & $counter)
		 _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $keputusanMesyuarat, "E" & $counter)

		 $tds = $oIE.document.GetElementsByTagName("td")

		 $found = True

		 Local $s = 0

		 For $td in $tds
			$found = False
			$sInnerText = _IEPropertyGet($td, "innertext")
			If $s = 0 Then
			   _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $sInnerText, "F" & $counter)
			EndIf

			If $s = 1 Then
			   _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $sInnerText, "G" & $counter)
			EndIf

			If $s = 2 Then
			   _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $sInnerText, "H" & $counter)
			EndIf

			If $s = 3 Then
			   _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $sInnerText, "I" & $counter)
			EndIf

			If $s = 4 Then
			   _Excel_RangeWrite($oWorkbook, $oWorkbook.Activesheet,  $sInnerText, "J" & $counter)
			EndIf

			If Mod($s, 4) = 0 And $s <> 0   Then
			   $s = 0
			   $counter = $counter + 1
			Else
			   $s = $s + 1
			EndIf
		 Next

		 If Not $found Then
			$counter = $counter + 1
		 EndIf

		 $numbering = $numbering + 1

		 Sleep(1000 * 2)
	  Next

   EndIf

EndSwitch

WEnd

EndFunc
