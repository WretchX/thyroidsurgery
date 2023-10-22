#Region UDF
; #FUNCTION# ====================================================================================================================
; Name...........: _StartTTS
; Description ...: Creates a object to be used with Text-to-Speak Functions.
; Syntax.........: _StartTTS()
; Parameters ....:
; Return values .: Success - Returns a object
; Author ........: bchris01

; Example .......: Yes
; ===============================================================================================================================
Func _StartTTS()
    Return ObjCreate("SAPI.SpVoice")
EndFunc   ;==>_StartTTS

; #FUNCTION# ====================================================================================================================
; Name...........: _SetRate
; Description ...: Sets the rendering rate of the voice. (How fast the voice talks.)
; Syntax.........: _SetRate(ByRef $Object, $iRate)
; Parameters ....: $Object        - Object returned from _StartTTS().
;                  $iRate         - Value specifying the speaking rate of the voice. Supported values range from -10 to 10
; Return values .:  None
; Author ........: bchris01
; Example .......: Yes
; ===============================================================================================================================
Func _SetRate(ByRef $Object, $iRate); Rates can be from -10 to 10
    $Object.Rate = $iRate
EndFunc   ;==>_SetRate

; #FUNCTION# ====================================================================================================================
; Name...........: _SetVolume
; Description ...: Sets the volume of the voice.
; Syntax.........: _SetVolume(ByRef $Object, $iVolume)
; Parameters ....: $Object        - Object returned from _StartTTS().
;                  $iVolume       - Value specifying the volume of the voice. Supported values range from 0-100. Default = 100
; Return values .:  None
; Author ........: bchris01
; Example .......: Yes
; ===============================================================================================================================
Func _SetVolume(ByRef $Object, $iVolume);Volume
    $Object.Volume = $iVolume
EndFunc   ;==>_SetVolume

; #FUNCTION# ====================================================================================================================
; Name...........: _SetVoice
; Description ...: Sets the identity of the voice used for text synthesis.
; Syntax.........: _SetVoice(ByRef $Object, $sVoiceName)
; Parameters ....: $Object        - Object returned from _StartTTS().
;                  $sVoiceName    - String matching one of the voices installed.
; Return values .:  Success - Sets object to voice.
;                   Failure - Sets @error to 1
; Author ........: bchris01
; Example .......: Yes
; ===============================================================================================================================
Func _SetVoice(ByRef $Object, $sVoiceName)
    Local $VoiceNames, $VoiceGroup = $Object.GetVoices
    For $VoiceNames In $VoiceGroup
        If $VoiceNames.GetDescription() = $sVoiceName Then
            $Object.Voice = $VoiceNames
            Return
        EndIf
    Next
    Return SetError(1)
EndFunc   ;==>_SetVoice

; #FUNCTION# ====================================================================================================================
; Name...........: _GetVoices
; Description ...: Retrives the currently installed voice identitys.
; Syntax.........: _GetVoices(ByRef $Object[, $Return = True])
; Parameters ....: $Object        - Object returned from _StartTTS().
;                  $bReturn       - String of text you want spoken.
;                  |If $bReturn = True then a 0-based array is returned.
;                  |If $bReturn = False then a string seperated by delimiter "|" is returned.
; Return values .:  Success - Returns an array or string containing installed voice identitys.
; Author ........: bchris01
; Example .......: Yes
; ===============================================================================================================================
Func _GetVoices(ByRef $Object, $bReturn = True)
    Local $sVoices, $VoiceGroup = $Object.GetVoices
    For $Voices In $VoiceGroup
        $sVoices &= $Voices.GetDescription() & '|'
    Next
    If $bReturn Then Return StringSplit(StringTrimRight($sVoices, 1), '|', 2)
    Return StringTrimRight($sVoices, 1)
EndFunc   ;==>_GetVoices

; #FUNCTION# ====================================================================================================================
; Name...........: _Speak
; Description ...: Speaks the contents of the text string.
; Syntax.........: _Speak(ByRef $Object, $sText)
; Parameters ....: $Object        - Object returned from _StartTTS().
;                  $sText         - String of text you want spoken.
; Return values .:  Success - Speaks the text.
; Author ........: bchris01
; Example .......: Yes
; ===============================================================================================================================
Func _Speak(ByRef $Object, $sText)
    $Object.Speak($sText)
EndFunc   ;==>_Speak
#EndRegion

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <SliderConstants.au3>
#include <ColorConstants.au3>
#include <EditConstants.au3>

HotKeySet("^e", "ActiveWindow")
HotKeySet("{HOME}","Terminate")

$fontSize = 20
$winTitle = "Thyroid"

#Region GUI
$GUI_MAIN = GUICreate($winTitle, 555, 450, -1, -1)
$Edit1 = GUICtrlCreateEdit("", 24, 110, 505, 321, Bitor($ES_WANTRETURN, $WS_VSCROLL, $ES_AUTOVSCROLL))
GUICtrlSetFont(-1, $fontSize, 500)
$butAlert = GUICtrlCreateButton("Alert" & @CR & @CR & "(CTRL+Q)", 24, 12, 81, 65, $BS_MULTILINE)
GUICtrlSetFont(-1, 10, 500)
$butTTS = GUICtrlCreateButton("TTS" & @CR & @CR & "(CTRL+S)", 125, 12, 81, 65, $BS_MULTILINE)
GUICtrlSetFont(-1, 10, 500)
$butClear = GUICtrlCreateButton("Clear" & @CR & @CR & "(CTRL+A)", 228, 12, 81, 65, $BS_MULTILINE)
GUICtrlSetFont(-1, 10, 500)
$chkOnTop = GUICtrlCreateCheckbox("Stay On Top", 325, 8, 81, 25)
$radSound1 = GUICtrlCreateRadio("Sound 1", 330, 35, 57, 25)
GUICtrlSetState($radSound1, $GUI_CHECKED)
$radSound2 = GUICtrlCreateRadio("Sound 2", 330, 55, 57, 25)
$radSound3 = GUICtrlCreateRadio("Sound 3", 330, 75, 57, 25)
$butYes = GUICtrlCreateButton("Yes", 420, 12, 54, 40)
GUICtrlSetFont(-1, 10, 900)
$butNo = GUICtrlCreateButton("No", 420, 58, 54, 40)
GUICtrlSetFont(-1, 10, 900)
$butFontSizeUp = GUICtrlCreateButton("Size Up", 480, 12, 50, 40)
$butFontSizeDown = GUICtrlCreateButton("Size Dn", 480, 58, 50, 40)
$labSlider = GUICtrlCreateLabel("Volume:", 24, 85, 45, 16)
$Slider1 = GUICtrlCreateSlider(64, 80, 220, 24, BitOR($TBS_TOOLTIPS, $TBS_AUTOTICKS, $TBS_ENABLESELRANGE))
$inVolume = GUICtrlCreateInput("50", 284, 82, 25, 20, $ES_NUMBER)
GUICtrlSetLimit(-1, 100, 0) ; control, max, min
GUICtrlSetData($Slider1, 50)
$labMuted = GUICtrlCreateLabel("Muted", 240, 431, 80, 20)
GUICtrlSetFont($labMuted, "", 900)
GUICtrlSetColor($labMuted, $COLOR_RED)
GUICtrlSetState($labMuted, $GUI_HIDE)
$labActivate = GUICtrlCreateLabel("Activate Window: CTRL+E", 25, 433, -1, 20)
GUICtrlSetFont(-1, -1, -1, $GUI_FONTITALIC)
$Slider2 = GUICtrlCreateSlider(385, 432, 150, 18, BitOR($TBS_TOOLTIPS, $TBS_AUTOTICKS, $TBS_ENABLESELRANGE))
GUICtrlSetLimit($Slider2, 5, -5) ; control, max, min
GUICtrlSetData($Slider2, 0)
$labRate = GUICtrlCreateLabel("Rate: " & GUICtrlRead($Slider2), 345, 432, -1, -1)
GUISetState(@SW_SHOW)
#EndRegion

$Default = _StartTTS()
If Not IsObj($Default) Then
    MsgBox(0, 'Error', 'Failed create object')
    Exit
EndIf

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $Slider1
			SetVol("slider")
		Case $Slider2
			GUICtrlSetData($labRate, "Rate: " & GUICtrlRead($Slider2))
			 _SetRate($Default, GUICtrlRead($Slider2))
		Case $inVolume
			SetVol("input")
		Case $butAlert
			Alert()
		Case $butTTS
			TTS()
		Case $butClear
			Clear()
		Case $butFontSizeUp
			$fontSize += 2
			AdjustSize()
		Case $butFontSizeDown
			$fontSize -= 2
			AdjustSize()
		Case $butYes
			PlaySound("yes")
		Case $butNo
			PlaySound("no")
		Case $chkOnTop
			CheckOnTopStatus()
		Case $GUI_EVENT_CLOSE
			Terminate()
	EndSwitch

	If WinActive($winTitle) = 0 Then
		HotKeySet("^s")
		HotKeySet("^a")
	Else
		HotKeySet("^q", "Alert")
		HotKeySet("^s", "TTS")
		HotKeySet("^a", "Clear")
	EndIf

WEnd

Func ActiveWindow()
	WinActivate($winTitle)
	ControlFocus("","",$Edit1)
EndFunc

Func CheckOnTopStatus()
	If GUICtrlRead($chkOnTop) = 1 Then ;checked
		WinSetOnTop($winTitle, "", $WINDOWS_ONTOP)
	Else
		WinSetOnTop($winTitle, "", $WINDOWS_NOONTOP)
	EndIf
EndFunc

Func Alert()
	If GUICtrlRead($radSound1) = 1 Then
		SoundPlay(@WindowsDir & "\Media\Alarm03.wav")
	ElseIf GUICtrlRead($radSound2) = 1 Then
		SoundPlay(@WindowsDir & "\Media\Alarm09.wav")
	ElseIf GUICtrlRead($radSound3) = 1 Then
		SoundPlay(@WindowsDir & "\Media\Alarm10.wav")
	Else
		SoundPlay(@WindowsDir & "\Media\Alarm01.wav")
	EndIf
EndFunc

Func PlaySound($a)
	If $a = "yes" Then
		SoundPlay(@ScriptDir & "\yes.mp3")
	ElseIf $a = "no" Then
		SoundPlay(@ScriptDir & "\no.mp3")
	EndIf
EndFunc

Func SetVol($a)
	If $a = "slider" Then
		SoundSetWaveVolume(GUICtrlRead($Slider1))
		GUICtrlSetData($inVolume, GUICtrlRead($Slider1))
	EndIf
	If $a = "input" Then
		SoundSetWaveVolume(GUICtrlRead($inVolume))
		GUICtrlSetData($Slider1, GUICtrlRead($inVolume))
	EndIf
	If GUICtrlRead($inVolume) = 0 Or GUICtrlRead($Slider1) = 0 Then
		GUICtrlSetState($labMuted, $GUI_SHOW)
	Else
		GUICtrlSetState($labMuted, $GUI_HIDE)
	EndIf
EndFunc

Func Clear()
	ControlSetText("","",$Edit1,"")
EndFunc

Func TTS()
	_Speak($Default, GUICtrlRead($Edit1))
EndFunc

Func AdjustSize()
	GUICtrlSetFont($Edit1, $fontSize, 500)
EndFunc

Func Terminate()
	Exit 1
EndFunc