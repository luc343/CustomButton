Attribute VB_Name = "Lib_Misc"
Option Explicit
Option Private Module

'---------------------------------------------------
'
'                    Lib_Misc
'
' Note: Not accessable outside the Add-in
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Public Type SSIZE
	cx As Single
	cy As Single
End Type

'**********************************
'
'         Pixel X to point
'
'**********************************

Public Function PixelXToPoints(PixelVal As Long) As Single
	Dim hdc As LongPtr

	hdc = GetDC(0)
	PixelXToPoints = PixelVal * 72 / GetDeviceCaps(hdc, LOGPIXELSX)
	ReleaseDC 0, hdc
End Function

'**********************************
'
'         Pixel Y to point
'
'**********************************

Public Function PixelYToPoints(PixelVal As Long) As Single
	Dim hdc As LongPtr

	hdc = GetDC(0)
	PixelYToPoints = PixelVal * 72# / GetDeviceCaps(hdc, LOGPIXELSY)
	ReleaseDC 0, hdc
End Function

'**********************************
'
'         Is array empty
'
'**********************************

Public Function IsEmptyArr(ByRef Arr As Variant) As Boolean
	Dim idx As Long

	On Error GoTo Err
	idx = UBound(Arr)
	On Error GoTo 0

	IsEmptyArr = False   'Arr is not empty
	Exit Function

Err:
	On Error GoTo 0
	IsEmptyArr = True    'Arr is empty
End Function

'**********************************
'
'    Is valid VBA varible name
'
'**********************************

Public Function IsValidVariableName(Name As String) As Boolean
	Dim VBScript As Object
	Dim i As Long

	If Len(Name) > 255 Then
		IsValidVariableName = False         'name cannot exceeds 255 chars
		Exit Function
	End If

	Set VBScript = CreateObject("VBScript.RegExp")

	With VBScript
		.Pattern = "^[A-Za-z][A-Za-z0-9_]*$"
		.IgnoreCase = True
		.Global = False

		If Not .test(Name) Then
			IsValidVariableName = False     'name does not match the pattern
			Set VBScript = Nothing

			Exit Function
		End If
	End With

	Set VBScript = Nothing

	If IsReservedKeyword(Name) Then
		IsValidVariableName = False         'name cannot be a reserved key word
		Exit Function
	End If

	IsValidVariableName = True  'valid variable name
End Function

Private Function IsReservedKeyword(Name As String) As Boolean
	Dim ReservedKeywords As Variant
	Dim Low As Long, High As Long, Mid As Long
	Dim CompareResult As Long

	ReservedKeywords = Array("Alias", "And", "As", "Boolean", "Byte", "ByRef", "ByVal", "Call", "Case", "Const", "Date", _
							 "Declare", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "End", "EndIf", "Eqv", "Erase", _
							 "Error", "Exit", "Explicit", "False", "For", "Friend", "Function", "Get", "GoSub", "GoTo", _
							 "If", "Imp", "In", "Integer", "Is", "Let", "Lib", "Like", "Long", "Loop", "Me", "Mod", _
							 "New", "Next", "Not", "Nothing", "Object", "On", "Option", "Optional", "Or", "ParamArray", _
							 "Private", "Public", "ReDim", "Resume", "Return", "Select", "Set", "Static", "Step", _
							 "Stop", "String", "Sub", "Then", "To", "True", "Type", "TypeOf", "Until", "Variant", _
							 "Wend", "While", "With", "Xor")

	Low = LBound(ReservedKeywords)
	High = UBound(ReservedKeywords)

	Do While Low <= High
		Mid = (Low + High) \ 2
		CompareResult = StrComp(Name, ReservedKeywords(Mid), vbTextCompare)
		Select Case CompareResult
			Case 0
				IsReservedKeyword = True
				Exit Function
			Case Is < 0
				High = Mid - 1
			Case Is > 0
				Low = Mid + 1
		End Select
	Loop

	IsReservedKeyword = False
End Function

'**********************************
'
'        Get vbKey constant
'
'**********************************

Public Function GetVbKey(ByVal KeyCode As MSForms.ReturnInteger) As String
	Dim KeyName As String

	KeyName = ""
	If (CustomButtons.Trace And TR_EVENTS) = TR_EVENTS Then
		Select Case KeyCode
			Case 3: KeyName = "Cancel"      'cancel
			Case 8: KeyName = "Back"        'backspace
			Case 9: KeyName = "Tab"         'tab
			Case 12: KeyName = "Clear"      'clear
			Case 13: KeyName = "Return"     'enter
			Case 16: KeyName = "Shift"      'shift
			Case 17: KeyName = "Control"    'ctrl
			Case 18: KeyName = "Menu"       'alt
			Case 19: KeyName = "Pause"      'pause
			Case 20: KeyName = "Capital"    'caps lock
			Case 27: KeyName = "Escape"     'escape
			Case 32: KeyName = "Space"      'spacebar
			Case 33: KeyName = "PageUp"     'page up
			Case 34: KeyName = "PageDown"   'page down
			Case 35: KeyName = "End"        'end
			Case 36: KeyName = "Home"       'home
			Case 37: KeyName = "Left"       'left arrow
			Case 38: KeyName = "Up"         'up arrow
			Case 39: KeyName = "Right"      'right arrow
			Case 40: KeyName = "Down"       'down arrow
			Case 41: KeyName = "Select"     'select
			Case 42: KeyName = "Print"      'print
			Case 43: KeyName = "Execute"    'execute
			Case 44: KeyName = "Snapshot"   'snapshot
			Case 45: KeyName = "Insert"     'insert
			Case 46: KeyName = "Delete"     'delete
			Case 47: KeyName = "Help"       'help
			Case 48 To 57: KeyName = Chr$(KeyCode)              'numbers 0-9
			Case 65 To 90: KeyName = Chr$(KeyCode)              'letters A-Z
			Case 96 To 105: KeyName = "Numpad" & KeyCode - 96   'numpad 0 to 9
			Case 106: KeyName = "Multiply"  'numpad multiply
			Case 107: KeyName = "Add"       'numpad add
			Case 108: KeyName = "Separator" 'numpad enter
			Case 109: KeyName = "Subtract"  'numpad subtract
			Case 110: KeyName = "Decimal"   'numpad decimal point
			Case 111: KeyName = "Divide"    'numpad divide
			Case 112 To 127: KeyName = "F" & KeyCode - 111   'F1 to F9
			Case 144: KeyName = "Numlock"   'numlock
		End Select

		If KeyName <> "" Then
			KeyName = "vbKey" & KeyName & ":"
		End If
	End If

	GetVbKey = KeyName & KeyCode
End Function

'**********************************
'
'     Get mouse vbKey constant
'
'**********************************

Public Function GetVbMKey(ByVal KeyCode As Integer) As String
	Dim KeyName As String

	KeyName = ""
	If (CustomButtons.Trace And TR_EVENTS) = TR_EVENTS Then
		Select Case KeyCode
			Case 1: KeyName = "LButton"     'left mouse button
			Case 2: KeyName = "RButton"     'right mouse button
			Case 4: KeyName = "MButton"     'middle mouse button
		End Select

		If KeyName <> "" Then
			KeyName = "vbKey" & KeyName & ":"
		End If
	End If

	GetVbMKey = KeyName & KeyCode
End Function

'**********************************
'
'        Get modifier keys
'
'**********************************

Public Function GetModKey(Shift As Integer) As String
	Dim KeyName As String

	KeyName = ""
	If (CustomButtons.Trace And TR_EVENTS) = TR_EVENTS Then
		If Shift And 1 Then
			KeyName = KeyName & "Shift+"
		End If

		If Shift And 2 Then
			KeyName = KeyName & "Ctrl+"
		End If

		If Shift And 4 Then
			KeyName = KeyName & "Alt+"
		End If
	End If

	If KeyName = "" Then
		GetModKey = Shift
	Else
		GetModKey = Mid$(KeyName, 1, Len(KeyName) - 1) & ":" & Shift
	End If
End Function
