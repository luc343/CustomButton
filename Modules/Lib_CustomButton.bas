Attribute VB_Name = "Lib_CustomButton"
Option Explicit

'---------------------------------------------------
'
'                  Lib_CustomButton
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Public CustomButtons As New CustomButtons

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef nlpRect As RECT) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'GetWindowRect API
Private Type RECT
	Left As Long
	Top As Long
	Right As Long
	Bottom As Long
End Type

'SetWindowPos API
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10

'***************************************
'
'           Centre a userform
'
'***************************************

Public Sub CentreUserForm(hwnd As LongPtr)
	Dim FormRect As RECT
	Dim MsgRect As RECT
	Dim XPos As Long, YPos As Long

	GetWindowRect Application.hwnd, FormRect
	GetWindowRect hwnd, MsgRect

	XPos = (FormRect.Left + (FormRect.Right - FormRect.Left) / 2) - ((MsgRect.Right - MsgRect.Left) / 2)
	YPos = (FormRect.Top + (FormRect.Bottom - FormRect.Top) / 2) - ((MsgRect.Bottom - MsgRect.Top) / 2)
	SetWindowPos hwnd, 0, XPos, YPos, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub

'**********************************
'
'   Output tab index properties
'
'**********************************

Public Sub ShowTabIndexes(ByRef UserForm As Object)
	Dim List As New Scripting.Dictionary
	Dim Ctrl As Variant
	Dim TabStop As String
	Dim i As Long

	On Error Resume Next
	For Each Ctrl In UserForm.Controls
		TabStop = "-"
		TabStop = Ctrl.TabStop

		List.Add Ctrl.TabIndex, PadRight(TabStop, 5) & " " & PadRight(Ctrl.Name, 15) & " " & TypeName(Ctrl)
	Next Ctrl
	On Error GoTo 0

	For i = 0 To List.Count - 1     'debug.print sorted indexes
		Debug.Print PadLeft(i, 3) & " " & List(i)
	Next i
	Debug.Print
End Sub

Private Function PadRight(ByVal Text As String, Num As Long, Optional Char As String = " ") As String
	PadRight = Left$(Text & String(Num, Char), Num)
End Function

Private Function PadLeft(ByVal Text As String, Num As Long, Optional Char As String = " ") As String
	PadLeft = Right$(String(Num, Char) & Text, Num)
End Function
