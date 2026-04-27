VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'
'                Class CustomButton
'
' Note: Class is PublicNotCreatable
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr      'button fade-in and fade-out

Private Theme As New CustomTheme

Private EventCatcher As New EventCatcher

'Label controls include the following events:
'    1. Click
'    2. DblClick
'    3. MouseDown
'    4. MouseUp
'    5. MouseMove
'
'CommandButton controls include the following additional events:
'    1. KeyPress
'    2. KeyDown
'    3. KeyUp
'
'Note: Enter & Exit events don't work with public events
'      Using EventCatcher to receive Enter & Exit events

Public Event OnEnter()
Public Event OnExit(ByVal Cancel As MSForms.ReturnBoolean)
Public Event AfterUpdate()
Public Event BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

Public Event Click()
Public Event DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Public Event KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Public Event KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Private WithEvents ButtonBody As MSForms.Label      'controls that make up a single button
Attribute ButtonBody.VB_VarHelpID = -1
Private WithEvents ButtonInnerBody As MSForms.Label
Attribute ButtonInnerBody.VB_VarHelpID = -1
Private WithEvents ButtonContent As MSForms.Label
Attribute ButtonContent.VB_VarHelpID = -1
Private WithEvents ButtonButton As MSForms.CommandButton
Attribute ButtonButton.VB_VarHelpID = -1

Private WithEvents ParentForm As MSForms.UserForm
Attribute ParentForm.VB_VarHelpID = -1
Private WithEvents ParentFrame As MSForms.Frame
Attribute ParentFrame.VB_VarHelpID = -1
Private WithEvents ParentPage As MSForms.MultiPage
Attribute ParentPage.VB_VarHelpID = -1

Private InFocus As Boolean
Private InFade As Boolean

Private This As ButtonTypes

Private Type ButtonTypes
	Parent As Object
	ParentName As String
	VarName As String

	Name As String
	Caption As String
	Accelerator As String
	Top As Single
	Left As Single
	Width As Single
	Height As Single

	BorderStyle As fmBorderStyle

	Font As New StdFont

	Tag As String
	TabIndex As Integer
	TabStop As Boolean
	Cancel As Boolean
	Default As Boolean
	Visible As Boolean
	Enabled As Boolean

	DefaultOverride As Boolean      'override default button look
End Type

Private BackGroundColor As LayoutColors

Private Enum LayoutColors
	CBBackColor = 0
	CBHoverDownColor = 1
	CBHoverBackColor = 2
End Enum

'######################## Button Methods ##########################

Private Sub Class_Initialize()
	Theme.Init  'set all the button colors

	This.ParentName = CustomButtons.Def_ParentName
	This.VarName = CustomButtons.Def_VarName

	This.Name = "CustomButtom" & CustomButtons.ButtonCount
	This.Caption = "CustomButton" & CustomButtons.ButtonCount
	This.Accelerator = CustomButtons.Def_Accelerator

	This.Font.Name = CustomButtons.Def_FontName
	This.Font.SIZE = CustomButtons.Def_FontSize
	This.Font.Bold = CustomButtons.Def_FontBold
	This.Font.Italic = CustomButtons.Def_FontItalic
	This.Font.Underline = CustomButtons.Def_FontUnderline
	This.Font.Strikethrough = CustomButtons.Def_FontStrikethrough

	This.Height = CustomButtons.Def_Height
	This.Width = CustomButtons.Def_Width

	This.BorderStyle = fmBorderStyleSingle

	This.Tag = CustomButtons.Def_Tag
	This.TabIndex = CustomButtons.Def_TabIndex
	This.TabStop = CustomButtons.Def_TabStop
	This.Cancel = CustomButtons.Def_Cancel
	This.Default = CustomButtons.Def_Default
	This.Visible = CustomButtons.Def_Visible
	This.Enabled = CustomButtons.Def_Enabled

	This.DefaultOverride = False

	InFocus = False
	CustomButtons.state = 0
End Sub

Friend Sub Initialize()
	Dim FontSize As SSIZE

	If This.Parent Is Nothing Or IsNull(This.Parent) Then
		Err.Raise vbObjectError + 1001, "CustomButton", "CustomButton cannot be initialized as the parent is not set."
	End If

	If TypeOf This.Parent Is MSForms.UserForm Then
		Set ParentForm = This.Parent
	ElseIf TypeOf This.Parent Is MSForms.Frame Then
		Set ParentFrame = This.Parent
	ElseIf TypeOf This.Parent Is MSForms.MultiPage Then
		Set ParentPage = This.Parent
	Else
		Err.Raise vbObjectError + 1002, "CustomButton", "Parent of CustomButton is not set to either a UserForm, Frame or MultiPage object."
	End If

	If ButtonButton Is Nothing Then
		Set ButtonButton = This.Parent.Controls.Add("Forms.CommandButton.1", This.Name & CustomButtons.ID, This.Visible)
	End If

	With ButtonButton       'hidden button
		.Name = This.Name & CustomButtons.ID
		.Accelerator = This.Accelerator
		.Height = This.Height
		.Width = This.Width
		.Top = This.Top
		.Left = This.Left
		If This.TabIndex <> CustomButtons.Def_TabStop Then .TabIndex = This.TabIndex
		.TabStop = This.TabStop
		.Cancel = This.Cancel
		.Default = This.Default
		.Visible = This.Visible
		.Enabled = This.Enabled
	End With

	If ButtonBody Is Nothing Then
		Set ButtonBody = This.Parent.Controls.Add("Forms.Label.1", This.Name, This.Visible)
	End If

	With ButtonBody         'button border, background color, and image
		.Name = This.Name & CustomButtons.ID
		.Height = This.Height
		.Width = This.Width
		.Top = This.Top
		.Left = This.Left
		.BorderStyle = This.BorderStyle
		.SpecialEffect = fmSpecialEffectFlat
		.Visible = This.Visible
		.Enabled = This.Enabled

		If This.Enabled Then
			.BackColor = Theme.BackColor
			.BorderColor = IIf(This.Default, Theme.DefaultBorderColor, Theme.BorderColor)
		Else
			.BackColor = Theme.DisabledBackColor
			.BorderColor = Theme.DisabledBorderColor
		End If
	End With

	If ButtonInnerBody Is Nothing Then
		Set ButtonInnerBody = This.Parent.Controls.Add("Forms.Label.1", This.Name, This.Visible)
	End If

	With ButtonInnerBody    'internal button area
		.Name = This.Name & CustomButtons.ID
		.Height = This.Height - PixelYToPoints(2)
		.Width = This.Width - PixelXToPoints(2)
		.Top = ButtonBody.Top + (ButtonBody.Height - .Height) / 2
		.Left = ButtonBody.Left + (ButtonBody.Width - .Width) / 2
		.BackStyle = fmBackStyleTransparent
		.BorderStyle = fmBorderStyleNone
		.BorderStyle = IIf(This.Enabled, .BorderStyle, fmBorderStyleNone)
		.BackColor = vbBlue
		.BorderColor = Theme.DefaultBorderColor
		.SpecialEffect = fmSpecialEffectFlat
		.Visible = This.Visible
		.Enabled = This.Enabled
	End With

	If ButtonContent Is Nothing Then
		Set ButtonContent = This.Parent.Controls.Add("Forms.Label.1", This.Name, This.Visible)
	End If

	With ButtonContent     'button text
		Set .Font = This.Font

		.AutoSize = False   'reset label size
		.WordWrap = False
		.Name = This.Name & CustomButtons.ID
		.Caption = This.Caption
		.Accelerator = This.Accelerator
		.TextAlign = fmTextAlignCenter
		.AutoSize = True    'label will autosize

		If .Width > This.Width Then .Width = This.Width
		If .Height > This.Height Then .Height = This.Height
		.Top = This.Top + (This.Height - .Height) / 2
		.Left = This.Left + (This.Width - .Width) / 2

		.BackStyle = fmBackStyleTransparent
		.BorderStyle = fmBorderStyleNone
		.BackColor = vbRed
		.SpecialEffect = fmSpecialEffectFlat
		.ForeColor = IIf(This.Enabled, Theme.ForeColor, Theme.DisabledForeColor)
		.Visible = This.Visible
		.Enabled = True
	End With

	CustomButtons.Reinitialize This.Parent  'update all global properties
End Sub

Friend Sub Reinitialize()
	This.Default = ButtonButton.Default
	This.Cancel = ButtonButton.Cancel
	This.TabIndex = ButtonButton.TabIndex
End Sub

Public Sub Repaint()
	If Not ButtonBody Is Nothing Then
		TraceMethod "Repaint()"

		Initialize
	End If
End Sub

Friend Sub Unregister()
	EventCatcher.Clear
End Sub

Friend Sub Register()
	EventCatcher.Ctrl = ButtonButton
	EventCatcher.Parent = Me
End Sub

Public Sub SetFocus()
	TraceMethod "SetFocus()"

	On Error Resume Next
	This.Parent.Controls(This.Name & CustomButtons.ID).SetFocus

	If CustomButtons.state = vbKeyLButton Then
		SetLayoutColor CBHoverDownColor, "SetFocus"
	Else
		SetLayoutColor CBBackColor, "SetFocus"
	End If
	On Error GoTo 0
End Sub

Private Sub SetLayoutColor(ByVal LayoutColor As LayoutColors, Caller As String)
	Dim Layout As String

	If This.Enabled = False Then Exit Sub

	If LayoutColor = CBBackColor Then
		InFade = False
		Layout = "CBBackColor"
	ElseIf LayoutColor = CBHoverBackColor Then
		InFade = True
		Layout = "CBHoverBackColor"
	ElseIf LayoutColor = CBHoverDownColor Then
		InFade = True
		Layout = "CBHoverDownColor"
	Else
		Layout = CStr(LayoutColor)
	End If

	TraceLayoutColor "SetLayoutColor(" & Layout & ", """ & Caller & """)"
	BackGroundColor = LayoutColor

	Select Case LayoutColor
		Case CBBackColor
			With ButtonBody
				If This.DefaultOverride Then
					ButtonInnerBody.BorderStyle = fmBorderStyleNone

					.BackColor = Theme.BackColor
					.BorderColor = Theme.BorderColor
					.BorderStyle = This.BorderStyle
				Else
					ButtonInnerBody.BorderStyle = IIf(InFocus, fmBorderStyleSingle, fmBorderStyleNone)

					.BackColor = Theme.BackColor
					.BorderColor = IIf(This.Default Or InFocus, Theme.DefaultBorderColor, Theme.BorderColor)
					.BorderStyle = This.BorderStyle
				End If
			End With

			ButtonContent.ForeColor = Theme.ForeColor

		Case CBHoverDownColor
			With ButtonBody
				ButtonInnerBody.BorderStyle = fmBorderStyleNone

				.BorderColor = Theme.HoverBorderColor
				.BorderStyle = This.BorderStyle

'                FadeBackgroundColor Theme.HoverBackColor, Theme.HoverDownColor
'
'                If InFade Then
					.BackColor = Theme.HoverDownColor
'                Else
'                   .BackColor = Theme.BackColor
'                End If
			End With

			ButtonContent.ForeColor = Theme.ForeColor

		Case CBHoverBackColor
			With ButtonBody
				ButtonInnerBody.BorderStyle = fmBorderStyleNone

				.BorderColor = Theme.HoverBorderColor
				.BorderStyle = This.BorderStyle

				FadeBackgroundColor Theme.BackColor, Theme.HoverBackColor

				If InFade Then
					.BackColor = Theme.HoverBackColor
				Else
					.BackColor = Theme.BackColor
				End If
			End With

			ButtonContent.ForeColor = Theme.HoverForeColor
	End Select

End Sub

Private Sub FadeBackgroundColor(ByVal StartColor As Long, ByVal FinalColor As Long)
	Dim CurrentColor As Long
	Dim ColorStep As Long
	Dim Tick As LongPtr

	With ButtonBody
		For ColorStep = 1 To 255
			CurrentColor = RGB( _
				InterpolateColorComponent(StartColor, FinalColor, ColorStep / 255, 1), _
				InterpolateColorComponent(StartColor, FinalColor, ColorStep / 255, 2), _
				InterpolateColorComponent(StartColor, FinalColor, ColorStep / 255, 3) _
			)

			.BackColor = CurrentColor

			Tick = GetTickCount()
			If ColorStep Mod 10 = 0 Then
				Do
					DoEvents
					If Not InFade Then Exit Sub     'exit fade
				Loop While GetTickCount() - Tick < 15
			End If
		Next ColorStep
	End With
End Sub

Private Function InterpolateColorComponent(StartColor As Long, FinalColor As Long, Progress As Single, ComponentIndex As Integer) As Integer
	Dim StartComponent As Integer
	Dim FinalComponent As Integer

	StartComponent = (StartColor Mod 256 ^ ComponentIndex) \ 256 ^ (ComponentIndex - 1)
	FinalComponent = (FinalColor Mod 256 ^ ComponentIndex) \ 256 ^ (ComponentIndex - 1)

	InterpolateColorComponent = StartComponent + (FinalComponent - StartComponent) * Progress
End Function

Private Sub Trace(Text As String)
	Dim Name As String, Parts() As String

	Name = This.ParentName & "." & This.VarName & "."
	If Name <> CustomButtons.LastName Then Debug.Print

	Debug.Print "    " & Text

	CustomButtons.LastName = Name
	CustomButtons.LastParentName = This.ParentName
End Sub

Private Sub TraceMethod(Text As String)
	If (CustomButtons.Trace And TR_METHODS) = TR_METHODS Then
		Trace This.ParentName & "." & This.VarName & "." & Text
	End If
End Sub

Private Sub TraceEvent(Text As String)
	If (CustomButtons.Trace And TR_EVENTS) = TR_EVENTS Then
		Trace This.ParentName & "." & This.VarName & "." & Text
	End If
End Sub

Private Sub TraceProperty(Text As String)
	Dim Name As String, Parts() As String

	If (CustomButtons.Trace And TR_PROPERTIES) = TR_PROPERTIES Then
		Name = This.ParentName & "." & This.VarName & "."

		If InStr(1, Text, " --> ") = 0 Then
			Text = Replace$(Text, "(", " = ")
			Text = Name & Mid$(Text, 1, Len(Text) - 1)
		Else
			Text = Name & Replace(Text, "()", "")

			Parts = Split(Text, " --> ")
			Text = Parts(1) & " <-- " & Parts(0)
		End If

		Trace Text
	End If
End Sub

Private Sub TraceLayoutColor(Text As String)
	If (CustomButtons.Trace And TR_LAYOUTCOLORS) = TR_LAYOUTCOLORS Then
		Trace This.ParentName & "." & This.VarName & "." & Text
	End If
End Sub

Private Sub TraceMouseMoveEvent(Text As String)
	If (CustomButtons.Trace And TR_MOUSEMOVE) = TR_MOUSEMOVE Then
		Trace This.ParentName & "." & This.VarName & "." & Text
	End If
End Sub

Public Sub TraceUserForm(Text As String)
	If (CustomButtons.Trace And TR_USERFORM) = TR_USERFORM Then
		Debug.Print "--> " & This.VarName & "_" & Text
	End If
End Sub

'######################### Button Events ##########################

'----------------------------------
'
'           Click Events
'
'----------------------------------

Private Sub ButtonBody_Click()
	TraceEvent "ButtonBody_Click()"

	SetLayoutColor CBBackColor, "ButtonBody_Click"
	RaiseEvent Click
End Sub

Private Sub ButtonInnerBody_Click()
	TraceEvent "ButtonInnerBody_Click()"

	SetLayoutColor CBBackColor, "ButtonInnerBody_Click"
	RaiseEvent Click
End Sub

Private Sub ButtonContent_Click()
	If Not This.Enabled Then Exit Sub
	TraceEvent "ButtonContent_Click()"

	SetLayoutColor CBBackColor, "ButtonContent_Click"
	RaiseEvent Click
End Sub

'----------------------------------
'
'       Double Click Events
'
'----------------------------------

Private Sub ButtonBody_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
	TraceEvent "ButtonBody_DblClick(" & Cancel & ")"

	SetLayoutColor CBBackColor, "ButtonBody_DblClick"
	RaiseEvent DblClick(Cancel)
End Sub

Private Sub ButtonInnerBody_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
	TraceEvent "ButtonInnerBody_DblClick(" & Cancel & ")"

	SetLayoutColor CBBackColor, "ButtonBody_DblClick"
	RaiseEvent DblClick(Cancel)
End Sub

Private Sub ButtonContent_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
	If Not This.Enabled Then Exit Sub

	TraceEvent "ButtonContent_DblClick(" & Cancel & ")"

	SetLayoutColor CBBackColor, "ButtonBody_DblClick"
	RaiseEvent DblClick(Cancel)
End Sub

'----------------------------------
'
'          KeyPress Events
'
'----------------------------------

Private Sub ButtonBody_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	TraceEvent "ButtonBody_KeyPress(""" & Chr$(KeyAscii) & """:" & KeyAscii & ")"

	RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub ButtonInnerBody_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	TraceEvent "ButtonInnerBody_KeyPress(""" & Chr$(KeyAscii) & """:" & KeyAscii & ")"

	RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub ButtonContent_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	TraceEvent "ButtonContent_KeyPress(""" & Chr$(KeyAscii) & """:" & KeyAscii & ")"

	RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub ButtonButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	TraceEvent "ButtonButton_KeyPress(""" & Chr$(KeyAscii) & """:" & KeyAscii & ")"

	RaiseEvent KeyPress(KeyAscii)
End Sub

'----------------------------------
'
'          KeyDown Events
'
'----------------------------------

Private Sub ButtonBody_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	TraceEvent "ButtonBody_KeyDown(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ButtonInnerBody_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	TraceEvent "ButtonInnerBody_KeyDown(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ButtonContent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	TraceEvent "ButtonContent_KeyDown(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ButtonButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	TraceEvent "ButtonButton_KeyDown(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyDown(KeyCode, Shift)

	If KeyCode = vbKeyTab Then      'tab key
		SetLayoutColor BackGroundColor, "ButtonButton_KeyDown"
	End If
End Sub

'----------------------------------
'
'           KeyUp Events
'
'----------------------------------

Private Sub ButtonBody_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	TraceEvent "ButtonBody_KeyUp(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ButtonInnerBody_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	TraceEvent "ButtonInnerBody_KeyUp(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ButtonContent_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	TraceEvent "ButtonContent_KeyUp(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ButtonButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
	If KeyCode = vbKeyTab Then          'tab key
		SetLayoutColor BackGroundColor, "ButtonButton_KeyUp"
	ElseIf KeyCode = vbKeyReturn Then   'enter key
		RaiseEvent Click
		Exit Sub
	ElseIf KeyCode = vbKeyEscape Then   'escape key
		RaiseEvent Click
		Exit Sub
	ElseIf Shift = 4 And LCase$(Chr$(KeyCode + 32)) = LCase$(This.Accelerator) Then
		RaiseEvent Click
		Exit Sub
	End If

	TraceEvent "ButtonButton_KeyUp(" & GetVbKey(KeyCode) & ", " & GetModKey(Shift) & ")"

	RaiseEvent KeyUp(KeyCode, Shift)

	If KeyCode = vbKeySpace Then        'space key
		RaiseEvent Click
		Exit Sub
	End If
End Sub

'----------------------------------
'
'        MouseMove Events
'
'----------------------------------

Private Sub ButtonBody_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	TraceMouseMoveEvent "ButtonBody_MouseMove(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	If BackGroundColor <> CBBackColor Then
		SetLayoutColor CBBackColor, "ButtonBody_MouseMove"
	End If

	CustomButtons.state = Button
	RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub ButtonInnerBody_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	TraceMouseMoveEvent "ButtonInnerBody_MouseMove(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	If BackGroundColor <> CBHoverBackColor Then
		SetLayoutColor IIf(Button = 1, CBHoverDownColor, CBHoverBackColor), "ButtonInnerBody_MouseMove"
	End If

	CustomButtons.state = Button
	RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub ButtonContent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	If Not This.Enabled Then Exit Sub   'only needed here

	TraceMouseMoveEvent "ButtonContent_MouseMove(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	If BackGroundColor <> CBHoverBackColor Then
		SetLayoutColor IIf(Button = 1, CBHoverDownColor, CBHoverBackColor), "ButtonContent_MouseMove"
	End If

	CustomButtons.state = Button
	RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub ParentForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	If BackGroundColor <> CBBackColor Then
		SetLayoutColor CBBackColor, "ParentForm_MouseMove"
	End If

	CustomButtons.state = Button
	RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'----------------------------------
'
'        MouseDown Events
'
'----------------------------------

Private Sub ButtonBody_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	TraceEvent "ButtonBody_MouseDown(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	CustomButtons.state = Button
	SetFocus
	RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub ButtonInnerBody_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	TraceEvent "ButtonInnerBody_MouseDown(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	If Button = vbKeyLButton Then
		If BackGroundColor <> CBHoverDownColor Then
			SetLayoutColor CBHoverDownColor, "ButtonInnerBody_MouseDown"
		End If

		CustomButtons.state = Button
		SetFocus
	End If

	RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub ButtonContent_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	If Not This.Enabled Then Exit Sub   'only needed here

	TraceEvent "ButtonContent_MouseDown(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	If Button = vbKeyLButton Then
		If BackGroundColor <> CBHoverDownColor Then
			SetLayoutColor CBHoverDownColor, "ButtonContent_MouseDown"
		End If

		CustomButtons.state = Button
		SetFocus
	End If

	RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'----------------------------------
'
'          MouseUp Events
'
'----------------------------------

Private Sub ButtonBody_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	TraceEvent "ButtonBody_MouseUp(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	CustomButtons.state = Button
	RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub ButtonInnerBody_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	If Button = vbKeyLButton Then
		If BackGroundColor <> CBBackColor Then
			SetLayoutColor CBBackColor, "ButtonInnerBody_MouseUp"
		End If
	End If

	TraceEvent "ButtonInnerBody_MouseUp(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	CustomButtons.state = Button
	RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub ButtonContent_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	If Not This.Enabled Then Exit Sub   'only needed here

	If Button = vbKeyLButton Then
		If BackGroundColor <> CBBackColor Then
			SetLayoutColor CBBackColor, "ButtonContent_MouseUp"
		End If
	End If

	TraceEvent "ButtonContent_MouseUp(" & GetVbMKey(Button) & ", " & GetModKey(Shift) & ", " & X & ", " & Y & ")"

	CustomButtons.state = Button
	RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'----------------------------------
'
'           OnEnter Event
'
'----------------------------------

Friend Sub RaiseEnter()
	InFocus = True
	TraceEvent "ButtonButton_Enter()"

	CustomButtons.OverrideDefault InFocus, This.Name, This.Parent

	If CustomButtons.state = vbKeyLButton Then    'left mouse button down
		SetLayoutColor CBHoverDownColor, "RaiseEnter"
	Else
		SetLayoutColor CBBackColor, "RaiseEnter"
	End If

	RaiseEvent OnEnter
End Sub

'----------------------------------
'
'           OnExit Event
'
'----------------------------------

Friend Sub RaiseExit(ByVal Cancel As MSForms.ReturnBoolean)
	InFocus = False
	TraceEvent "ButtonButton_OnExit(" & Cancel & ")"

	CustomButtons.OverrideDefault InFocus, This.Name, This.Parent

	SetLayoutColor CBBackColor, "RaiseExit"
	RaiseEvent OnExit(Cancel)
End Sub

'----------------------------------
'
'         AfterUpdate Event
'
'----------------------------------

Friend Sub RaiseAfterUpdate()
	TraceEvent "ButtonButton_AfterUpdate()"

	RaiseEvent AfterUpdate
End Sub

'----------------------------------
'
'         BeforeUpdate Event
'
'----------------------------------

Friend Sub RaiseBeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
	TraceEvent "ButtonButton_BeforeUpdate()"

	RaiseEvent BeforeUpdate(Cancel)
End Sub

'################### Get/Let Button Properties ####################

'Parent
Public Property Get Parent() As Object
	TraceProperty "Parent() --> " & This.ParentName

	Set Parent = This.Parent
End Property

Friend Property Let Parent(Value As Object)
	If Value Is Nothing Or IsNull(Value) Then
		Err.Raise vbObjectError + 1003, "CustomButton", "Parent of CustomButton cannot be set to an object containing nothing."
	Else
		If ButtonBody Is Nothing Then
			Set This.Parent = Value
			This.ParentName = This.Parent.Name
		Else
			Err.Raise vbObjectError + 1004, "CustomButton", "Parent of CustomButton is already set to """ & This.ParentName & """ and cannot be changed."
		End If
	End If

	TraceProperty "Parent(" & This.ParentName & ")"
End Property

'VarName
Public Property Get VarName() As String
	TraceProperty "VarName() --> " & This.VarName

	VarName = This.VarName
End Property

Friend Property Let VarName(Value As String)
	If Not IsValidVariableName(Value) Then
		Err.Raise vbObjectError + 1005, "CustomButton", "Variable name for CustomButton is not valid and cannot be set to """ & Value & """."
	Else
		If ButtonBody Is Nothing Then
			This.VarName = Value
		Else
			Err.Raise vbObjectError + 1006, "CustomButton", "Variable name for CustomButton is already set to """ & This.VarName & """ and cannot be changed."
		End If
	End If

	TraceProperty "VarName(""" & Value & """)"
End Property

'Name
Public Property Get Name() As String
	TraceProperty "Name() --> " & This.Name

	Name = This.Name
End Property

Friend Property Let Name(Value As String)
	TraceProperty "Name(""" & Value & """)"

	If ButtonBody Is Nothing Then
		This.Name = Value
	Else
		Err.Raise vbObjectError + 1007, "CustomButton", "Name of CustomButton is already set to " & This.Name & " and cannot be changed."
	End If

End Property

'Caption
Public Property Get Caption() As String
	TraceProperty "Caption() --> " & This.Caption

	Caption = This.Caption
End Property

Public Property Let Caption(Value As String)
	TraceProperty "Caption(""" & Value & """)"

	If Not ButtonContent Is Nothing Then
		With ButtonContent
			.AutoSize = False   'reset label size
			.Caption = Value
			.AutoSize = True    'label will autosize

			If .Width > This.Width Then .Width = This.Width
			If .Height > This.Height Then .Height = This.Height
			.Top = This.Top + (This.Height - .Height) / 2
			.Left = This.Left + (This.Width - .Width) / 2
		End With
	End If

	This.Caption = Value
End Property

'Accelerator
Public Property Get Accelerator() As String
	TraceProperty "Accelerator() --> " & This.Accelerator

	Accelerator = This.Accelerator
End Property

Public Property Let Accelerator(Value As String)
	Dim Ch As String

	Ch = IIf(Len(Value) > 0, Mid$(Value, 1, 1), "")

	TraceProperty "Accelerator(""" & Ch & """)"

	If Not ButtonContent Is Nothing Then
		ButtonContent.Accelerator = Ch
	End If

	This.Accelerator = Ch
End Property

'BackColor
Public Property Get BackColor() As Long
	TraceProperty "BackColor() --> " & Theme.BackColor

	BackColor = Theme.BackColor
End Property

Public Property Let BackColor(Value As Long)
	TraceProperty "BackColor(" & Value & ")"

	Theme.BackColor = Value
End Property

'HoverBackColor
Public Property Get HoverBackColor() As Long
	TraceProperty "HoverBackColor() --> " & Theme.HoverBackColor

	HoverBackColor = Theme.HoverBackColor
End Property

Public Property Let HoverBackColor(Value As Long)
	TraceProperty "HoverBackColor(" & Value & ")"

	Theme.HoverBackColor = Value
End Property

'HoverBorderColor
Public Property Get HoverBorderColor() As Long
	TraceProperty "HoverBorderColor() --> " & Theme.HoverBorderColor

	HoverBorderColor = Theme.HoverBorderColor
End Property

Public Property Let HoverBorderColor(Value As Long)
	TraceProperty "HoverBorderColor(" & Value & ")"

	Theme.HoverBorderColor = Value
End Property

'HoverDownColor
Public Property Get HoverDownColor() As Long
	TraceProperty "HoverDownColor() --> " & Theme.HoverDownColor

	HoverDownColor = Theme.HoverDownColor
End Property

Public Property Let HoverDownColor(Value As Long)
	TraceProperty "HoverDownColor(" & Value & ")"

	Theme.HoverDownColor = Value
End Property

'DisabledBackColor
Public Property Get DisabledBackColor() As Long
	TraceProperty "DisabledBackColor() --> " & Theme.DisabledBackColor

	DisabledBackColor = Theme.DisabledBackColor
End Property

Public Property Let DisabledBackColor(Value As Long)
	TraceProperty "DisabledBackColor(" & Value & ")"

	Theme.DisabledBackColor = Value
End Property

'DisabledForeColor
Public Property Get DisabledForeColor() As Long
	TraceProperty "DisabledForeColor() --> " & Theme.DisabledForeColor

	DisabledForeColor = Theme.DisabledForeColor
End Property

Public Property Let DisabledForeColor(Value As Long)
	TraceProperty "DisabledForeColor(" & Value & ")"

	Theme.DisabledForeColor = Value
End Property

'DisabledBorderColor
Public Property Get DisabledBorderColor() As Long
	TraceProperty "DisabledBorderColor() --> " & Theme.DisabledBorderColor

	DisabledBorderColor = Theme.DisabledBorderColor
End Property

Public Property Let DisabledBorderColor(Value As Long)
	TraceProperty "DisabledBorderColor(" & Value & ")"

	Theme.DisabledBorderColor = Value
End Property

'DefaultBorderColor
Public Property Get DefaultBorderColor() As Long
	TraceProperty "DefaultBorderColor() --> " & Theme.DefaultBorderColor

	DefaultBorderColor = Theme.DefaultBorderColor
End Property

Public Property Let DefaultBorderColor(Value As Long)
	TraceProperty "DefaultBorderColor(" & Value & ")"

	Theme.HoverBorderColor = Value
End Property

'ForeColor
Public Property Get ForeColor() As Long
	TraceProperty "ForeColor() --> " & Theme.ForeColor

	ForeColor = Theme.ForeColor
End Property

Public Property Let ForeColor(Value As Long)
	TraceProperty "ForeColor(" & Value & ")"

	Theme.ForeColor = Value
End Property

'HoverForeColor
Public Property Get HoverForeColor() As Long
	TraceProperty "HoverForeColor() --> " & Theme.HoverForeColor

	HoverForeColor = Theme.HoverForeColor
End Property

Public Property Let HoverForeColor(Value As Long)
	TraceProperty "HoverForeColor(" & Value & ")"

	Theme.HoverForeColor = Value
End Property

'BorderColor
Public Property Get BorderColor() As Long
	TraceProperty "BorderColor() --> " & Theme.BorderColor

	BorderColor = Theme.BorderColor
End Property

Public Property Let BorderColor(Value As Long)
	TraceProperty "BorderColor(" & Value & ")"

	Theme.BorderColor = Value
End Property

'ExcelGreen
Friend Property Get ExcelGreen() As Boolean
	TraceProperty "ExcelGreen() --> " & Theme.ExcelGreen

	ExcelGreen = Theme.ExcelGreen
End Property

Friend Property Let ExcelGreen(Value As Boolean)
	TraceProperty "ExcelGreen(" & Value & ")"

	Theme.ExcelGreen = Value
End Property

'Top
Public Property Get Top() As Single
	TraceProperty "Top() --> " & This.Top

	Top = This.Top
End Property

Public Property Let Top(Value As Single)
	TraceProperty "Top(" & Value & ")"

	This.Top = Value

	Repaint
End Property

'Left
Public Property Get Left() As Single
	TraceProperty "Left() --> " & This.Left

	Left = This.Left
End Property

Public Property Let Left(Value As Single)
	TraceProperty "Left(" & Value & ")"

	This.Left = Value

	Repaint
End Property

'Width
Public Property Get Width() As Single
	TraceProperty "Width() --> " & This.Width

	Width = This.Width
End Property

Public Property Let Width(Value As Single)
	TraceProperty "Width(" & Value & ")"

	This.Width = Value

	Repaint
End Property

'Height
Public Property Get Height() As Single
	TraceProperty "Height() --> " & This.Height

	Height = This.Height
End Property

Public Property Let Height(Value As Single)
	TraceProperty "Height(" & Value & ")"

	This.Height = Value

	Repaint
End Property

'Font
Public Property Get Font() As StdFont
	TraceProperty "Font() --> " & "[""" & This.Font.Name & """, " & This.Font.SIZE & ", B=" & This.Font.Bold & ", I=" & This.Font.Italic & ", U=" & This.Font.Underline & ", S=" & This.Font.Strikethrough & "]"

	Set Font = This.Font
End Property

'Font
Public Property Set Font(Value As StdFont)
	TraceProperty "Font(Font)"

	If Not ButtonContent Is Nothing Then
'        Set ButtonContent.Font = Value

		ButtonContent.Font.Name = Value.Name
		ButtonContent.Font.SIZE = Value.SIZE
		ButtonContent.Font.Bold = Value.Bold
		ButtonContent.Font.Italic = Value.Italic
		ButtonContent.Font.Underline = Value.Underline
		ButtonContent.Font.Strikethrough = Value.Strikethrough
	End If

	Set This.Font = Value

	Repaint
End Property

'Tag
Public Property Get Tag() As String
	TraceProperty "Tag() --> " & This.Tag

	Tag = This.Tag
End Property

Public Property Let Tag(Value As String)
	TraceProperty "Tag(" & Value & ")"

	This.Tag = Value
End Property

'TabIndex
Public Property Get TabIndex() As Integer
	TraceProperty "TabIndex() --> " & This.TabIndex

	TabIndex = This.TabIndex
End Property

Public Property Let TabIndex(Value As Integer)
	TraceProperty "TabIndex(" & Value & ")"

	If Not ButtonButton Is Nothing Then
		ButtonButton.TabIndex = Value
		CustomButtons.Reinitialize This.Parent  'global button setting
	End If

	This.TabIndex = Value
End Property

'TabStop
Public Property Get TabStop() As Boolean
	TraceProperty "TabStop() --> " & This.TabStop

	TabStop = This.TabStop
End Property

Public Property Let TabStop(Value As Boolean)
	TraceProperty "TabStop(" & Value & ")"

	If Not ButtonButton Is Nothing Then
		ButtonButton.TabStop = Value
	End If

	This.TabStop = Value
End Property

'Cancel
Public Property Get Cancel() As Boolean
	TraceProperty "Cancel() --> " & This.Cancel

	Cancel = This.Cancel
End Property

Public Property Let Cancel(Value As Boolean)
	TraceProperty "Cancel(" & Value & ")"

	If Not ButtonButton Is Nothing Then
		ButtonButton.Cancel = Value
		CustomButtons.Reinitialize This.Parent  'global button setting
	End If

	This.Cancel = Value
End Property

'Default
Public Property Get Default() As Boolean
	TraceProperty "Default() --> " & This.Default

	Default = This.Default
End Property

Public Property Let Default(Value As Boolean)
	TraceProperty "Default(" & Value & ")"

	If Not ButtonButton Is Nothing Then
		ButtonButton.Default = Value
		CustomButtons.Reinitialize This.Parent  'global button setting
	End If

	This.Default = Value
End Property

'Visible
Public Property Get Visible() As Boolean
	TraceProperty "Visible() --> " & This.Visible

	Visible = This.Visible
End Property

Public Property Let Visible(Value As Boolean)
	TraceProperty "Visible(" & Value & ")"

	If Not ButtonBody Is Nothing Then
		ButtonBody.Visible = Value
		ButtonInnerBody.Visible = Value
		ButtonContent.Visible = Value
		ButtonButton.visable = Value
	End If

	This.Visible = Value
End Property

'Enabled
Public Property Get Enabled() As Boolean
	TraceProperty "Enabled() --> " & This.Enabled

	Enabled = This.Enabled
End Property

Public Property Let Enabled(Value As Boolean)
	TraceProperty "Enabled(" & Value & ")"

	If Not ButtonBody Is Nothing Then
		ButtonBody.Enabled = Value
		ButtonInnerBody.Enabled = Value
		'exclude ButtonContent as it includes text
		ButtonButton.Enabled = Value

		If Value Then
			ButtonBody.BackColor = Theme.BackColor
			ButtonBody.BorderColor = IIf(This.Default, Theme.DefaultBorderColor, Theme.BorderColor)

			ButtonContent.ForeColor = Theme.ForeColor
			ButtonInnerBody.BorderStyle = IIf(This.Default, fmBorderStyleSingle, fmBorderStyleNone)
		Else
			ButtonBody.BackColor = Theme.DisabledBackColor
			ButtonBody.BorderColor = Theme.DisabledBorderColor

			ButtonContent.ForeColor = Theme.DisabledForeColor
			ButtonInnerBody.BorderStyle = fmBorderStyleNone
		End If

	End If

	This.Enabled = Value
End Property

'Override
Friend Property Get DefaultOverride() As Boolean
	TraceProperty "DefaultOverride() --> " & This.DefaultOverride

	DefaultOverride = This.DefaultOverride
End Property

Friend Property Let DefaultOverride(Value As Boolean)
	TraceProperty "DefaultOverride(" & Value & ")"

	This.DefaultOverride = Value
	SetLayoutColor CBBackColor, "DefaultOverride"
End Property
