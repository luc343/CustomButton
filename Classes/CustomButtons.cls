VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------
'
'                Class CustomButtons
'
' Note: Class is PublicNotCreatable
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

Private ButtonList() As New CustomButton    'list of all buttons created
Private ButtonUsed() As Boolean             'which buttons have been created
Private ButtonIdx As Long                   'next button index

Private MouseButtonState As Integer         'current mouse button state

Private ExcelGreenFlag As Boolean           'excel green

Private TraceCustomButtons As Long          'trace bitwise flags
Private TraceLastName As String             'last button name displayed (.ParentName.ButtonName.)
Private TraceLastParentName As String       'last parent name displayed

Public Enum TraceFlags
	TR_NONE = 0
	TR_EVENTS = 1
	TR_METHODS = 2
	TR_PROPERTIES = 4
	TR_LAYOUTCOLORS = 8
	TR_USERFORM = 16
	TR_MOUSEMOVE = 32
	TR_CREATE = 64
	TR_DISTORYALL = 128
	TR_OVERRIDEDEFAULT = 256

	TR_ALL = -1
End Enum

Private Const CustomButtonID As String = "#65435"    'unique number used by CustomButton controls

Private Const CB_ParentName = "Unknown"
Private Const CB_VarName As String = "Unknown"
Private Const CB_Accelerator As String = ""
Private Const CB_ExcelGreen As Boolean = False
Private Const CB_FontName As String = "Tahoma"  '"Segoe UI"
Private Const CB_FontSize As Single = 8
Private Const CB_FontBold As Boolean = False
Private Const CB_FontItalic As Boolean = False
Private Const CB_FontUnderline As Boolean = False
Private Const CB_FontStrikethrough As Boolean = False

Private Const CB_Height As Single = 20
Private Const CB_Width As Single = 68

Private Const CB_Tag As String = ""
Private Const CB_TabIndex As Integer = -1
Private Const CB_TabStop As Boolean = False     'must be false
Private Const CB_Cancel As Boolean = False
Private Const CB_Default As Boolean = False
Private Const CB_Visible As Boolean = True
Private Const CB_Enabled As Boolean = True

'######################## Button Methods ##########################

'**********************************
'
'      Create a custom button
'
'**********************************

Public Function Create(ByRef Parent As Object, VariableName As String, Name As String, Caption As String, Top As Single, Left As Single, _
					   Optional Width As Variant, _
					   Optional Height As Variant, _
					   Optional FontName As Variant, _
					   Optional FontSize As Variant, _
					   Optional FontBold As Variant, _
					   Optional FontItalic As Variant, _
					   Optional FontUnderline As Variant, _
					   Optional FontStrikethrough As Variant, _
					   Optional Accelerator As Variant, _
					   Optional Cancel As Variant, _
					   Optional TabIndex As Variant, _
					   Optional TabStop As Variant, _
					   Optional Default As Variant, _
					   Optional Visible As Variant, _
					   Optional Enabled As Variant) As CustomButton
	Dim TempTrace As Long

	If CommandButtonExists(Parent, Name) Then
		Err.Raise vbObjectError + 1008, "CustomButton", "Creating a button with a duplicate name of '" & Name & "' is not permtted."
		Exit Function
	End If

	TempTrace = TraceCustomButtons
	If (TempTrace And TR_CREATE) <> TR_CREATE Then TraceCustomButtons = 0

	Set Create = New CustomButton

	ReDim Preserve ButtonList(ButtonIdx)
	Set ButtonList(ButtonIdx) = Create

	ReDim Preserve ButtonUsed(ButtonIdx)
	ButtonUsed(ButtonIdx) = True

	ButtonIdx = ButtonIdx + 1

	With Create    'call each class property/function
		.Parent = Parent
		.VarName = VariableName
		.Name = Name
		.Caption = Caption
		.Top = Top
		.Left = Left

		If Not IsMissing(Height) Then .Height = Height
		If Not IsMissing(Width) Then .Width = Width

		.ExcelGreen = ExcelGreenFlag

		If Not IsMissing(Accelerator) Then .Accelerator = Accelerator
		If Not IsMissing(TabIndex) Then .TabIndex = TabIndex
		If Not IsMissing(TabStop) Then .TabStop = TabStop
		If Not IsMissing(Cancel) Then .Cancel = Cancel
		If Not IsMissing(Default) Then .Default = Default
		If Not IsMissing(Visible) Then .Visible = Visible
		If Not IsMissing(Enabled) Then .Enabled = Enabled

		If Not IsMissing(FontName) Then .Font.Name = FontName
		If Not IsMissing(FontSize) Then .Font.SIZE = FontSize
		If Not IsMissing(FontBold) Then .Font.Bold = FontBold
		If Not IsMissing(FontItalic) Then .Font.Italic = FontItalic
		If Not IsMissing(FontUnderline) Then .Font.Underline = FontUnderline
		If Not IsMissing(FontStrikethrough) Then .Font.Strikethrough = FontStrikethrough

		.Initialize
	End With

	Create.Register     'connect to event catcher

	If (TempTrace And TR_CREATE) <> TR_CREATE Then TraceCustomButtons = TempTrace
End Function

Private Function CommandButtonExists(ByRef Parent As Object, Name As String) As Boolean
	Dim Ctrl As Variant

	CommandButtonExists = False

	On Error Resume Next
	For Each Ctrl In Parent.Controls
		If TypeName(Ctrl) = "CommandButton" And Ctrl.Name = Name Then
			CommandButtonExists = True
			Exit For
		End If
	Next Ctrl
	On Error GoTo 0
End Function

'**********************************
'
'    Distroy all custom buttons
'
'**********************************

Public Sub DestroyAll(ByRef Parent As Object)
	Dim i As Long, TempTrace As Long
	Dim Done As Boolean

	TempTrace = TraceCustomButtons
	If (TempTrace And TR_DISTORYALL) <> TR_DISTORYALL Then TraceCustomButtons = 0

	Done = True
	If Not IsEmptyArr(ButtonList) Then
		For i = 0 To UBound(ButtonList)
			If ButtonUsed(i) Then
				If ButtonList(i).Parent.Name = Parent.Name Then
						ButtonList(i).Unregister
						ButtonUsed(i) = False
						Set ButtonList(i) = Nothing
				Else
					Done = False
				End If
			End If
		Next i

		DoEvents
	End If

	If Done Then
		Erase ButtonList
		Erase ButtonUsed
		ButtonIdx = 0
	End If

	If (TempTrace And TR_DISTORYALL) <> TR_DISTORYALL Then TraceCustomButtons = TempTrace
End Sub

'**********************************
'
'       Output userform code
'
'**********************************

Public Sub ShowUserFormCode(ByRef UserForm As Object)
	Dim Params As String, VarName As String, Buttons As String, Events As String
	Dim i As Long, OptCnt As Long
	Dim Ctrl As Variant

	On Error Resume Next

	Debug.Print "Option Explicit"

	Debug.Print
	For Each Ctrl In UserForm.Controls
		If TypeName(Ctrl) = "CommandButton" And InStr(1, Ctrl.Name, CustomButtonID) = 0 Then
			VarName = Ctrl.Name     'variable name

			Debug.Print "Private WithEvents " & VarName & " As CustomButton"
		End If
	Next Ctrl

	Debug.Print
	Debug.Print "'--------- START OF USERFORM CONTROL EVENTS ---------"

	Events = ""
	For Each Ctrl In UserForm.Controls
		If TypeName(Ctrl) = "CommandButton" And InStr(1, Ctrl.Name, CustomButtonID) = 0 Then
			VarName = Ctrl.Name     'variable name

			Events = "End Sub" & vbLf & Events
			Events = "    MsgBox ""Clicked on "" & " & VarName & ".Caption, vbInformation, ""Message""" & vbLf & Events
			Events = "Private Sub " & VarName & "_Click()" & vbLf & Events
			Events = vbLf & Events
		End If
	Next Ctrl
	Debug.Print Events

	Debug.Print "'--------- END OF USERFORM CONTROL EVENTS ---------"

	Debug.Print
	Debug.Print "Private Sub UserForm_Activate()"

	Buttons = ""
	For Each Ctrl In UserForm.Controls
		If TypeName(Ctrl) = "CommandButton" And InStr(1, Ctrl.Name, CustomButtonID) = 0 Then
			Params = ""

			VarName = Ctrl.Name     'variable name

			Params = """" & VarName & """, " & _
					 """" & Ctrl.Name & """, " & _
					 """" & Ctrl.Caption & """, " & _
					 Ctrl.Top & ", " & Ctrl.Left

			'need to maintain actual order with optional params for create()

			OptCnt = 1
			If Ctrl.Width <> Def_Width Then
				Params = Params & ", " & Ctrl.Width
				OptCnt = OptCnt + 1
			End If

			If Ctrl.Height <> Def_Height Then
				If OptCnt = 2 Then
					Params = Params & ", " & Ctrl.Height
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", Height:=" & Ctrl.Height
				End If
			End If

			If Ctrl.Font.Name <> Def_FontName Then
				If OptCnt = 3 Then
					Params = Params & ", """ & Ctrl.Font.Name & """"
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", FontName:=""" & Ctrl.Font.Name & """"
				End If
			End If

			If Ctrl.Font.SIZE <> Def_FontSize Then
				If OptCnt = 4 Then
					Params = Params & ", " & Ctrl.Font.SIZE
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", FontSize:=" & Ctrl.Font.SIZE
				End If
			End If

			If Ctrl.Font.Bold <> Def_FontBold Then
				If OptCnt = 5 Then
					Params = Params & ", " & Ctrl.Font.Bold
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", FontBold:=" & Ctrl.Font.Bold
				End If
			End If

			If Ctrl.Font.Italic <> Def_FontItalic Then
				If OptCnt = 6 Then
					Params = Params & ", " & Ctrl.Font.Italic
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", FontItalic:=" & Ctrl.Font.Italic
				End If
			End If

			If Ctrl.Font.Underline <> Def_FontUnderline Then
				If OptCnt = 7 Then
					Params = Params & ", " & Ctrl.Font.Underline
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", FontUnderline:=" & Ctrl.Font.Underline
				End If
			End If

			If Ctrl.Font.Strikethrough <> Def_FontStrikethrough Then
				If OptCnt = 8 Then
					Params = Params & ", " & Ctrl.Font.Strikethrough
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", FontStrikethrough:=" & Ctrl.Font.Strikethrough
				End If
			End If

			If Ctrl.Accelerator <> Def_Accelerator Then
				If OptCnt = 9 Then
					Params = Params & ", """ & Ctrl.Accelerator & """"
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", Accelerator:=""" & Ctrl.Accelerator & """"
				End If
			End If

			If Ctrl.Cancel <> Def_Cancel Then
				If OptCnt = 10 Then
					Params = Params & ", " & Ctrl.Cancel
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", Cancel:=" & Ctrl.Cancel
				End If
			End If

			If Ctrl.TabStop <> Def_TabStop Then
				If OptCnt = 12 Then
					Params = Params & ", " & Ctrl.TabStop
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", TabStop:=" & Ctrl.TabStop
				End If
			End If

			If Ctrl.TabStop = True Then
				If OptCnt = 11 Then
					Params = Params & ", " & Ctrl.TabIndex
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", TabIndex:=" & Ctrl.TabIndex
				End If
			End If

			If Ctrl.Default <> Def_Default Then
				If OptCnt = 13 Then
					Params = Params & ", " & Ctrl.Default
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", Default:=" & Ctrl.Default
				End If
			End If

			If Ctrl.Visible <> Def_Visible Then
				If OptCnt = 14 Then
					Params = Params & ", " & Ctrl.Visible
					OptCnt = OptCnt + 1
				Else
					Params = Params & ", Visible:=" & Ctrl.Visible
				End If
			End If

			If Ctrl.Enabled <> Def_Enabled Then
				 If OptCnt = 15 Then
					Params = Params & ", " & Ctrl.Enabled
				Else
				   Params = Params & ", Enabled:=" & Ctrl.Enabled
				End If
			End If

			Buttons = vbTab & "Set " & VarName & " = CustomButtons.Create(Me, " & Params & ")" & vbLf & Buttons
		End If
	Next Ctrl

	Debug.Print Buttons & "End Sub"
	Debug.Print ""
	Debug.Print "Private Sub UserForm_Initialize()"
	Debug.Print "    Dim hWnd As LongPtr"
	Debug.Print ""
	Debug.Print "    With " & UserForm.Name
	Debug.Print "        .Height = " & Format(UserForm.Height, "#0.00") & "        'fix screen driver issue"
	Debug.Print "        .Width = " & Format(UserForm.Width, "#0.00") & ""
	Debug.Print ""
	Debug.Print "        .StartUpPosition = 0    'allow manual positioning"
	Debug.Print ""
	Debug.Print "        hWnd = FindWindow(""ThunderDFrame"", .Caption)"
	Debug.Print "        CentreUserForm hWnd"
	Debug.Print "    End With"
	Debug.Print "End Sub"
	Debug.Print ""
	Debug.Print "Private Sub UserForm_Terminate()"
	Debug.Print "    CustomButtons.DestroyAll Me     'finished with all custombuttons"
	Debug.Print "End Sub"
	Debug.Print

	On Error GoTo 0
End Sub

Friend Sub OverrideDefault(Value As Boolean, Name As String, ByRef Parent As Object)
	Dim i As Long, TempTrace As Long

	TempTrace = TraceCustomButtons
	If (TempTrace And TR_OVERRIDEDEFAULT) <> TR_OVERRIDEDEFAULT Then TraceCustomButtons = 0

	For i = 0 To ButtonIdx - 1
		If ButtonUsed(i) Then
			If ButtonList(i).Parent.Name = Parent.Name And ButtonList(i).Name <> Name And ButtonList(i).Default Then
				ButtonList(i).DefaultOverride = Value
				Exit For
			End If
		End If

		DoEvents
	Next i

	If (TempTrace And TR_OVERRIDEDEFAULT) <> TR_OVERRIDEDEFAULT Then TraceCustomButtons = TempTrace
End Sub

Friend Sub Reinitialize(ByRef Parent As Object)
	Dim i As Long, TempTrace As Long

	TempTrace = TraceCustomButtons
	TraceCustomButtons = 0      'switch off tracing

	For i = 0 To ButtonIdx - 1
		If ButtonUsed(i) Then
			If ButtonList(i).Parent.Name = Parent.Name Then
				ButtonList(i).Reinitialize
			End If
		End If

		DoEvents
	Next i

	TraceCustomButtons = TempTrace
End Sub

'################### Get/Let Button Properties ####################

'State
Friend Property Get state() As Integer
	state = MouseButtonState
End Property

Friend Property Let state(Value As Integer)
	MouseButtonState = Value
End Property

'Trace
Public Property Get Trace() As Long
	Trace = TraceCustomButtons
End Property

Public Property Let Trace(Value As Long)
	TraceCustomButtons = Value
End Property

'ExcelGreen
Public Property Get ExcelGreen() As Long
	ExcelGreen = ExcelGreenFlag
End Property

Public Property Let ExcelGreen(Value As Long)
	Dim i As Long, TempTrace As Long

	ExcelGreenFlag = Value

	TempTrace = TraceCustomButtons
	TraceCustomButtons = 0      'switch off tracing

	For i = 0 To ButtonIdx - 1
		If ButtonUsed(i) Then
			ButtonList(i).ExcelGreen = ExcelGreenFlag
		End If

		DoEvents
	Next i

	TraceCustomButtons = TempTrace
End Property

'LastName
Friend Property Get LastName() As String
	LastName = TraceLastName
End Property

Friend Property Let LastName(Value As String)
	TraceLastName = Value
End Property

'LastParentName
Friend Property Get LastParentName() As String
	LastParentName = TraceLastParentName
End Property

Friend Property Let LastParentName(Value As String)
	TraceLastParentName = Value
End Property

'ButtonCount
Friend Property Get ButtonCount() As Long
	ButtonCount = ButtonIdx
End Property

'ID
Friend Property Get ID() As String
	ID = CustomButtonID
End Property


'Def_ParentName
Friend Property Get Def_ParentName() As String
	Def_ParentName = CB_ParentName
End Property

'Def_VarName
Friend Property Get Def_VarName() As String
	Def_VarName = CB_VarName
End Property

'Def_Accelerator
Friend Property Get Def_Accelerator() As String
	Def_Accelerator = CB_Accelerator
End Property

'Def_ExcelGreen
Friend Property Get Def_ExcelGreen() As Boolean
	Def_ExcelGreen = CB_ExcelGreen
End Property

'Def_FontName
Friend Property Get Def_FontName() As String
	Def_FontName = CB_FontName
End Property

'Def_FontSize
Friend Property Get Def_FontSize() As Single
	Def_FontSize = CB_FontSize
End Property

'Def_FontBold
Friend Property Get Def_FontBold() As Boolean
	Def_FontBold = CB_FontBold
End Property

'Def_FontItalic
Friend Property Get Def_FontItalic() As Boolean
	Def_FontItalic = CB_FontItalic
End Property

'Def_FontUnderline
Friend Property Get Def_FontUnderline() As Boolean
	Def_FontUnderline = CB_FontUnderline
End Property

'Def_FontStrikethrough
Friend Property Get Def_FontStrikethrough() As Boolean
	Def_FontStrikethrough = CB_FontStrikethrough
End Property

'Def_Height
Friend Property Get Def_Height() As Single
	Def_Height = CB_Height
End Property

'Def_Width
Friend Property Get Def_Width() As Single
	Def_Width = CB_Width
End Property

'Def_Tag
Friend Property Get Def_Tag() As String
	Def_Tag = CB_Tag
End Property

'Def_TabIndex
Friend Property Get Def_TabIndex() As Integer
	Def_TabIndex = CB_TabIndex
End Property

'Def_TabStop
Friend Property Get Def_TabStop() As Boolean
	Def_TabStop = CB_TabStop
End Property

'Def_Cancel
Friend Property Get Def_Cancel() As Boolean
	Def_Cancel = CB_Cancel
End Property

'Def_Default
Friend Property Get Def_Default() As Boolean
	Def_Default = CB_Default
End Property

'Def_Visible
Friend Property Get Def_Visible() As Boolean
	Def_Visible = CB_Visible
End Property

'Def_Enabled
Friend Property Get Def_Enabled() As Boolean
	Def_Enabled = CB_Enabled
End Property
