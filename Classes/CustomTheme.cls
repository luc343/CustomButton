VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------
'
'                Class CustomTheme
'
' Note: Class is Private
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
'---------------------------------------------------

'GetThemeColor() and GetThemeSysColor() APIs do not return the required button colors for Windows 10 & 11

Private Declare PtrSafe Function RtlGetVersion Lib "ntdll.dll" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

'RtlGetVersion
Private Type OSVERSIONINFOEX
	dwOSVersionInfoSize As Long
	dwMajorVersion As Long
	dwMinorVersion As Long
	dwBuildNumber As Long
	dwPlatformId As Long
	szCSDVersion As String * 128
End Type

'Windows 10
Private Const Def_10_BackColor = 14803425               'RGB(225, 225, 225)
Private Const Def_10_BorderColor = 11382189             'RGB(173, 173, 173)
Private Const Def_10_HoverBackColor = 16511461          'RGB(229, 241, 251)
Private Const Def_10_HoverDownColor = 16245964          'RGB(204, 228, 247)
Private Const Def_10_HoverBorderColor = 14120960        'RGB(0, 120, 215)
Private Const Def_10_DisabledBackColor = 13421772       'RGB(204, 204, 204)
Private Const Def_10_DisabledBorderColor = 12566463     'RGB(191, 191, 191)
Private Const Def_10_DisabledForeColor = 8618883        'RGB(131, 131, 131)
Private Const Def_10_DefaultBorderColor = 14120960      'RGB(0, 120, 215)

'Windows 11
Private Const Def_11_BackColor = 16514043               'RGB(251, 251, 251)
Private Const Def_11_BorderColor = 15066597             'RGB(229, 229, 229)
Private Const Def_11_HoverBackColor = 16379616          'RGB(224, 238, 249)
Private Const Def_11_HoverDownColor = 16245964          'RGB(204, 228, 247)
Private Const Def_11_HoverBorderColor = 13924352        'RGB(0, 120, 212)
Private Const Def_11_DisabledBackColor = 16382457       'RGB(249, 249, 249)
Private Const Def_11_DisabledBorderColor = 15329769     'RGB(233, 233, 233)
Private Const Def_11_DisabledForeColor = 10395294       'RGB(158, 158, 158)
Private Const Def_11_DefaultBorderColor = 12609280      'RGB(0, 103, 192)

'Other Windows
Private Const Def_0_ForeColor = 0                       'RGB(0, 0, 0)
Private Const Def_0_BackColor = 15790320                'RGB(240, 240, 240)
Private Const Def_0_BorderColor = 11382189              'RGB(173, 173, 173)

Private Const Def_0_HoverForeColor = Def_0_ForeColor
Private Const Def_0_HoverBackColor = Def_0_BackColor
Private Const Def_0_HoverDownColor = Def_0_BackColor
Private Const Def_0_HoverBorderColor = Def_0_BorderColor

Private Const Def_0_DisabledBackColor = Def_0_BackColor
Private Const Def_0_DisabledBorderColor = Def_0_BorderColor
Private Const Def_0_DisabledForeColor = 10526880        'RGB(160, 160, 160)
Private Const Def_0_DefaultBorderColor = Def_0_BorderColor

'Excel Green
Private Const Def_Gr_HoverBackColor = 15660521          'RGB(233, 245, 238)
Private Const Def_Gr_HoverDownColor = 11783583          'RGB(159, 205, 179)
Private Const Def_Gr_HoverBorderColor = 6854222         'RGB(78, 150, 104)
Private Const Def_Gr_DefaultBorderColor = 6854222       'RGB(78, 150, 104)

Private Col_ForeColor As Long
Private Col_BackColor As Long
Private Col_BorderColor As Long

Private Col_HoverForeColor As Long
Private Col_HoverBackColor As Long
Private Col_HoverDownColor As Long
Private Col_HoverBorderColor As Long

Private Col_DisabledBackColor As Long
Private Col_DisabledBorderColor As Long
Private Col_DisabledForeColor As Long
Private Col_DefaultBorderColor As Long

Private ExcelGreenFlag As Boolean               'Custom Excel color flag
Private InitFlag As Boolean                     'Initalization flag

'######################## Color Methods ##########################

Public Sub Init()
	If InitFlag Then Exit Sub

	ExcelGreenFlag = CustomButtons.Def_ExcelGreen

	Col_ForeColor = Def_0_ForeColor
	Col_HoverForeColor = Def_0_HoverForeColor

	Select Case GetWindowsVersion()
		Case 10     'Windows 10
			Col_BackColor = Def_10_BackColor
			Col_BorderColor = Def_10_BorderColor

			Col_HoverBackColor = Def_10_HoverBackColor
			Col_HoverDownColor = Def_10_HoverDownColor
			Col_HoverBorderColor = Def_10_HoverBorderColor

			Col_DisabledBackColor = Def_10_DisabledBackColor
			Col_DisabledBorderColor = Def_10_DisabledBorderColor
			Col_DisabledForeColor = Def_10_DisabledForeColor

			Col_DefaultBorderColor = Def_10_DefaultBorderColor

		Case 11     'Windows 11
			Col_BackColor = Def_11_BackColor
			Col_BorderColor = Def_11_BorderColor

			Col_HoverBackColor = Def_11_HoverBackColor
			Col_HoverDownColor = Def_11_HoverDownColor
			Col_HoverBorderColor = Def_11_HoverBorderColor

			Col_DisabledBackColor = Def_11_DisabledBackColor
			Col_DisabledBorderColor = Def_11_DisabledBorderColor
			Col_DisabledForeColor = Def_11_DisabledForeColor

			Col_DefaultBorderColor = Def_11_DefaultBorderColor

		Case Else   'other Windows
			Col_BackColor = Def_0_BackColor
			Col_BorderColor = Def_0_BorderColor

			Col_HoverBackColor = Def_0_HoverBackColor
			Col_HoverDownColor = Def_0_HoverDownColor
			Col_HoverBorderColor = Def_0_HoverBorderColor

			Col_DisabledBackColor = Def_0_DisabledBackColor
			Col_DisabledBorderColor = Def_0_DisabledBorderColor
			Col_DisabledForeColor = Def_0_DisabledForeColor

			Col_DefaultBorderColor = Def_0_DefaultBorderColor
	End Select

	InitFlag = True
End Sub

Private Function GetWindowsVersion() As Long
	Dim osvi As OSVERSIONINFOEX

	osvi.dwOSVersionInfoSize = Len(osvi)
	If RtlGetVersion(osvi) = 0 Then
		If osvi.dwMajorVersion = 10 Then
			If osvi.dwBuildNumber >= 22000 Then
				GetWindowsVersion = 11       'windows 11
			Else
				GetWindowsVersion = 10       'windows 10
			End If
		Else
			GetWindowsVersion = 0            'all others
		End If
	Else
		GetWindowsVersion = 0    'default
	End If
End Function

Private Sub SplitRGB(ByVal RGBColor As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
	r = RGBColor And 255
	g = (RGBColor \ 256) And 255
	b = (RGBColor \ 65536) And 255
End Sub

Private Sub DebugPrint(Text As String, RGBColor As Long)
	Dim r As Long, g As Long, b As Long

	SplitRGB RGBColor, r, g, b

	Debug.Print Text & ": RGB(" & r & ", " & g & ", " & b & ")"
End Sub

'#################### Get Color Properties #######################

'ExcelGreen
Friend Property Get ExcelGreen() As Boolean
	ExcelGreen = ExcelGreenFlag
End Property

Friend Property Let ExcelGreen(Value As Boolean)
	ExcelGreenFlag = Value

	If Value = True Then
		'Excel
		Col_HoverBackColor = Def_Gr_HoverBackColor
		Col_HoverDownColor = Def_Gr_HoverDownColor
		Col_HoverBorderColor = Def_Gr_HoverBorderColor

		Col_DefaultBorderColor = Def_Gr_DefaultBorderColor
	Else
		Select Case GetWindowsVersion()
			Case 10     'Windows 10
				Col_HoverBackColor = Def_10_HoverBackColor
				Col_HoverDownColor = Def_10_HoverDownColor
				Col_HoverBorderColor = Def_10_HoverBorderColor

				Col_DefaultBorderColor = Def_10_DefaultBorderColor

			Case 11     'Windows 11
				Col_HoverBackColor = Def_11_HoverBackColor
				Col_HoverDownColor = Def_11_HoverDownColor
				Col_HoverBorderColor = Def_11_HoverBorderColor

				Col_DefaultBorderColor = Def_11_DefaultBorderColor

			Case Else   'all other windows versions
				Col_HoverBackColor = Def_0_HoverBackColor
				Col_HoverDownColor = Def_0_HoverDownColor
				Col_HoverBorderColor = Def_0_HoverBorderColor

				Col_DefaultBorderColor = Def_0_DefaultBorderColor
		End Select
	End If
End Property

'BackColor
Friend Property Get BackColor() As Long
	BackColor = Col_BackColor
End Property

Friend Property Let BackColor(Value As Long)
	Col_BackColor = Value
End Property

'HoverBackColor
Friend Property Get HoverBackColor() As Long
	HoverBackColor = Col_HoverBackColor
End Property

Friend Property Let HoverBackColor(Value As Long)
	Col_HoverBackColor = Value
End Property

'HoverBorderColor
Friend Property Get HoverBorderColor() As Long
	HoverBorderColor = Col_HoverBorderColor
End Property

Friend Property Let HoverBorderColor(Value As Long)
	Col_HoverBorderColor = Value
End Property

'HoverDownColor
Friend Property Get HoverDownColor() As Long
	HoverDownColor = Col_HoverDownColor
End Property

Friend Property Let HoverDownColor(Value As Long)
	Col_HoverDownColor = Value
End Property

'DisabledBackColor
Friend Property Get DisabledBackColor() As Long
	DisabledBackColor = Col_DisabledBackColor
End Property

Friend Property Let DisabledBackColor(Value As Long)
	Col_DisabledBackColor = Value
End Property

'DisabledForeColor
Friend Property Get DisabledForeColor() As Long
	DisabledForeColor = Col_DisabledForeColor
End Property

Friend Property Let DisabledForeColor(Value As Long)
	Col_DisabledForeColor = Value
End Property

'DisabledBorderColor
Friend Property Get DisabledBorderColor() As Long
	DisabledBorderColor = Col_DisabledBorderColor
End Property

Friend Property Let DisabledBorderColor(Value As Long)
	Col_DisabledBorderColor = Value
End Property

'DefaultBorderColor
Friend Property Get DefaultBorderColor() As Long
	DefaultBorderColor = Col_DefaultBorderColor
End Property

Friend Property Let DefaultBorderColor(Value As Long)
	Col_HoverBorderColor = Value
End Property

'ForeColor
Friend Property Get ForeColor() As Long
	ForeColor = Col_ForeColor
End Property

Friend Property Let ForeColor(Value As Long)
	Col_ForeColor = Value
End Property

'HoverForeColor
Friend Property Get HoverForeColor() As Long
	HoverForeColor = Col_HoverForeColor
End Property

Friend Property Let HoverForeColor(Value As Long)
	Col_HoverForeColor = Value
End Property

'BorderColor
Friend Property Get BorderColor() As Long
	BorderColor = Col_BorderColor
End Property

Friend Property Let BorderColor(Value As Long)
	Col_BorderColor = Value
End Property
