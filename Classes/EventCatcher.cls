VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventCatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------
'
'                Class EventCatcher
'
' Copyright (c) Lucien Cinc 2025
'
' Available under the MIT license: see the LICENSE
' file at the root of this project.
'
' Based on the work from the site:
'   http://addinbox.sakura.ne.jp/English_Page.htm
'
'---------------------------------------------------

Private Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long

Private Type GUID
	Data1 As Long
	Data2 As Integer
	Data3 As Integer
	Data4(0 To 7) As Byte
End Type

Private Cookie As Long              'returned by ConnectToConnectionPoint

Private MyCtrl As Object
Attribute MyCtrl.VB_VarHelpID = -1
Private MyParent As CustomButton    'reference to CustomButton object

'######################### Event Methods ##########################

Private Sub Class_Initialize()
	Set MyCtrl = Nothing
	Cookie = 0
End Sub

Public Sub Clear()
	If (Cookie <> 0) Then
		ConnectEvent False
		Cookie = 0
	End If

	Set MyCtrl = Nothing
End Sub

Private Function ConnectEvent(ByVal Connect As Boolean) As Long
	Dim IID_IDispatch As GUID

	With IID_IDispatch      'GUID {00020400-0000-0000-C000000000000046}
		.Data1 = &H20400
		.Data4(0) = &HC0
		.Data4(7) = &H46
	End With

	ConnectEvent = ConnectToConnectionPoint(Me, IID_IDispatch, Connect, MyCtrl, Cookie, 0&)
End Function

'######################## Listener Events #########################

Public Sub HookEnter()
Attribute HookEnter.VB_UserMemId = -2147384830
	MyParent.RaiseEnter
End Sub

Public Sub HookExit(ByVal Cancel As MSForms.ReturnBoolean)
Attribute HookExit.VB_UserMemId = -2147384829
	MyParent.RaiseExit Cancel
End Sub

Public Sub HookBeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Attribute HookBeforeUpdate.VB_UserMemId = -2147384831
	MyParent.RaiseBeforeUpdate Cancel
End Sub

Public Sub HookAfterUpdate()
Attribute HookAfterUpdate.VB_UserMemId = -2147384832
	MyParent.RaiseAfterUpdate
End Sub

'#################### Get/Let Event Properties ####################

Public Property Let Ctrl(NewCtrl As Object)
	Set MyCtrl = NewCtrl

	ConnectEvent True
End Property

Public Property Get Ctrl() As Object
	Set Ctrl = MyCtrl
End Property

Public Property Let Parent(NewParent As CustomButton)
	Set MyParent = NewParent
End Property

Public Property Get Parent() As CustomButton
	Set Parent = MyParent
End Property
