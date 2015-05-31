VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "General HTML Generator"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FormStart()
Dim strOldCaption As String


  strOldCaption = Me.Caption
  
  modCommon.Init
  frmFirst.Show
  
  Me.Caption = strOldCaption
  
End Sub

Private Sub MDIForm_Activate()
  FormActivateOnce Me, False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  FormActivateOnce Me, True
End Sub
