VERSION 5.00
Begin VB.Form frmFirst 
   Caption         =   "ExecuteApp"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1395
   ScaleWidth      =   4605
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbApps 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "&App Name"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbApps_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    ExecuteStuff
  End If
End Sub

Private Sub cmdExecute_Click()
  ExecuteStuff
End Sub

Private Sub Form_Activate()
  FormActivateOnce Me, False
End Sub

Sub FormStart()
Dim strOldCaption As String
  strOldCaption = Me.Caption
  
  FillCombo mconnSettingsDB, "tblApp", "Name", Me.cmbApps
  Me.cmbApps.ListIndex = 0
  DropCombo Me.cmbApps
  
  Me.Caption = strOldCaption
End Sub
Private Sub Form_Unload(Cancel As Integer)
  FormActivateOnce Me, True
End Sub
Sub ExecuteStuff()
Dim strAppName As String
Dim strOldSelfCaption As String
Dim strOldMainCaption As String

  strOldSelfCaption = Me.Caption
  strOldMainCaption = frmMain.Caption
  
  strAppName = Me.cmbApps.Text
  Me.Caption = "Executing " & strAppName
  ExecuteApp strAppName
  
  Me.Caption = strOldSelfCaption
  frmMain.Caption = strOldMainCaption

End Sub
