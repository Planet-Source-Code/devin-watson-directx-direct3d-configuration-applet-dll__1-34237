VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "D3DConfig Calling Program"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Applet"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents myConfig As D3DConfigForm
Attribute myConfig.VB_VarHelpID = -1

Private Sub cmdOpen_Click()
    myConfig.Show
End Sub

Private Sub Form_Load()
    Set myConfig = New D3DConfigForm
    myConfig.Caption = "DirectX Configuration"
End Sub


Private Sub myConfig_OnCloseClick()
    MsgBox "The user pressed close.", vbOKOnly + vbInformation, "Close Pressed"
End Sub


Private Sub myConfig_OnOKClick(AdapterSelected As Long, ModeSelected As Long, Fullscreen As Boolean)
    MsgBox "Adapter Selected: " & AdapterSelected & vbCrLf & "Mode Selected: " & ModeSelected & _
            vbCrLf & "Fullscreen: " & Fullscreen
End Sub


