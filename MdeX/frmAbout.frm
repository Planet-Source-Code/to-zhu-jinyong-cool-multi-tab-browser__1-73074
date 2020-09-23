VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   7155
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   540
      Left            =   4095
      MousePointer    =   10  'Up Arrow
      Top             =   6510
      Width           =   1800
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    frmWeb.Enabled = True
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub

