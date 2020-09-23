VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStart 
   BorderStyle     =   0  'Íåò
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   3270
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar LunaBar1 
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   945
      Top             =   1575
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
Dim hRgn As Long
hRgn = CreateRoundRectRgn(1, 1, Me.Width / 15, Me.Height / 15, 650, 250)
SetWindowRgn Me.hWnd, hRgn, True

i = True

    Add_url 0, "http://www.google.com/"
    Insert_Fav
    GetHistory
    frmWeb.TabStrip1.Tabs.Remove 1
    frmWeb.Web(frmWeb.TabStrip1.Value).Navigate2 GetHomePage
    frmWeb.TabStrip1.Tabs(0).Caption = GetHomePage
    frmWeb.Combo1.Text = GetHomePage
    LunaBar1.Max = 101
    LunaBar1.Value = 1

x = 1
End Sub

Private Sub Timer1_Timer()
x = x + 1
frmStart.LunaBar1.Value = frmStart.LunaBar1.Value + 1
If x = 101 Then
    frmWeb.Visible = True
    Unload Me
End If
End Sub
