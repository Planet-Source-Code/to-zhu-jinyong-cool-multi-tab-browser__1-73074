VERSION 5.00
Begin VB.Form frmFavDelCh 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change URL"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MdeX.xpcmdbutton cmdCancel 
      Height          =   330
      Left            =   5040
      TabIndex        =   3
      Top             =   1575
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MdeX.xpcmdbutton cmdChange 
      Height          =   330
      Left            =   3675
      TabIndex        =   2
      Top             =   1575
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "Change"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MdeX.xpgroupbox xpgroupbox2 
      Height          =   645
      Left            =   105
      TabIndex        =   1
      Top             =   840
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   1138
      Caption         =   "New Name:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin VB.TextBox txtNew 
         Appearance      =   0  'Ïëîñêà
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6015
      End
   End
   Begin MdeX.xpgroupbox xpgroupbox1 
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   1138
      Caption         =   "Old Name:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin VB.TextBox txtOld 
         Appearance      =   0  'Ïëîñêà
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmFavDelCh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdChange_Click()
Add_url frmFavDel.List1.ListIndex, txtNew.Text
Del_Fav frmFavDel.List1.ListIndex
Add_Fav frmFavDel.List1.List(frmFavDel.List1.ListIndex), txtNew.Text, frmFavDel.List1.ListIndex
Unload Me
End Sub

Private Sub Form_Load()
txtOld.Text = Get_Url(frmFavDel.List1.ListIndex)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFavDel.Enabled = True
End Sub

Private Sub txtNew_Change()
If txtNew.Text <> "" Then cmdChange.Enabled = True Else cmdChange.Enabled = False
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub

