VERSION 5.00
Begin VB.Form frmFavDel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete Favorite"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUrl 
      Appearance      =   0  'Ïëîñêà
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5055
   End
   Begin MdeX.xpcmdbutton cmdChURL 
      Height          =   330
      Left            =   1575
      TabIndex        =   4
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "Change URL"
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
   Begin MdeX.xpcmdbutton cmdRename 
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   525
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "Rename"
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
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3150
      Top             =   105
   End
   Begin MdeX.xpcmdbutton cmdClose 
      Height          =   330
      Left            =   3780
      TabIndex        =   2
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      Caption         =   "Close"
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
   Begin MdeX.xpcmdbutton cmdDel 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "Delete"
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
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   105
      TabIndex        =   0
      Top             =   1365
      Width           =   5055
   End
End
Attribute VB_Name = "frmFavDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChURL_Click()
    frmFavDelCh.Visible = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub cmdDel_Click()
'Del_Fav(index)
txtUrl.Text = ""
Del_Fav List1.ListIndex
Add_url List1.ListIndex, Get_Url(frmWeb.mnuItemSite.Count - 1)
Add_url frmWeb.mnuItemSite.Count - 1, ""
frmWeb.mnuItemSite(List1.ListIndex).Caption = frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count - 1).Caption
Unload frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count - 1)
Form_Load
End Sub

Private Sub cmdRename_Click()
Dim tName As String
x = InputBox("Enter name:", "Favorites")
tName = x
If x <> "" Then
    frmWeb.mnuItemSite(List1.ListIndex).Caption = x
    Del_Fav frmFavDel.List1.ListIndex
    Add_Fav tName, Get_Url(frmFavDel.List1.ListIndex), frmFavDel.List1.ListIndex
End If


Form_Load
End Sub


Private Sub Form_Load()
cmdDel.Enabled = False
cmdRename.Enabled = False
cmdChURL.Enabled = False
List1.Clear
    For i = 0 To frmWeb.mnuItemSite.Count - 1
        List1.AddItem frmWeb.mnuItemSite(i).Caption
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmWeb.Enabled = True
End Sub


Private Sub List1_Click()
txtUrl.Text = Get_Url(List1.ListIndex)
If List1.ListIndex <> 0 Then cmdDel.Enabled = True Else cmdDel.Enabled = False
If List1.ListIndex <> 0 Then cmdRename.Enabled = True Else cmdRename.Enabled = False
If List1.ListIndex <> 0 Then cmdChURL.Enabled = True Else cmdChURL.Enabled = False

End Sub


'Razbor
'    x = GetSetting("MdeX", "Favorites", j)
'    A = InStr(1, x, "&")
'    x1 = Left(x, A - 1)
'    x2 = Right(x, Len(x) - A)

