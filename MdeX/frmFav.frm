VERSION 5.00
Begin VB.Form frmFav 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Favorites"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MdeX.xpgroupbox xpgroupbox2 
      Height          =   1065
      Left            =   105
      TabIndex        =   4
      Top             =   1155
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   1879
      Caption         =   "Address"
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
      Begin VB.TextBox txtURL 
         Appearance      =   0  'Ïëîñêà
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DA9F18&
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "URL:"
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Width           =   1275
      End
   End
   Begin MdeX.xpcmdbutton cmdReset 
      Height          =   330
      Left            =   6090
      TabIndex        =   3
      Top             =   1890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "Reset"
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
   Begin MdeX.xpgroupbox xpgroupbox1 
      Height          =   1065
      Left            =   105
      TabIndex        =   2
      Top             =   0
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   1879
      Caption         =   "Name"
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
      Begin VB.TextBox txtName 
         Appearance      =   0  'Ïëîñêà
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DA9F18&
         BackStyle       =   0  'Ïðîçðà÷íî
         Caption         =   "Name:"
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   420
         Width           =   960
      End
   End
   Begin MdeX.xpcmdbutton cmdCancel 
      Height          =   330
      Left            =   6090
      TabIndex        =   1
      Top             =   525
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
   Begin MdeX.xpcmdbutton cmdAdd 
      Height          =   330
      Left            =   6090
      TabIndex        =   0
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "Add"
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
End
Attribute VB_Name = "frmFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

    Load frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count)
    frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count - 1).Caption = txtName.Text
    frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count - 1).Visible = True
    Add_url frmWeb.mnuItemSite.Count - 1, txtUrl.Text
    Add_Fav txtName.Text, txtUrl.Text, frmWeb.mnuItemSite.Count - 1
    
    Unload Me
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
    txtName.Text = frmWeb.Web(frmWeb.TabStrip1.Value).LocationName
    txtUrl.Text = frmWeb.Web(frmWeb.TabStrip1.Value).LocationURL
End Sub

Private Sub Form_Load()
    txtName.Text = frmWeb.Web(frmWeb.TabStrip1.Value).LocationName
    txtUrl.Text = frmWeb.Web(frmWeb.TabStrip1.Value).LocationURL
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmWeb.Enabled = True
End Sub


