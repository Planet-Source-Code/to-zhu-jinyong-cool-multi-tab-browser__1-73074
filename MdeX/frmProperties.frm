VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   6105
   ClientLeft      =   4125
   ClientTop       =   2355
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Ïèêñåëü
   ScaleWidth      =   361
   Begin VB.Frame Frame1 
      BackColor       =   &H00DA9F18&
      Caption         =   "Properties"
      Height          =   5475
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   5160
      Begin MdeX.xpgroupbox xpgroupbox1 
         Height          =   5475
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   5160
         _ExtentX        =   5953
         _ExtentY        =   3916
         Caption         =   "Properties"
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
         Begin MdeX.xpgroupbox xpgroupbox3 
            Height          =   1065
            Left            =   105
            TabIndex        =   11
            Top             =   315
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   1879
            Caption         =   "History"
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
            Begin MdeX.xpcmdbutton Command1 
               Height          =   330
               Left            =   3465
               TabIndex        =   12
               Top             =   630
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   582
               Caption         =   "Clear History"
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Ïðîçðà÷íî
               Caption         =   "All the addresses you entered were saved. If you want to clear history then press that button."
               Height          =   645
               Left            =   210
               TabIndex        =   13
               Top             =   315
               Width           =   3270
            End
         End
         Begin MdeX.xpgroupbox xpgroupbox2 
            Height          =   1380
            Left            =   105
            TabIndex        =   5
            Top             =   1470
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   2434
            Caption         =   "Home Page"
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
            Begin VB.TextBox txtAddress 
               Appearance      =   0  'Ïëîñêà
               Height          =   285
               Left            =   1080
               TabIndex        =   14
               Top             =   600
               Width           =   3615
            End
            Begin MdeX.xpcmdbutton command4 
               Height          =   330
               Left            =   1050
               TabIndex        =   6
               Top             =   945
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   582
               Caption         =   "Use default"
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
            Begin MdeX.xpcmdbutton command3 
               Height          =   330
               Left            =   2310
               TabIndex        =   7
               Top             =   945
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   582
               Caption         =   "Use current"
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
            Begin MdeX.xpcmdbutton command2 
               Height          =   330
               Left            =   3570
               TabIndex        =   8
               Top             =   945
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   582
               Caption         =   "Use blank"
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
            Begin VB.Label Label2 
               BackStyle       =   0  'Ïðîçðà÷íî
               Caption         =   "You can change your start page."
               Height          =   225
               Left            =   105
               TabIndex        =   10
               Top             =   315
               Width           =   4740
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Ïðîçðà÷íî
               Caption         =   "Address:"
               Height          =   225
               Left            =   105
               TabIndex        =   9
               Top             =   630
               Width           =   960
            End
         End
      End
   End
   Begin MdeX.xpcmdbutton cmdCancel 
      Height          =   330
      Left            =   3990
      TabIndex        =   2
      Top             =   5670
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
   Begin MdeX.xpcmdbutton cmdApply 
      Height          =   330
      Left            =   1470
      TabIndex        =   1
      Top             =   5670
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "Apply"
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
   Begin MdeX.xpcmdbutton cmdOk 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   5670
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "Ok"
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
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
Apply
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub cmdOk_Click()
    Apply
    Unload Me
End Sub

Private Sub Command1_Click()
x = MsgBox("Are you sure you want to delete your History?", vbYesNo, "History")
    If x = vbYes Then
    Open "c:\WebHis.dat" For Output As #1
    Close #1
    frmWeb.Combo1.Clear
    End If
End Sub

Private Sub Command2_Click()
    txtAddress.Text = "about:blank"
End Sub

Private Sub Command3_Click()
    txtAddress.Text = frmWeb.Web(frmWeb.TabStrip1.Value).LocationURL
End Sub

Private Sub Command4_Click()
    txtAddress.Text = "www.google.com"
End Sub

Public Sub Apply()
    SetHomePage txtAddress.Text
End Sub

Private Sub Form_Load()
    txtAddress.Text = GetHomePage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmWeb.Enabled = True
End Sub
