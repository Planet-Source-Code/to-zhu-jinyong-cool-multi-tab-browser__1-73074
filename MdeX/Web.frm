VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmWeb 
   Caption         =   "MustDie-Explorer"
   ClientHeight    =   6780
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Ïðèâÿçàòü ââåðõ
      Height          =   1065
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1879
      BandCount       =   2
      _CBWidth        =   9960
      _CBHeight       =   1065
      _Version        =   "6.0.8169"
      Child1          =   "Picture1"
      MinHeight1      =   645
      Width1          =   960
      NewRow1         =   0   'False
      Child2          =   "Picture2"
      MinHeight2      =   330
      Width2          =   1695
      NewRow2         =   -1  'True
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'Íåò
         Height          =   330
         Left            =   165
         ScaleHeight     =   330
         ScaleWidth      =   9705
         TabIndex        =   5
         Top             =   705
         Width           =   9705
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   735
            TabIndex        =   7
            Top             =   0
            Width           =   5895
         End
         Begin MdeX.xpcmdbutton cmdGo 
            Height          =   300
            Left            =   6720
            TabIndex        =   6
            Top             =   5
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   529
            Caption         =   "==>"
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
            Caption         =   "Address"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   50
            TabIndex        =   8
            Top             =   60
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Íåò
         Height          =   645
         Left            =   165
         ScaleHeight     =   645
         ScaleWidth      =   9705
         TabIndex        =   4
         Top             =   30
         Width           =   9705
         Begin VB.CommandButton newww 
            Caption         =   "New"
            Height          =   625
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   750
         End
         Begin VB.CommandButton cmdBack 
            Caption         =   "<<<"
            Height          =   625
            Left            =   840
            TabIndex        =   15
            Top             =   0
            Width           =   750
         End
         Begin VB.CommandButton cmdForward 
            Caption         =   ">>>"
            Height          =   625
            Left            =   1680
            TabIndex        =   14
            Top             =   0
            Width           =   750
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Height          =   625
            Left            =   2520
            TabIndex        =   13
            Top             =   0
            Width           =   750
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   625
            Left            =   3360
            TabIndex        =   12
            Top             =   0
            Width           =   750
         End
         Begin VB.CommandButton cmdHome 
            Caption         =   "Home"
            Height          =   625
            Left            =   4200
            TabIndex        =   11
            Top             =   0
            Width           =   750
         End
      End
   End
   Begin MdeX.xpcmdbutton cmdClose 
      Height          =   280
      Left            =   5985
      TabIndex        =   2
      Top             =   1060
      Width           =   290
      _ExtentX        =   503
      _ExtentY        =   503
      Caption         =   "X"
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
      Left            =   2760
      Top             =   1140
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   3495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   6975
      ExtentX         =   12303
      ExtentY         =   6165
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Ïðèâÿçàòü âíèç
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   6465
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Status"
            TextSave        =   "Status"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSForms.TabStrip TabStrip1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7200
      ListIndex       =   0
      Size            =   "12700;450"
      Items           =   "Tab1;Tab2;"
      TipStrings      =   ";;"
      Names           =   "Tab1;Tab2;"
      NewVersion      =   -1  'True
      TabsAllocated   =   2
      Tags            =   ";;"
      TabData         =   2
      Accelerator     =   ";;"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
      TabState        =   "3;3"
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu new_b 
         Caption         =   "New Browser               "
         Shortcut        =   ^T
      End
      Begin VB.Menu save_as 
         Caption         =   "Save as..."
         Shortcut        =   ^S
      End
      Begin VB.Menu PrP 
         Caption         =   "Properties"
      End
      Begin VB.Menu s 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu asss 
         Caption         =   "Print Options"
         Begin VB.Menu print 
            Caption         =   "Print..."
         End
         Begin VB.Menu Page_Setup 
            Caption         =   "Page Setup..."
         End
         Begin VB.Menu PrPr 
            Caption         =   "Print Preview..."
         End
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "Paste                        "
      End
      Begin VB.Menu se 
         Caption         =   "-"
      End
      Begin VB.Menu selall 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu Favorites 
      Caption         =   "Favorites"
      Begin VB.Menu AddToFav 
         Caption         =   "Add to Favorites..."
      End
      Begin VB.Menu DelFav 
         Caption         =   "Organize Favorite..."
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemSite 
         Caption         =   "Google"
         Index           =   0
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Begin VB.Menu mnuTop 
         Caption         =   "Always on Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu prop 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuTab 
      Caption         =   "Tab Menu"
      Visible         =   0   'False
      Begin VB.Menu cmdClosee 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuHelpGen 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu spsp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About            "
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


Dim Zt As String
Dim t As Boolean
Dim t_ind As Integer


Private Sub AddToFav_Click()
    Me.Enabled = False
    frmFav.Visible = True
End Sub

Private Sub Close_Click()
    End
End Sub
Private Sub cmdBack_Click()
On Error Resume Next

    Web(TabStrip1.Value).Stop
    Web(TabStrip1.Value).GoBack

End Sub

Private Sub cmdClose_Click()
If Tab_Count = 1 And Web(TabStrip1.Value).LocationURL = GetHomePage Then
Else

    If Tab_Count > 1 Then
        Web(TabStrip1.Value).Visible = False
        TabStrip1.Tabs(TabStrip1.Value).Visible = False
        'Web(TabStrip1.Value).ZOrder 0

Else
    If Is_First(TabStrip1.Value) And Tab_Count > 1 Then
        x = TabStrip1.Value
        'Unload Web(x)
        Web(x).Visible = False
        TabStrip1.Value = Next_Tab
        TabStrip1.Tabs(x).Visible = False
    
    Else
        Web(TabStrip1.Value).Navigate2 GetHomePage
    End If


End If
End If
    Combo1.Text = Web(index).LocationURL
    frmWeb.Caption = Web(index).LocationName & " - MustDie-ExploreR"
End Sub

Private Sub cmdClosee_Click()
    cmdClose_Click
End Sub

Private Sub cmdForward_Click()
On Error Resume Next

    Web(TabStrip1.Value).Stop
    Web(TabStrip1.Value).GoForward
    
End Sub
Private Sub cmdGo_Click()

    
    Url = Combo1.Text
    chd = Url
    addr = Url
    
    If InStr(1, Url, "http://www.") = 1 Then
        addr = Url
    Else
        
        If InStr(1, Url, "http://") = 1 Then addr = Url Else If InStr(1, Url, "www.") = 1 Then addr = "http://" & Url Else addr = "http://www." & Url
    End If
    addr2 = addr
    If InStr(1, Url, "http://") = 1 And InStr(1, Url, "http://www") <> 1 Then chd = Mid(Url, 7 + 1, Len(Url) - 7)
    Combo1.Text = chd
    If Check(Combo1.Text) = True Then
    

        Combo1.AddItem addr2
        AddHistory Combo1.Text
    End If
    
    Web(TabStrip1.Value).Stop
    Web(TabStrip1.Value).Navigate2 Combo1.Text
    
End Sub

Private Sub cmdHome_Click()
    Web(TabStrip1.Value).Navigate2 GetHomePage
End Sub

Private Sub cmdRefresh_Click()
    
    Bar1.Visible = True
    Web(TabStrip1.Value).Stop
    Web(TabStrip1.Value).Refresh2
    
End Sub

Private Sub cmdstop_Click()

    Web(TabStrip1.Value).Stop
    Bar1.Visible = False

End Sub

Private Sub Combo1_Click()

    Web(TabStrip1.Value).Navigate2 Combo1.Text
    
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then cmdGo_Click

End Sub
Private Sub Copy_Click()

    Clipboard.SetText Combo1.SelText
    
End Sub

Private Sub Cut_Click()

    x = Combo1.SelStart
    a1 = Left(Combo1.Text, x)
    a2 = Mid(Combo1.Text, x - Combo1.SelLength + 1, Len(Combo1.Text) - x - Combo1.SelLength)
    Clipboard.SetText Combo1.SelText
    Combo1.Text = a1 & a2
    
End Sub


Private Sub DelFav_Click()
    Me.Enabled = False
    frmFavDel.Visible = True
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    
    t = True
    mnuTop.Checked = False
    SetProc hWnd

End Sub



Private Sub mnuAbout_Click()
    Me.Enabled = False
    frmAbout.Visible = True
End Sub

Private Sub mnuHelp_Click()
    MsgBox "What for?!" & Chr(13) & "What is'nt clear?!" & Chr(13) & "It's a simple Web-Browser!!!", vbCritical, "Lamer's Window"
End Sub

Private Sub mnuItemSite_Click(index As Integer)
    Dim li As Integer
    li = index
    Web(TabStrip1.Value).Navigate2 Get_Url(li)
End Sub

Private Sub mnuTop_Click()
    mnuTop.Checked = Not (mnuTop.Checked)
    
    If mnuTop.Checked Then
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub new_b_Click()
    newww_Click
End Sub

Private Sub newww_Click()

    TabStrip1.Tabs.Add , GetHomePage
    Load Web(Web.Count)
    Web(Web.Count - 1).Visible = True
    Web(Web.Count - 1).Navigate2 "about:blank"
    TabStrip1.Value = TabStrip1.Tabs.Count - 1
    Web(TabStrip1.Tabs.Count - 1).ZOrder 0
    
End Sub

Private Sub Page_Setup_Click()
    Web(TabStrip1.Value).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub paste_Click()
    x = Combo1.SelStart
    a1 = Left(Combo1.Text, x)
    a2 = Mid(Combo1.Text, x - Combo1.SelLength + 1, Len(Combo1.Text) - x - Combo1.SelLength)
    Combo1.Text = a1 & Clipboard.GetText & a2
    With Combo1
        .SelStart = x
        .SelLength = Len(Clipboard.GetText)
    End With
End Sub


Private Sub print_Click()
    Web(TabStrip1.Value).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub prop_Click()
    frmProperties.Visible = True
    Me.Enabled = False
End Sub

Private Sub Prp_Click()
    Web(TabStrip1.Value).ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub PrPr_Click()
    Web(TabStrip1.Value).ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub save_as_Click()
    Web(TabStrip1.Value).ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub selall_Click()

    With Combo1
        .SelStart = 0
        .SelLength = Len(Combo1.Text)
    End With
    
End Sub

Private Sub TabStrip1_Click(index As Long)

    Web(index).ZOrder 0
    Combo1.Text = Web(index).LocationURL
    frmWeb.Caption = Web(index).LocationName & " - MustDie-ExploreR"
    
End Sub

Private Sub TabStrip1_MouseDown(index As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuTab
    t_ind = index
End If
TabStrip1_Click TabStrip1.Value
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Frame1.Left = 0
    Frame1.Width = Me.Width
    Bar1.Left = StatusBar1.Width - Bar1.Width - 850
    Bar1.Top = StatusBar1.Top + 30
    Web(TabStrip1.Value).Height = StatusBar1.Top - Web(TabStrip1.Value).Top
    Web(TabStrip1.Value).Width = StatusBar1.Width
    Combo1.Width = Picture2.Width - cmdGo.Width - Combo1.Left
    cmdGo.Left = Combo1.Left + Combo1.Width
    cmdClose.Left = Me.Width - cmdClose.Width - 120
End Sub
Private Sub Web_BeforeNavigate2(index As Integer, ByVal pDisp As Object, Url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    'If index = TabStrip1.Value Then Combo1.Text = Url
    Bar1.Visible = True

End Sub
Private Sub Web_DocumentComplete(index As Integer, ByVal pDisp As Object, Url As Variant)

    Bar1.Visible = False

        If index = TabStrip1.Value Then Combo1.Text = Web(index).LocationURL

End Sub

Private Sub Web_NavigateComplete2(index As Integer, ByVal pDisp As Object, Url As Variant)
    'If index = TabStrip1.Value Then Combo1.Text = Url
End Sub

Private Sub Web_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    t = False
    If Mid(Zt, 1, 7) = "http://" Then
        Web(TabStrip1.Value).Stop
        Cancel = True
        TabStrip1.Tabs.Add , Zt
        Load Web(TabStrip1.Tabs.Count - 1)
        Web(TabStrip1.Tabs.Count - 1).Visible = True
        TabStrip1.Value = (TabStrip1.Tabs.Count - 1)
        'Web(TabStrip1.Value).Navigate2 "about:blank"
        Web(TabStrip1.Value).Navigate2 Zt
        Web(TabStrip1.Value).ZOrder 0
        Combo1.Text = Zt
        frmWeb.Caption = Zt
    End If
    t = True
End Sub

Private Sub Web_ProgressChange(index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next

    If index = TabStrip1.Value Then
        Bar1.Max = ProgressMax
        Bar1.Value = Progress
    End If

End Sub


Private Sub Web_StatusTextChange(index As Integer, ByVal Text As String)

        If index = TabStrip1.Value Then StatusBar1.Panels(1).Text = Text
        If t Then Zt = Text
        

End Sub
Private Sub Web_TitleChange(index As Integer, ByVal Text As String)

    Me.Caption = Text & " - " & " - MustDie-ExploreR"
    If Len(Text) > 20 Then TabStrip1.Tabs(index).Caption = Left(Text, 20) & "..." Else TabStrip1.Tabs(index).Caption = Text
    
End Sub

