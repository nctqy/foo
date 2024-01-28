'vb
'''

Private Sub DownNetFile(ByVal nUrl As String, ByVal nFile As String)

    Dim iOpenFileFlg As Integer
    Dim XmlHttp, b() As Byte
    
    iOpenFileFlg = FreeFile
    Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    
    XmlHttp.open "GET", nUrl, False
    
    XmlHttp.Send
    
    If XmlHttp.ReadyState = 4 Then
    
    b() = XmlHttp.ResponseBody
    
    'Open nFile For Binary As #1
    Open nFile For Binary As #iOpenFileFlg
    
    'Put #1, , B()
    Put #iOpenFileFlg, , b()
    
    Close #iOpenFileFlg
    
    End If
    
    Set XmlHttp = Nothing

End Sub


Public Function LoadPicture(ByVal strFileName As String) As Picture
Dim IID As TGUID
With IID
.Data1 = &H7BF80980
.Data2 = &HBF32
.Data3 = &H101A
.Data4(0) = &H8B
.Data4(1) = &HBB
.Data4(2) = &H0
.Data4(3) = &HAA
.Data4(4) = &H0
.Data4(5) = &H30
.Data4(6) = &HC
.Data4(7) = &HAB
End With

On Error GoTo LocalErr

OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
Exit Function
LocalErr:
Set LoadPicture = VB.LoadPicture(strFileName)
Err.Clear
End Function

Sub CheckAndDowndata()
Dim sTmpSql As String
Dim sDownloadSql As String

Dim sDownloadSqlOld As String
Dim sMyCurrentDateStart As String
Dim sMyCurrentDateEnd As String

Dim sDay As String
Dim sTableTmp As String
sMyCurrentDateEnd = Format(Date, "yyyy-mm-dd")

sMyCurrentDateStart = DateAdd("d", -6, sMyCurrentDateEnd)

sMyCurrentDateStart = Mid(sMyCurrentDateStart, 1, 4) & "-" & Mid(sMyCurrentDateStart, 5, 2) & "-" & Mid(sMyCurrentDateStart, 7, 2)
 sTmpSql = "select *  from stockday  where xqday>='" & sMyCurrentDateStart & "' and xqday<='" & sMyCurrentDateEnd & "'"

sDownloadSqlOld = "DownStckTradeData,select * from "
Dim cnStr As String
Dim czCnn0321 As New adodb.Connection
Dim czRss0321 As New adodb.Recordset
  cnStr = getConstr
    Set czCnn0321 = New adodb.Connection
    czCnn0321.open cnStr
    Set czRss0321 = New adodb.Recordset
    czRss0321.open sTmpSql, czCnn0321, 1, 3
    Do While Not czRss0321.EOF
    
    sDay = czRss0321.fields("xqday").Value
    
    sTableTmp = "ALLSTCK_" & Replace(sDay, "-", "")
    sDownloadTable = sTableTmp
    sDownloadSql = sDownloadSqlOld & sTableTmp
    
'    Debug.Print sDownloadSql
    If Weekday(sDay) <> vbSaturday And Weekday(sDay) <> vbSunday Then


        If getIsNotTrade(sDay) = "1" Then

          If isTableExist(cnStr, sTableTmp) = False Then
          ''down data from server
          myFlg = 3
          sendCommand sDownloadSql
Debug.Print sDownloadSql
            Debug.Print sTableTmp

          End If
        End If

        End If
    czRss0321.MoveNext
    Loop
czRss0321.close
czCnn0321.close
End Sub




'''
