Attribute VB_Name = "modCurl"
Option Explicit
Private mcolCookieFiles As New Collection
Private mcolParams As New Collection

Public Function Curl(sUrl As String, sPostData As String, sReferer As String, sEBayUser As String, Optional bWait As Boolean = True) As String
    
    On Error GoTo ERROR_HANDLER
    
    Dim i As Integer
    Dim F As Long
    Dim t As String
    Dim p1 As Integer
    Dim p2 As Integer
    Dim c As String
    Dim sCookieFile As String
    Dim lProcessID As Long
    Dim lTimestamp As Long
    Dim sKey As String
    Dim a As Variant
    Dim v As Variant
    Dim colCurlConfig As Collection
    Dim colCurlOptions As Collection
    
    Set colCurlConfig = New Collection
    Set colCurlOptions = New Collection

    If sEBayUser = "" Then sEBayUser = "default"
    a = Array("FILE_HEADER", "FILE_STDOUT", "FILE_STDERR", "FILE_TRACE", "FILE_CONFIG")
    
    
    If ExistCollectionKey(mcolCookieFiles, sEBayUser) Then
        sCookieFile = mcolCookieFiles(sEBayUser)
    Else
        sCookieFile = MakeTempFile()
        mcolCookieFiles.Add sCookieFile, sEBayUser
    End If
    
    With colCurlConfig
        .Add sCookieFile, "FILE_COOKIE"
        
        For Each v In a
            Call .Add(MakeTempFile(), CStr(v))
            Call SaveToFile("", colCurlConfig(CStr(v)))
        Next
        
        Call .Add("Accept-Language: " & gsBrowserLanguage, "HEADER")
        Call .Add(gsBrowserIdString, "USER_AGENT")
        Call .Add(Int(glHttpTimeOut / 1000), "CONNECT_TIMEOUT")
        Call .Add(Int(glHttpTimeOut / 1000), "TRANSFER_TIMEOUT")
        Call .Add(10, "MAX_REDIRECT")
        Call .Add(3, "MAX_RETRY")
        Call .Add(gsProxyName & IIf(giProxyPort > 0, ":" & CStr(giProxyPort), ""), "PROXY_SERVER")
        Call .Add(gsProxyUser & IIf(gsProxyPass > "", ":" & gsProxyPass, ""), "PROXY_CREDENTIALS")
        Call .Add(sUrl, "URL")
        Call .Add(sReferer, "REFERER")
        Call .Add(sPostData, "DATA")
    End With 'colCurlConfig
    
    With colCurlOptions
        Call .Add("--dump-header {FILE_HEADER}")
        Call .Add("--compressed")
'       Call .Add("--fail")
        Call .Add("--silent")
        Call .Add("--show-error")
        Call .Add("--user-agent {USER_AGENT}")
        Call .Add("--connect-timeout {CONNECT_TIMEOUT}")
        Call .Add("--location")
        Call .Add("--max-time {TRANSFER_TIMEOUT}")
        Call .Add("--max-redirs {MAX_REDIRECT}")
        Call .Add("--retry {MAX_RETRY}")
        Call .Add("--output {FILE_STDOUT}")
        Call .Add("--stderr {FILE_STDERR}")
        Call .Add("--header {HEADER}")
        Call .Add("--url {URL}")
        
        If sReferer > "" Then Call .Add("--referer {REFERER}")
        If sPostData > "" Then Call .Add("--data {DATA}")
        
        If sEBayUser <> "anonymous" Then
          Call .Add("--cookie {FILE_COOKIE}")
          Call .Add("--cookie-jar {FILE_COOKIE}")
        End If
        
        If gbUseProxy Then
            Call .Add("--proxy {PROXY_SERVER}")
            If gsProxyUser > "" Then
                Call .Add("--proxy-user {PROXY_USER}:{PROXY_PASS}")
                Call .Add("--proxy-anyauth")
            End If
        End If
        
        If Dir(App.Path & "\curl-ca-bundle.crt") = "" Then
            Call .Add("--insecure")
        End If
        
        Call .Add("#--trace-ascii {FILE_TRACE}")
        Call .Add("#--trace-time")
        Call .Add("#--limit-rate 7168")
    End With 'colCurlOptions
    
    F = FreeFile()
    Open colCurlConfig("FILE_CONFIG") For Output As F
        For i = 1 To colCurlOptions.Count
            t = colCurlOptions.Item(i)
            p1 = InStr(1, t, "{")
            p2 = InStr(p1 + 1, t, "}")
            Do While p1 > 0 And p2 > 0
                c = Mid(t, p1 + 1, p2 - p1 - 1)
                c = Replace(colCurlConfig.Item(c), IIf(Left(c, 5) = "FILE_", "\", "/"), "/")
                t = Replace(t, Mid(t, p1, p2 - p1 + 1), """" & c & """")
                p1 = InStr(1, t, "{")
                p2 = InStr(p1 + 1, t, "}")
            Loop
            Print #F, t
        Next i
    Close F
  
    lProcessID = ShellStart("""" & App.Path & "\curl.exe"" --config """ & colCurlConfig("FILE_CONFIG") & """", vbHide)
    
    If lProcessID <= 0 Then DebugPrint "no curl process id, url = " & sUrl
    
    lTimestamp = Timer * 100
    sKey = lProcessID & "/" & lTimestamp
    
    mcolParams.Add a, "a" & sKey
    mcolParams.Add colCurlConfig, "c" & sKey
    mcolParams.Add sUrl, "u" & sKey
    mcolParams.Add sKey, "k" & sKey
    
    If bWait Then
        Do
            Call Sleep(10)
            DoEvents
        Loop While ShellStillRunning(lProcessID)
        Curl = CurlGetData(sKey)
    Else
        mcolParams.Add lProcessID, sKey
        Curl = lProcessID
    End If
    
    Set colCurlConfig = Nothing
    Set colCurlOptions = Nothing
    
Exit Function
ERROR_HANDLER:
    DebugPrint "error in function curl: " & Err.Description
    Err.Clear
        
End Function

Private Function CurlGetData(sKey As String, Optional ByRef sUrlReturn As String) As String
    
    Dim a As Variant
    Dim v As Variant
    Dim vntCurlConfig As Variant
    Dim sUrl As String
    Dim bOk As Boolean
    Dim b() As Byte
    Dim sServerHeader As String
    Dim sError As String

    If ExistCollectionKey(mcolParams, "c" & sKey) Then
    
        a = mcolParams("a" & sKey)
        Set vntCurlConfig = mcolParams("c" & sKey)
        sUrl = mcolParams("u" & sKey)
        
        sUrlReturn = sUrl
        
        If FileLen(vntCurlConfig("FILE_STDOUT")) > 0 Then
            b() = ReadFromFile(vntCurlConfig("FILE_STDOUT"), True)
            If UBound(b()) >= LBound(b()) Then
                bOk = True
            End If
        End If
        
        If bOk Then
            
            sServerHeader = ReadFromFile(vntCurlConfig("FILE_HEADER"))
            'MsgBox "SiteEncoding: " & gsSiteEncoding & vbCrLf & "ServerHeader: " & vbCrLf & sServerHeader
            If InStr(1, sServerHeader, "charset=utf-8", vbTextCompare) > 0 Then
                CurlGetData = ByteArray2String(Decode_UTF8(b))
                CurlGetData = Replace(CurlGetData, "charset=utf-8", "charset=ISO-8859-1", , , vbTextCompare)
            Else
                CurlGetData = StrConv(b, vbUnicode)
            End If
        Else
            sError = ReadFromFile(vntCurlConfig("FILE_STDERR"))
            Do While Right(sError, 1) = vbCr Or Right(sError, 1) = vbLf: sError = Left(sError, Len(sError) - 1): Loop
            frmHaupt.SetStatus sError, True
            DoEvents
            Call DebugPrint(sError & " (" & sUrl & ")")
        End If
        
        On Error Resume Next
        For Each v In a
            Call Kill(vntCurlConfig.Item(v))
        Next
        On Error GoTo 0
        
        Do While vntCurlConfig.Count > 0
            vntCurlConfig.Remove 1
        Loop
        
        mcolParams.Remove "a" & sKey
        mcolParams.Remove "c" & sKey
        mcolParams.Remove "u" & sKey
        mcolParams.Remove "k" & sKey
        If ExistCollectionKey(mcolParams, sKey) Then mcolParams.Remove sKey
    End If 'ExistCollectionKey(mcolParams, "c" & sKey)
    
End Function

Public Function PollPendingCurls(sUrlReturn As String, sDataReturn As String) As Boolean
    
    Dim i As Integer
    Dim sKey As String
    
    For i = 1 To mcolParams.Count
        If TypeName(mcolParams(i)) = "String" Then
            sKey = mcolParams(i)
            If ExistCollectionKey(mcolParams, "" & sKey) And ExistCollectionKey(mcolParams, "k" & sKey) Then
                If Not ShellStillRunning(mcolParams(sKey)) Then
                    sDataReturn = CurlGetData(sKey, sUrlReturn)
                    PollPendingCurls = CBool(sDataReturn > "")
                    Exit For
                End If
            End If
        End If
    Next i
    
End Function

Public Sub RemoveCookies()
    
    On Error Resume Next
    Dim i As Integer
    
    For i = 1 To mcolCookieFiles.Count
        Call Kill(mcolCookieFiles(i))
    Next
    On Error GoTo 0
    
End Sub

Public Function TestForCurl() As Boolean
    
    If Dir(App.Path & "\curl.exe") > "" Then TestForCurl = True
    
End Function
