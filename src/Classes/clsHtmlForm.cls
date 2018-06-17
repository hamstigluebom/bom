VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHtmlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msFormSource As String
Private mcolFields As Collection
Private mcolTypes As Collection
Private mcolValues As Collection

Private Sub Class_Initialize()
    
    Set mcolFields = New Collection
    Set mcolTypes = New Collection
    Set mcolValues = New Collection
    
End Sub

Private Sub AddField(sFieldName As String, sFieldType As String, sFieldValue As String)

  On Error Resume Next
  If sFieldType = "" Then sFieldType = "text"
  mcolFields.Add sFieldName, sFieldName
  mcolTypes.Add sFieldType, sFieldName
  mcolValues.Add HtmlZeichenConvert(sFieldValue), sFieldName

End Sub

Public Sub PutField(sName As String, sValue As String)
    
    Dim i As Long
    
    For i = 1 To mcolFields.Count
        If LCase(mcolFields(i)) = LCase(sName) Then
            mcolValues.Remove mcolFields(i)
            mcolValues.Add sValue, mcolFields(i)
            Exit For
        End If
    Next i
    
End Sub

Public Function GetField(sName As String) As String
    
    Dim i As Long
    
    For i = 1 To mcolFields.Count
        If LCase(mcolFields(i)) = LCase(sName) Then
            GetField = mcolValues(mcolFields(i))
            Exit For
        End If
    Next i
    
End Function

Public Function GetFieldType(sName As String) As String
    
    Dim i As Long
    
    For i = 1 To mcolFields.Count
        If LCase(mcolFields(i)) = LCase(sName) Then
            GetFieldType = LCase(mcolTypes(mcolFields(i)))
            Exit For
        End If
    Next i
    
End Function

Public Sub ReadForm(sHtmlCode As String, sFormName As String)
    
    msFormSource = GetHtmlForm(sHtmlCode, sFormName)
    Call ReadFields
    
End Sub

Public Function GetFields(sEncoding As String) As String
    
    Dim i As Long
    Dim sTmp As String
    Dim sTmpVal As String
    
    For i = 1 To mcolFields.Count
        If LCase(mcolTypes(mcolFields(i))) <> "submit" And LCase(mcolTypes(mcolFields(i))) <> "image" Then
            sTmpVal = mcolValues(mcolFields(i))
            If sEncoding = "utf-8" Then sTmpVal = Encode_UTF8(sTmpVal)
            sTmp = sTmp & mcolFields(i) & "=" & URLEncode(sTmpVal) & "&"
        End If
    Next i
    
    If Len(sTmp) > 0 Then sTmp = Left(sTmp, Len(sTmp) - 1)
    GetFields = sTmp
    
End Function

Public Function GetAction() As String

  GetAction = GetHtmlValue(msFormSource, "action=""", "", """")
  If GetAction Like "*[?]*" Then GetAction = Left(GetAction, InStr(1, GetAction, "?") - 1)

End Function

Private Function GetHtmlForm(sTxt As String, sFormName As String) As String

  Dim lPosStart As Long
  Dim lPosEnde As Long
  Dim lPos As Long
  
  lPosStart = InStr(lPos + 1, sTxt, "<form ", vbTextCompare)
  Do While lPosStart > 0
    
    lPosEnde = InStr(lPosStart, sTxt, "</form>", vbTextCompare)
    If lPosEnde = 0 Then lPosEnde = 2000000
    
'    lPos = InStr(lPosStart, sTxt, "method=""post""", vbTextCompare)
'    If lPos > lPosStart And lPos < lPosEnde Then 'method=post gefunden
      
      lPos = InStr(lPosStart, sTxt, "name=""" & sFormName & """", vbTextCompare)
      If lPos = 0 Then lPos = InStr(lPosStart, sTxt, "name=" & sFormName & " ", vbTextCompare)
      ' Wenn die gesuchte Form ohne Namen ist, schauen wir mal ob die jetzige Form überhaupt ein Namens-Attribut hat
      If lPos = 0 And sFormName = "" Then
        lPos = InStr(lPosStart, sTxt, ">")
        If lPos > 0 Then
          If InStr(lPosStart, sTxt, "name=", vbTextCompare) < lPos Then lPos = 0
        End If
      End If
      If lPos > lPosStart And lPos < lPosEnde Then 'richtige form gefunden
        GetHtmlForm = Mid(sTxt, lPosStart, lPosEnde - lPosStart + 7)
        Exit Function
      Else 'nächste form suchen
        lPos = lPosEnde
      End If
      
'    Else 'nächste form suchen
'      lPos = lPosEnde
'    End If
    
    lPosStart = InStr(lPos + 1, sTxt, "<form ", vbTextCompare)
    
  Loop

End Function

Private Function GetHtmlValue(sTxt As String, sSuch1 As String, sSuch2 As String, sSuch3 As String)

  Dim lPos As Long
  Dim lPosStart As Long
  
  lPos = InStr(1, sTxt, sSuch1, vbTextCompare) 'such1
  If lPos > 0 Then
    lPos = lPos + Len(sSuch1)
    If sSuch2 <> "" Then
      lPos = InStr(lPos, sTxt, sSuch2, vbTextCompare) 'such2
      If lPos > 0 Then lPos = lPos + Len(sSuch2)
    End If
    If lPos > 0 Then
      lPosStart = lPos
      lPos = InStr(lPosStart, sTxt, sSuch3) 'such3
      If lPos > lPosStart Then
        GetHtmlValue = Mid(sTxt, lPosStart, lPos - lPosStart)
      End If
    End If
  End If

End Function

Private Sub ReadFields()
    
    Dim lPos As Long
    Dim lPosStart As Long
    Dim lPosEnde As Long
    Dim lPosEnde2 As Long
    Dim sBuffer As String
    
    lPos = InStr(lPos + 1, msFormSource, "<input", vbTextCompare)
    Do While lPos > 0
        sBuffer = ""
        lPosStart = lPos
        lPosEnde = InStr(lPosStart, msFormSource, """>")
        lPosEnde2 = InStr(lPosStart, msFormSource, "/>")
        If lPosEnde2 > 0 And (lPosEnde2 < lPosEnde Or lPosEnde = 0) Then lPosEnde = lPosEnde2
        
        If lPosEnde > lPosStart Then sBuffer = Mid(msFormSource, lPosStart, lPosEnde + 2 - lPosStart)
        
        AddField GetHtmlValue(sBuffer, "name=""", "", """"), _
            GetHtmlValue(sBuffer, "type=""", "", """"), _
            GetHtmlValue(sBuffer, "value=""", "", """")
        
        lPos = InStr(lPos + 1, msFormSource, "<input", vbTextCompare)
    Loop
    
    lPos = InStr(lPos + 1, msFormSource, "<select", vbTextCompare)
    Do While lPos > 0
        sBuffer = ""
        lPosStart = lPos
        lPosEnde = InStr(lPosStart, msFormSource, "</select>")
        If lPosEnde > lPosStart Then sBuffer = Mid(msFormSource, lPosStart, lPosEnde + 9 - lPosStart)
        
        AddField GetHtmlValue(sBuffer, "name=""", "", """"), _
            "select", _
            GetSelectedOrFirstOptionValue(sBuffer)
        
        lPos = InStr(lPos + 1, msFormSource, "<select", vbTextCompare)
    Loop
    
End Sub

Private Function GetSelectedOrFirstOptionValue(sSelectSource As String) As String

    Dim lPos As Long
    Dim lPosStart As Long
    Dim lPosEnde As Long
    Dim sBuffer As String
    
    lPos = InStr(lPos + 1, sSelectSource, "<option", vbTextCompare)
    Do While lPos > 0
        sBuffer = ""
        lPosStart = lPos
        lPosEnde = InStr(lPosStart, sSelectSource, ">")
        If lPosEnde > lPosStart Then sBuffer = Mid(sSelectSource, lPosStart, lPosEnde + 1 - lPosStart)
        
        If GetSelectedOrFirstOptionValue = "" Or InStr(1, sBuffer, "selected", vbTextCompare) > 0 Then
            GetSelectedOrFirstOptionValue = GetHtmlValue(sBuffer, "value=""", "", """")
        End If
        
        lPos = InStr(lPos + 1, sSelectSource, "<option", vbTextCompare)
    Loop

End Function

Private Sub Class_Terminate()
    
    Set mcolFields = Nothing
    Set mcolTypes = Nothing
    Set mcolValues = Nothing
    
End Sub

Public Sub Clear()

  Dim i As Long
  
  For i = 1 To mcolFields.Count
    mcolFields.Remove 1
    mcolValues.Remove 1
    mcolTypes.Remove 1
    msFormSource = ""
  Next i

End Sub

Public Function FormFound() As Boolean

  FormFound = CBool(msFormSource <> "")

End Function

Public Function ClickImage(sImageName As String) As String

  If GetFieldType(sImageName) = "image" Then ClickImage = "&" & sImageName & ".x=" & CStr(Int(Rnd() * 9) + 1) & "&" & sImageName & ".y=" & CStr(Int(Rnd() * 9) + 1)
  
End Function
