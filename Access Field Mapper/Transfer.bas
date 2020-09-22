Attribute VB_Name = "Transfer"
Public Sub TransferData()

    On Error GoTo ErrHandler
    
    If Len(FileTo) = 0 Or _
        Len(FileFrom) = 0 Then Exit Sub
    
    If frmMain.cmbTo & "" = "" Or _
        frmMain.cmbFrom & "" = "" Then Exit Sub
    
    Dim Con1 As New ADODB.Connection
    Dim Con2 As New ADODB.Connection
    Dim Rst1 As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass

    With Con1
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileFrom & ";"
        .CursorLocation = adUseServer
        .Open
    End With
    
    With Con2
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileTo & ";"
        .CursorLocation = adUseServer
        .Open
    End With
    
    Rst1.Open "SELECT * FROM [" & frmMain.cmbFrom & "]", Con1, adOpenDynamic, adLockPessimistic
    Rst2.Open "SELECT * FROM [" & frmMain.cmbTo & "]", Con2, adOpenDynamic, adLockPessimistic
    
    If frmMain.Opt1 Then 'Overwrite existing data
        Con2.Execute "Delete * From [" & frmMain.cmbTo & "]"
    End If
    
    Dim i As Integer
    Dim str1 As String, str2 As String, strSQL As String
    
    Rst1.MoveFirst
    Do Until Rst1.EOF
        Rst2.AddNew
        For i = 1 To Lines.Count
            str1 = Trim(Left(frmMain.lstFrom.List(Lines(i).Index1), intFld))
            str2 = Trim(Left(frmMain.lstTo.List(Lines(i).Index2), intFld))
            Rst2(str2) = Coerce(GetType(Rst1(str1).Type), GetType(Rst2(str2).Type), Rst1(str1).Value)
        Next i
        Rst2.Update
        Rst1.MoveNext
    Loop
        
    Screen.MousePointer = vbDefault
    
NormalExit:
    Rst1.Close
    Rst2.Close
    Con1.Close
    Con2.Close
    Set Rst1 = Nothing
    Set Rst2 = Nothing
    Set Con1 = Nothing
    Set Con2 = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly + vbCritical, "Error"
    GoTo NormalExit
    
End Sub

'Used to coerce different data types on transfer
Public Function Coerce(strType1 As String, strType2 As String, varData As Variant) As Variant
    Select Case strType1
        Case "Text"
            Select Case strType2
                Case "Text"
                    Coerce = varData
                Case "Memo"
                    Coerce = varData
                Case Else
                    Coerce = Null
            End Select
        Case "Memo"
            Select Case strType2
                Case "Text"
                    Coerce = varData
                Case "Memo"
                    Coerce = varData
                Case Else
                    Coerce = Null
            End Select
        Case "OLE Object"
            Select Case strType2
                Case "OLE Object"
                    Coerce = varData
                Case Else
                    Coerce = Null
            End Select
        Case "Byte"
            Select Case strType2
                Case "Byte"
                    Coerce = varData
                Case Else
                    Coerce = Null
            End Select
        Case "Long Integer"
            Select Case strType2
                Case "Long Integer"
                    Coerce = varData
                Case "Integer"
                    Coerce = Int(varData)
                Case "Single"
                    Coerce = CSng(varData)
                Case "Double"
                    Coerce = CDbl(varData)
                Case "Yes/No"
                    Coerce = CBool(varData)
                Case "Text"
                    Coerce = Str(varData)
                Case "Memo"
                    Coerce = Str(varData)
                Case Else
                    Coerce = Null
            End Select
        Case "Integer"
            Select Case strType2
                Case "Long Integer"
                    Coerce = varData
                Case "Integer"
                    Coerce = varData
                Case "Single"
                    Coerce = CSng(varData)
                Case "Double"
                    Coerce = CDbl(varData)
                Case "Yes/No"
                    Coerce = CBool(varData)
                Case "Text"
                    Coerce = Str(varData)
                Case "Memo"
                    Coerce = Str(varData)
                Case Else
                    Coerce = Null
            End Select
        Case "Single"
            Select Case strType2
                Case "Long Integer"
                    Coerce = CLng(varData)
                Case "Integer"
                    Coerce = CInt(varData)
                Case "Single"
                    Coerce = varData
                Case "Double"
                    Coerce = varData
                Case "Text"
                    Coerce = Str(varData)
                Case "Memo"
                    Coerce = Str(varData)
                Case Else
                    Coerce = Null
            End Select
        Case "Double"
            Select Case strType2
                Case "Long Integer"
                    Coerce = CLng(varData)
                Case "Integer"
                    Coerce = CInt(varData)
                Case "Single"
                    Coerce = CSng(varData)
                Case "Double"
                    Coerce = varData
                Case "Text"
                    Coerce = Str(varData)
                Case "Memo"
                    Coerce = Str(varData)
                Case Else
                    Coerce = Null
            End Select
        Case "Yes/No"
            Select Case strType2
                Case "Yes/No"
                    Coerce = varData
                Case "Text", "Memo"
                    Coerce = CStr(varData)
                Case "Integer", "Long Integer"
                    Coerce = CInt(varData)
                Case Else
                    Coerce = Null
            End Select
        Case "Replication ID"
            Select Case strType2
                Case "Replication ID"
                    Coerce = varData
                Case Else
                    Coerce = Null
            End Select
        Case "Date/Time"
            Select Case strType2
                Case "Date/Time"
                    Coerce = varData
                Case "Text"
                    Coerce = CDate(varData)
                Case "Memo"
                    Coerce = CDate(varData)
                Case Else
                    Coerce = Null
            End Select
        Case Else
            Coerce = Null
    End Select
End Function

