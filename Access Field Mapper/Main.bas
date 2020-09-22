Attribute VB_Name = "modMain"
Public FileTo As String 'Name of Database to transfer data to
Public FileFrom As String 'Name of Database to transfer data from
Public cn1 As ADODB.Connection 'Connection object
Public rs1 As ADODB.Recordset 'Recordset object
Public DB1Tables() As String 'Array to hold table names
Public DB2Tables() As String 'Array to hold table names
Public LineFrom As Long 'X value of right side of lstFrom
Public LineTo As Long 'X value of left side of lstTo
Public Cnr As Connector 'Current Connector to draw to screen
Public Lines As New Collection 'our collection of connectors
Public lstEnabled As Boolean 'Keeps track of whether or not both lists have data
Public LineNo As Long 'Used for creating unique names in the initialization of a connector

'Variable to keep track of which type of connection is currently in process
'1 = From-To connection
'2 = To-From connection
Public DrawConnector As Integer

Public Const LineY = 211 'Height of list item
Public Const yOffset = 135 'so that line sits in middle of list item
Public Const intDisplayed = 16 'Number of fields displayed in a list box
Public Const intFld = 21 'width of field name text in the list box
Public Const intType = 14 'width of the type text in the list box
Public Const intLen = 3 'width of the field size in the list box

Public Sub Main()
    LineFrom = frmMain.lstFrom.Left + frmMain.lstFrom.Width
    LineTo = frmMain.lstTo.Left
    frmMain.Show
    
    'For Testing purposes
    '+++++++++++++++++++++++++++++++++++++++++++++++++++
    'FileTo = "C:\Program Files\MSU Extension Services\Mailroom\Mailroom.mdb"
    'FileFrom = "C:\Program Files\MSU Extension Services\Mailroom\Mailroom.mdb"
    'frmMain.cmbFrom = "aaTest"
    'frmMain.cmbTo = "aaTest2"
    'PopulateList 2
    'PopulateList 1
    'frmMain.AutoMap
    '+++++++++++++++++++++++++++++++++++++++++++++++++++
    
End Sub

'Populate the list with the field names
Public Sub PopulateList(iWhichDB As Integer)

    If iWhichDB = 2 Then
        If frmMain.cmbTo & "" = "" Then Exit Sub
    Else
        If frmMain.cmbFrom & "" = "" Then Exit Sub
    End If
    
    Dim rsField As ADODB.Field
    Screen.MousePointer = vbHourglass
    
    Set cn1 = New ADODB.Connection
    With cn1
        'setup the connection
        If iWhichDB = 2 Then
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileTo & ";"
        Else
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileFrom & ";"
        End If
        .CursorLocation = adUseClient
        .Open
    End With
    
    'Clear the list
    Dim i As Integer
    If iWhichDB = 1 Then
        For i = frmMain.lstFrom.ListCount - 1 To 0 Step -1
            frmMain.lstFrom.RemoveItem i
        Next i
    Else
        For i = frmMain.lstTo.ListCount - 1 To 0 Step -1
            frmMain.lstTo.RemoveItem i
        Next i
    End If
    
    'Open the recordset
    Set rs1 = New ADODB.Recordset
    rs1.ActiveConnection = cn1
    If iWhichDB = 2 Then
        rs1.Open "SELECT * FROM [" & frmMain.cmbTo & "]"
    Else
        rs1.Open "SELECT * FROM [" & frmMain.cmbFrom & "]"
    End If
        
    'Add the field names to the list
    With rs1
        For Each rsField In .Fields
            If iWhichDB = 2 Then
                frmMain.lstTo.AddItem FixedStr(rsField.Name, intFld) _
                & " " & FixedStr(GetType(rsField.Type), intType) _
                & " " & FixedStr(FixedSize(rsField.DefinedSize), intLen)
            Else
                frmMain.lstFrom.AddItem FixedStr(rsField.Name, intFld) _
                & " " & FixedStr(GetType(rsField.Type), intType) _
                & " " & FixedStr(FixedSize(rsField.DefinedSize), intLen)
            End If
        Next
    End With
    
    Screen.MousePointer = vbDefault
    rs1.Close
    Set rs1 = Nothing
    
    'if both lists have fields then enable them for connectors
    If frmMain.lstFrom.ListCount > 0 And frmMain.lstTo.ListCount > 0 Then
        frmMain.lstFrom.Enabled = True
        frmMain.lstTo.Enabled = True
        lstEnabled = True
    Else
        frmMain.lstFrom.Enabled = False
        frmMain.lstTo.Enabled = False
        lstEnabled = False
    End If
End Sub

'Get the table names from the database
Public Function GetDBTables(iWhichDB As Integer) As Boolean
    
    Dim sTablename As String
    
    If iWhichDB = 1 Then
        If Len(FileTo) = 0 Then Exit Function
    Else
        If Len(FileFrom) = 0 Then Exit Function
    End If
    
    On Error GoTo ErrGetDBTables
    
    Screen.MousePointer = vbHourglass

    Set cn1 = New ADODB.Connection
    With cn1
        'setup the connection
        If iWhichDB = 1 Then
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileTo & ";"
        Else
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileFrom & ";"
        End If
        .CursorLocation = adUseClient
        .Open
        DoEvents
        
        'get the list of tables using the  Openschema method
        'The OpenSchema method has three arguments (third is optional)
        'The first argument identifies the type of schema information to return
        '   Identified as the Query Type
        'The second is an array that sets the Constraints and the Column names
        'This array is diffent for each of the Query Types
        'When working with MS Access you need to replace all constraints that
        '   are invalid with empty.  As in this case only the Table Name is needed
        'Once the Recordset is built then the Remaining of the processing
        'remains the same as when processing actual data.
        Set rs1 = .OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
        
        'Setup the Arrays
        If iWhichDB = 1 Then
            ReDim Preserve DB1Tables(0)
        Else
            ReDim Preserve DB2Tables(0)
        End If
        
        'Loop the Recordset and retrieve the table names
        Do While Not rs1.EOF
            sTablename = rs1!TABLE_NAME
            If iWhichDB = 1 Then
                ReDim Preserve DB1Tables(UBound(DB1Tables) + 1)
                DB1Tables(UBound(DB1Tables)) = sTablename
            Else
                ReDim Preserve DB2Tables(UBound(DB2Tables) + 1)
                DB2Tables(UBound(DB2Tables)) = sTablename
            End If
            rs1.MoveNext
        Loop
        
        'Clean Up
        rs1.Close
        DoEvents
        .Close
        DoEvents
    End With
    GetDBTables = True
    
exitGetDBTables:
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrGetDBTables:
    GetDBTables = False
    Resume exitGetDBTables
    
End Function

Public Sub DragDefault()
    Screen.MousePointer = vbDefault
    If lstEnabled Then 'if both lists have data then enable them
        frmMain.lstTo.Enabled = True
        frmMain.lstFrom.Enabled = True
    End If
End Sub

Public Sub DragTo()
    Screen.MousePointer = vbNoDrop
    frmMain.lstTo.Enabled = False
End Sub

Public Sub DragFrom()
    Screen.MousePointer = vbNoDrop
    frmMain.lstFrom.Enabled = False
End Sub

Public Sub DropHere()
    Screen.MousePointer = vbCustom
End Sub

Public Function FixedStr(strString As String, Characters As Integer) As String
    If Len(strString) >= Characters Then
        FixedStr = Left(strString, Characters)
    Else
        FixedStr = strString & Space(Characters - Len(strString))
    End If
End Function

Public Function FixedSize(lngSize As Long) As String
    If lngSize > 255 Then
        FixedSize = "N/A"
    Else
        FixedSize = Trim(Str(lngSize))
    End If
End Function

Public Function GetType(TypeVal As Integer) As String
    Select Case TypeVal
        Case 202
            GetType = "Text"
        Case 203
            GetType = "Memo"
        Case 205
            GetType = "OLE Object"
        Case 17
            GetType = "Byte"
        Case 3
            GetType = "Long Integer"
        Case 2
            GetType = "Integer"
        Case 4
            GetType = "Single"
        Case 5
            GetType = "Double"
        Case 11
            GetType = "Yes/No"
        Case 72
            GetType = "Replication ID"
        Case 7
            GetType = "Date/Time"
        Case Else
            GetType = "Not Recognized"
    End Select
    
End Function
'Used to see if two fields are of compatible types to transfer the data
'Returns    0 for false
'           1 for true
'           2 for true with possible loss of data
Public Function Compatible(strType1 As String, strType2 As String) As Integer
    Select Case strType1
        Case "Text"
            Select Case strType2
                Case "Text"
                    Compatible = 1
                Case "Memo"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Memo"
            Select Case strType2
                Case "Text"
                    Compatible = 2
                Case "Memo"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "OLE Object"
            Select Case strType2
                Case "OLE Object"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Byte"
            Select Case strType2
                Case "Byte"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Long Integer"
            Select Case strType2
                Case "Long Integer"
                    Compatible = 1
                Case "Integer"
                    Compatible = 2
                Case "Single"
                    Compatible = 1
                Case "Double"
                    Compatible = 1
                Case "Yes/No"
                    Compatible = 2
                Case "Text"
                    Compatible = 1
                Case "Memo"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Integer"
            Select Case strType2
                Case "Long Integer"
                    Compatible = 1
                Case "Integer"
                    Compatible = 1
                Case "Single"
                    Compatible = 1
                Case "Double"
                    Compatible = 1
                Case "Yes/No"
                    Compatible = 2
                Case "Text"
                    Compatible = 1
                Case "Memo"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Single"
            Select Case strType2
                Case "Long Integer"
                    Compatible = 2
                Case "Integer"
                    Compatible = 2
                Case "Single"
                    Compatible = 1
                Case "Double"
                    Compatible = 1
                Case "Text"
                    Compatible = 1
                Case "Memo"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Double"
            Select Case strType2
                Case "Long Integer"
                    Compatible = 2
                Case "Integer"
                    Compatible = 2
                Case "Single"
                    Compatible = 2
                Case "Double"
                    Compatible = 1
                Case "Text"
                    Compatible = 1
                Case "Memo"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Yes/No"
            Select Case strType2
                Case "Yes/No"
                    Compatible = 1
                Case "Text", "Memo"
                    Compatible = 1
                Case "Integer", "Long Integer"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Replication ID"
            Select Case strType2
                Case "Replication ID"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case "Date/Time"
            Select Case strType2
                Case "Date/Time"
                    Compatible = 1
                Case "Text"
                    Compatible = 1
                Case "Memo"
                    Compatible = 1
                Case Else
                    Compatible = 0
            End Select
        Case Else
            Compatible = 0
    End Select
End Function

'This is a check for field size of text fields
Public Function LenCompatible(strType1 As String, strType2 As String, fldSize1 As String, fldSize2 As String) As Boolean
    Select Case strType1
        Case "Text"
            If strType2 = "Text" Then
                If Val(fldSize1) > Val(fldSize2) Then
                    LenCompatible = False
                Else
                    LenCompatible = True
                End If
            Else
                LenCompatible = True
            End If
        Case Else
            LenCompatible = True
    End Select
End Function
