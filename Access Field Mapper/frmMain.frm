VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Access Field Mapper"
   ClientHeight    =   6705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":030A
   ScaleHeight     =   6705
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmOption 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8520
      TabIndex        =   15
      Top             =   5760
      Width           =   2175
      Begin VB.OptionButton Opt2 
         Caption         =   "Append"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Overwrite"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Execute Auto Map"
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer Data"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   6240
      Width           =   2775
   End
   Begin VB.CommandButton cmdRemAll 
      Caption         =   "Remove All Mappings"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "AutoMap"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   960
      Value           =   1  'Checked
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CDB1 
      Left            =   3600
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Choose To/From Database File"
      Filter          =   "*.mdb"
      InitDir         =   "C:\windows\Desktop\"
   End
   Begin VB.ComboBox cmbFrom 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   4575
   End
   Begin VB.ComboBox cmbTo 
      Height          =   315
      Left            =   7200
      TabIndex        =   8
      Top             =   960
      Width           =   4575
   End
   Begin VB.ListBox lstFrom 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      ItemData        =   "frmMain.frx":088C
      Left            =   120
      List            =   "frmMain.frx":088E
      TabIndex        =   7
      Top             =   1800
      Width           =   4575
   End
   Begin VB.ListBox lstTo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      ItemData        =   "frmMain.frx":0890
      Left            =   7200
      List            =   "frmMain.frx":0892
      TabIndex        =   6
      Top             =   1800
      Width           =   4575
   End
   Begin VB.CommandButton cmdFrom 
      Caption         =   "Open From Database"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdTo 
      Caption         =   "Open To Database"
      Height          =   255
      Left            =   10080
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   6120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label10 
      Caption         =   "Field Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   24
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   23
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Field Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "To Table:"
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "From Table:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "From Database"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "To Database"
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFrom_click()
    PopulateList 1 'Get field names
    ClearLines 'Delete all connector previously created
    AutoMap 'Run Automap
End Sub

Private Sub cmbTo_Click()
    PopulateList 2
    ClearLines
    AutoMap
End Sub

Private Sub cmdAuto_Click()
    AutoMap
End Sub

Private Sub cmdFrom_Click()
    On Error GoTo ErrHandler
    CDB1.ShowOpen 'Open Common Dialog
    FileFrom = CDB1.FileName 'Get filename
    If FileFrom = "" Then Exit Sub
    txtFrom = FileFrom 'Show value returned
    
    'Clear ComboBox of table names
    For i = cmbFrom.ListCount - 1 To 0 Step -1
        cmbFrom.RemoveItem i
    Next i
    
    'Get table names
    If GetDBTables(2) = True Then
        For i = 1 To UBound(DB2Tables)
            cmbFrom.AddItem DB2Tables(i)
        Next i
    End If
    
    'Delete all previously existing connectors
    ClearLines
    For i = lstFrom.ListCount - 1 To 0 Step -1
        lstFrom.RemoveItem i
    Next i
    Exit Sub
    
ErrHandler:
    If Err = 32755 Then 'Cancel was pressed
        Exit Sub
    End If
End Sub

Private Sub cmdRemAll_Click()
    ClearLines
End Sub

Private Sub cmdTo_Click()
    On Error GoTo ErrHandler
    CDB1.ShowOpen
    FileTo = CDB1.FileName
    If FileTo = "" Then
        Exit Sub
    End If
    txtTo = FileTo
    
    For i = cmbTo.ListCount - 1 To 0 Step -1
        cmbTo.RemoveItem i
    Next i
    
    If GetDBTables(1) = True Then
        For i = 1 To UBound(DB1Tables)
            cmbTo.AddItem DB1Tables(i)
        Next i
    End If
    
    ClearLines
    For i = lstTo.ListCount - 1 To 0 Step -1
        lstTo.RemoveItem i
    Next i
    Exit Sub
    
ErrHandler:
    If Err = 32755 Then 'Cancel was pressed
        Exit Sub
    End If
End Sub

Private Sub cmdTransfer_Click()
    TransferData
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Load the Custom Cursor
    Screen.MouseIcon = LoadPicture(App.Path & "\DragCursor.cur")
    Randomize 'Used in creating random colors for connectors
End Sub





'Used to draw the line
'Generates movement of the line by redrawing on every mouse move
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If DrawConnector = 1 Then 'From - To
            Cnr.Draw LineFrom, lstFrom.Top + yOffset + LineY * _
                    (lstFrom.ListIndex - lstFrom.TopIndex), CLng(X), CLng(Y)
        
    ElseIf DrawConnector = 2 Then 'To - From
            Cnr.Draw LineTo, lstTo.Top + yOffset + LineY * _
                    (lstTo.ListIndex - lstTo.TopIndex), CLng(X), CLng(Y)
    End If
End Sub

'Used for resetting the connection if needed
'You discontinue a connector by clicking on its origin list
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawConnector = 1 Then 'From - To
        If Within(lstFrom, X, Y) Then 'if we click inside the From list after we have started a connector
            DragDefault 'Return cursor to normal
            DrawConnector = 0 'Reset current connection state to normal
            Set Cnr = Nothing 'Delete our connector
        End If
    ElseIf DrawConnector = 2 Then 'To - From
        If Within(lstTo, X, Y) Then
            DragDefault
            DrawConnector = 0
            Set Cnr = Nothing
        End If
    Else
        DragDefault
        DrawConnector = 0
    End If
End Sub











Private Sub lstFrom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If DrawConnector = 2 Then 'if a connector is in the process of being connected
            'check for existing key
            If KeyExists(IndexXY(lstFrom, Y), 1) Then
                MsgBox "A field has already been mapped here", vbOKOnly + vbExclamation
                Set Cnr = Nothing
                DrawConnector = 0
                DragDefault
                Exit Sub
            End If
            Cnr.Index1 = IndexXY(lstFrom, Y)
            lstFrom.Selected(Cnr.Index1) = True
            If MapThem Then
                Lines.Add Cnr, CStr(hListIndex)
            End If
            Set Cnr = Nothing
            DrawConnector = 0
            DragDefault
        ElseIf IndexXY(lstFrom, Y) < lstFrom.ListCount Then
            'check for existing key
            If KeyExists(IndexXY(lstFrom, Y), 1) Then
                MsgBox "A field has already been mapped here", vbOKOnly + vbExclamation
                Set Cnr = Nothing 'Delete connector
                DrawConnector = 0
                DragDefault
                Exit Sub
            End If
            DragFrom 'Set cursor
            Dim Index As Integer
            Set Cnr = Nothing 'Reset Temp connector
            Set Cnr = New Connector 'Create a new connector
            'Draw the connector
            Cnr.Draw LineFrom, lstFrom.Top + yOffset + LineY * (lstFrom.ListIndex - lstFrom.TopIndex), _
                    CLng(X), CLng(Y), lstFrom.ListIndex, 0, NextColor
            DrawConnector = 1 'Current connector type (From-To)
        End If
    ElseIf Button = 2 Then 'Right-click to remove
        If KeyExists(lstFrom.ListIndex, 1) Then 'If a connection exists
            Dim i As Integer
            For i = Lines.Count To 1 Step -1 'iterate all connectors
                If Lines(i).Index1 = lstFrom.ListIndex Then
                    Lines(i).Visible = False
                    Lines.Remove i 'Remove it
                End If
            Next i
        End If
    End If
End Sub

'When the mouse moves over the list box you want the item under the mouse to be selected
Private Sub lstFrom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawConnector = 2 Then 'To - From
        DropHere 'change cursor
        If X < LineFrom Then
            X = CSng(LineFrom)
        End If
        
        lstFrom.Selected(IndexXY(lstFrom, Y)) = True
        Y = (lstFrom.ListIndex - lstFrom.TopIndex) * LineY + lstFrom.Top + yOffset
        If Y < lstFrom.Top Then Y = lstFrom.Top + yOffset
        
        'Draw the line
        Cnr.Draw LineTo, lstTo.Top + yOffset + LineY * (lstTo.ListIndex - lstTo.TopIndex), _
                CLng(X), CLng(Y)
    ElseIf DrawConnector = 0 Then
        If IndexXY(lstFrom, Y) < lstFrom.ListCount Then lstFrom.Selected(IndexXY(lstFrom, Y)) = True
    End If
End Sub

Private Sub lstFrom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawConnector = 2 Then
        'check for existing key
        If KeyExists(IndexXY(lstFrom, Y), 1) Then
            MsgBox "A field has already been mapped here", vbOKOnly + vbExclamation
            Set Cnr = Nothing
            DrawConnector = 0
            DragDefault
            Exit Sub
        End If
        Cnr.Index1 = IndexXY(lstFrom, Y)
        lstFrom.Selected(Cnr.Index1) = True
        If MapThem Then
            Lines.Add Cnr, CStr(hListIndex)
        End If
        Set Cnr = Nothing
        DrawConnector = 0
        DragDefault
    End If
End Sub

'On scroll redraw all the lines
Private Sub lstFrom_Scroll()
    Dim i As Integer
    For i = 1 To Lines.Count
        If Lines(i).FromX < Lines(i).ToX Then 'Originated from lstFrom
            'Check to see if line should go to top
            If Lines(i).Index1 < lstFrom.TopIndex Then
                Lines(i).Draw LineFrom, lstFrom.Top, Lines(i).ToX, Lines(i).ToY
            ElseIf Lines(i).Index1 >= lstFrom.TopIndex And Lines(i).Index1 <= lstFrom.TopIndex + intDisplayed Then
                Lines(i).Draw LineFrom, ((Lines(i).Index1 - lstFrom.TopIndex) * LineY + lstFrom.Top + yOffset), Lines(i).ToX, Lines(i).ToY
            ElseIf Lines(i).Index1 > lstFrom.TopIndex + intDisplayed Then
                Lines(i).Draw LineFrom, lstFrom.Top + lstFrom.Height, Lines(i).ToX, Lines(i).ToY
            End If
        Else
            'Check to see if line should go to top
            If Lines(i).Index1 < lstFrom.TopIndex Then
                Lines(i).Draw Lines(i).FromX, Lines(i).FromY, LineFrom, lstFrom.Top
            ElseIf Lines(i).Index1 >= lstFrom.TopIndex And Lines(i).Index1 <= lstFrom.TopIndex + intDisplayed Then
                Lines(i).Draw Lines(i).FromX, Lines(i).FromY, LineFrom, ((Lines(i).Index1 - lstFrom.TopIndex) * LineY + lstFrom.Top + yOffset)
            ElseIf Lines(i).Index1 > lstFrom.TopIndex + intDisplayed Then
                Lines(i).Draw Lines(i).FromX, Lines(i).FromY, LineFrom, lstFrom.Top + lstFrom.Height
            End If
        End If
    Next i
End Sub










Private Sub lstTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If DrawConnector = 1 Then
            'check for existing key
            If KeyExists(IndexXY(lstTo, Y), 2) Then
                MsgBox "A field has already been mapped here", vbOKOnly + vbExclamation
                Set Cnr = Nothing
                DrawConnector = 0
                DragDefault
                Exit Sub
            End If
            Cnr.Index2 = IndexXY(lstTo, Y)
            lstTo.Selected(Cnr.Index2) = True
            If MapThem Then
                Lines.Add Cnr, CStr(hListIndex)
            End If
            Set Cnr = Nothing
            DrawConnector = 0
            DragDefault
        ElseIf IndexXY(lstTo, Y) < lstTo.ListCount Then
            'check for existing key
            If KeyExists(IndexXY(lstTo, Y), 2) Then
                MsgBox "A field has already been mapped here", vbOKOnly + vbExclamation
                Set Cnr = Nothing
                DrawConnector = 0
                DragDefault
                Exit Sub
            End If
            DragTo
            Dim Index As Integer
            Set Cnr = Nothing
            Set Cnr = New Connector
            Cnr.Draw LineTo, lstTo.Top + yOffset + LineY * (lstTo.ListIndex - lstTo.TopIndex), _
                    CLng(X), CLng(Y), 0, lstTo.ListIndex, NextColor
            DrawConnector = 2
        End If
    ElseIf Button = 2 Then 'Right-Click to Remove
        If KeyExists(lstTo.ListIndex, 2) Then
            Dim i As Integer
            For i = Lines.Count To 1 Step -1
                If Lines(i).Index2 = lstTo.ListIndex Then
                    Lines(i).Visible = False
                    Lines.Remove i
                End If
            Next i
        End If
    End If
End Sub

Private Sub lstTo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawConnector = 1 Then 'From - To
        DropHere
        X = CSng(LineTo)
        lstTo.Selected(IndexXY(lstTo, Y)) = True
        Y = (lstTo.ListIndex - lstTo.TopIndex) * LineY + lstTo.Top + yOffset
        If Y < lstTo.Top Then Y = lstTo.Top + yOffset
        Cnr.Draw LineFrom, lstFrom.Top + yOffset + LineY * _
                (lstFrom.ListIndex - lstFrom.TopIndex), CLng(X), CLng(Y)
    ElseIf DrawConnector = 0 Then
        If IndexXY(lstTo, Y) < lstTo.ListCount Then lstTo.Selected(IndexXY(lstTo, Y)) = True
    End If
End Sub

Private Sub lstTo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawConnector = 1 Then
        'check for existing key
        If KeyExists(IndexXY(lstTo, Y), 2) Then
            MsgBox "A field has already been mapped here", vbOKOnly + vbExclamation
            Set Cnr = Nothing
            DrawConnector = 0
            DragDefault
            Exit Sub
        End If
        
        Cnr.Index2 = IndexXY(lstTo, Y)
        lstTo.Selected(Cnr.Index2) = True
        
        If MapThem Then
            Lines.Add Cnr, CStr(hListIndex)
        End If
        
        Set Cnr = Nothing
        DrawConnector = 0
        DragDefault
    End If
End Sub

Private Sub lstto_Scroll()
    Dim i As Integer
    For i = 1 To Lines.Count
        If Lines(i).FromX < Lines(i).ToX Then 'Originated from lstFrom
            'Check to see if line should go to top
            If Lines(i).Index2 < lstTo.TopIndex Then
                Lines(i).Draw Lines(i).FromX, Lines(i).FromY, LineTo, lstTo.Top
            ElseIf Lines(i).Index2 >= lstTo.TopIndex And Lines(i).Index2 <= lstTo.TopIndex + intDisplayed Then
                Lines(i).Draw Lines(i).FromX, Lines(i).FromY, LineTo, ((Lines(i).Index2 - lstTo.TopIndex) * LineY + lstTo.Top + yOffset)
            ElseIf Lines(i).Index2 > lstTo.TopIndex + intDisplayed Then
                Lines(i).Draw Lines(i).FromX, Lines(i).FromY, LineTo, lstTo.Top + lstTo.Height
            End If
        Else
            'Check to see if line should go to top
            If Lines(i).Index2 < lstTo.TopIndex Then
                Lines(i).Draw LineTo, lstTo.Top, Lines(i).ToX, Lines(i).ToY
            ElseIf Lines(i).Index2 >= lstTo.TopIndex And Lines(i).Index2 <= lstTo.TopIndex + intDisplayed Then
                Lines(i).Draw LineTo, ((Lines(i).Index2 - lstTo.TopIndex) * LineY + lstTo.Top + yOffset), Lines(i).ToX, Lines(i).ToY
            ElseIf Lines(i).Index2 > lstTo.TopIndex + intDisplayed Then
                Lines(i).Draw LineTo, lstTo.Top + lstTo.Height, Lines(i).ToX, Lines(i).ToY
            End If
        End If
    Next i
End Sub













'Used to determine whether or not a connector has
'already been associated with a field
Function KeyExists(KeyVal As Integer, WhichKey As Integer) As Boolean
    Dim i As Integer
    KeyExists = False
    For i = 1 To Lines.Count 'Process all existing connectors
        If WhichKey = 1 Then 'From list
            If Lines(i).Index1 = KeyVal Then 'Look for the KeyVal in Index1
                KeyExists = True 'Found it
                Exit For 'stop processing
            End If
        Else 'To list
            If Lines(i).Index2 = KeyVal Then 'Look for the KeyVal in Index2
                KeyExists = True 'Found it
                Exit For 'stop processing
            End If
        End If
    Next i
End Function

'Returns the index of the item that is at position Y from the top of the control
Function IndexXY(ctrl As Control, Y As Single) As Integer
    If Y < 0 Then
        IndexXY = 0
    Else
        'Convert Y to an Int and divide by the LineY constant
        'Drop the Fraction and add it to the TopIndex of the control
        IndexXY = Int((CInt(Y) / LineY)) + ctrl.TopIndex
        If IndexXY > ctrl.ListCount - 1 Then
            IndexXY = ctrl.ListCount - 1
        End If
    End If
End Function

'Returns True if coordinates X,Y are within the boundaries of ctrl
Function Within(ctrl As Control, X As Single, Y As Single) As Boolean
    Within = True
    If X < ctrl.Left Or X > ctrl.Left + ctrl.Width Then
        Within = False
    ElseIf Y < ctrl.Top Or Y > ctrl.Top + ctrl.Height Then
        Within = False
    End If
End Function

'Gets a random color
Public Function NextColor() As Long
    NextColor = CLng(8777215 * Rnd)
    Do Until NextColor <> frmMain.BackColor
        NextColor = CLng(8777215 * Rnd)
    Loop
End Function

'Returns the higher list index of the two lists
Public Function hListIndex() As Integer
    If lstTo.ListCount > lstFrom.ListCount Then
        hListIndex = lstTo.ListIndex
    Else
        hListIndex = lstFrom.ListIndex
    End If
End Function

'Removes all existing connectors
Public Sub ClearLines()
    Dim i As Integer
    For i = Lines.Count To 1 Step -1
        Lines(i).Visible = False
        Lines.Remove i
    Next i
End Sub

'Creates connectors for all fields with matching field names
Public Sub AutoMap()
    If chkAuto <> 1 Then Exit Sub
    
    Dim i As Integer, j As Integer, y1 As Long, y2 As Long
    For i = 0 To lstFrom.ListCount - 1 'iterate lstfrom
        For j = 0 To lstTo.ListCount - 1 'iterate lstto
            If Left(lstFrom.List(i), intFld + intType) = Left(lstTo.List(j), intFld + intType) Then
                
                'check for field size incompatibility
                Dim strTemp As String
                strTemp1 = Right(lstFrom.List(i), intLen) 'Get field size of lstFrom
                If strTemp1 <> "N/A" Then
                    Dim strTemp2 As String
                    strTemp2 = Right(lstTo.List(j), intLen) 'get field size of lstTo
                    If Val(strTemp1) > Val(strTemp2) Then
                        Exit For
                    End If
                End If
                
                'check for existing key
                If KeyExists(j, 2) Or KeyExists(i, 1) Then
                    Set Cnr = Nothing
                    Exit For
                End If
                Dim Index As Integer
                Set Cnr = Nothing
                Set Cnr = New Connector
                
                Cnr.Draw LineFrom, 0, LineTo, 0, i, j, NextColor
                If lstTo.ListCount > lstFrom.ListCount Then
                    Lines.Add Cnr, CStr(j)
                Else
                    Lines.Add Cnr, CStr(i)
                End If
                
                Set Cnr = Nothing
                lstto_Scroll
                lstFrom_Scroll
                Exit For
            End If
        Next j
    Next i
End Sub

Public Function MapThem() As Boolean
    'check for field incompatibility
    Dim fldType1 As String, fldSize1 As String
    Dim fldType2 As String, fldSize2 As String
    Dim intComp As Integer
    
    fldType1 = Trim(Mid(lstFrom.List(Cnr.Index1), intFld + 2, intType)) 'Get field type of lstFrom
    fldType2 = Trim(Mid(lstTo.List(Cnr.Index2), intFld + 2, intType)) 'Get field type of lstto
    
    'first check for field compatibility
    intComp = Compatible(fldType1, fldType2)
    If intComp = 1 Then
        fldSize1 = Trim(Right(lstFrom.List(Cnr.Index1), intLen))
        fldSize2 = Trim(Right(lstTo.List(Cnr.Index2), intLen))
        'Second check for length compatibility
        If LenCompatible(fldType1, fldType2, fldSize1, fldSize2) Then
            MapThem = True
        Else
            'Ask user if they want to lose data
            If MsgBox("The size of the field you are mapping to is smaller than the field you are mapping from. This could result in a loss of data upon transfer." & vbCrLf & vbCrLf & "Do you wish to make the mapping anyway?", vbYesNo + vbCritical) = vbYes Then
                MapThem = True
            Else
                MapThem = False
            End If
        End If
    ElseIf intComp = 2 Then
        If MsgBox("This mapping may result in a loss of data upon transfer because of the difference in field types." & vbCrLf & vbCrLf & "Do you wish to make the mapping anyway?", vbYesNo + vbCritical) = vbYes Then
            MapThem = True
        Else
            MapThem = False
        End If
    Else
        MsgBox "The fields you are trying to map are not of compatible types.", vbOKOnly + vbCritical
        MapThem = False
    End If
    
End Function

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub
