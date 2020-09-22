VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'A connector consists of 7 lines
'2 lines for the '>' wedge next to the From list (lnEndTo)
'2 lines for the '<' wedge next to the To list   (lnEndFrom)
'1 straight line '-----' next to the from wedge  (lnFrom)
'1 straight line '-----' next to the to wedge    (lnTo)
'1 diagonal line connecting the straight lines   (lnDiag)
'                  ------<
'                 /
'                /
'               /
'              /
'             /
'       >-----

'Connector End; consists of 2 lines
Private Type ConEnd
    lnUp As Line
    lnDown As Line
    lnUpName As String
    lnDownName As String
End Type
Private lnEndTo As ConEnd '> wedge
Private lnEndFrom As ConEnd '< wedge
Private lnTo As Line '----- line
Private lnFrom As Line '----- line
Private lnDiag As Line 'Diagonal Line
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Keep a reference of the names of the lines so that
'you can remove them when the connector is destroyed
Private lnToName As String
Private lnFromName As String
Private lnDiagName As String

'Keeps up with the current x,y coordinates of the two ends
'of the connector; used mainly for scrolling the lists
Public FromX As Long
Public ToX As Long
Public FromY As Long
Public ToY As Long

Private ParentForm As Form 'reference to the parent form
Private myColor As Long 'Color of connector
Public Index1 As Integer 'Index of item in lstFrom
Public Index2 As Integer 'Index of item in lstTo



'Set the visible property
Property Let Visible(bln1 As Boolean)
    If Not bln1 Then
        lnTo.Visible = False
        lnDiag.Visible = False
        lnFrom.Visible = False
        lnEndTo.lnDown.Visible = False
        lnEndTo.lnUp.Visible = False
        lnEndFrom.lnDown.Visible = False
        lnEndFrom.lnUp.Visible = False
    Else
        lnTo.Visible = True
        lnDiag.Visible = True
        lnFrom.Visible = True
        lnEndTo.lnDown.Visible = True
        lnEndTo.lnUp.Visible = True
        lnEndFrom.lnDown.Visible = True
        lnEndFrom.lnUp.Visible = True
    End If
End Property

'Define the coordinates of the connector
Public Sub Draw(fromX1 As Long, fromY1 As Long, toX2 As Long, toY2 As Long, Optional IndexA As Integer = 0, Optional IndexB As Integer = 0, Optional lngColor As Long = 0)
    FromX = fromX1 'Save values so that they can be referenced
    FromY = fromY1 'These values are needed for scrolling
    ToX = toX2
    ToY = toY2
    
    If lngColor <> 0 Then
        myColor = lngColor
    End If
    
    If IndexA <> 0 Then Index1 = IndexA 'Set reference to Index
    If IndexB <> 0 Then Index2 = IndexB
    
    Dim dif As Long, LeftOffset As Long, RightOffset As Long
    dif = Abs(fromX1 - toX2)
    LeftOffset = dif / 4
    RightOffset = dif - LeftOffset * 3
    
    If fromX1 < toX2 Then 'Start from Left
        lnFrom.X1 = fromX1
        lnFrom.y1 = fromY1
        lnFrom.X2 = fromX1 + LeftOffset
        lnFrom.y2 = fromY1
        lnFrom.BorderColor = myColor
        
        lnTo.X1 = toX2 - RightOffset
        lnTo.y1 = toY2
        lnTo.X2 = toX2
        lnTo.y2 = toY2
        lnTo.BorderColor = myColor
        
        lnDiag.X1 = lnFrom.X2
        lnDiag.y1 = lnFrom.y2
        lnDiag.X2 = lnTo.X1
        lnDiag.y2 = lnTo.y1
        lnDiag.BorderColor = myColor
        
        lnEndTo.lnUp.X1 = toX2
        lnEndTo.lnUp.y1 = toY2 - yOffset / 2
        lnEndTo.lnUp.X2 = toX2 - 50
        lnEndTo.lnUp.y2 = toY2
        lnEndTo.lnUp.BorderColor = myColor
        
        lnEndTo.lnDown.X1 = toX2
        lnEndTo.lnDown.y1 = toY2 + yOffset / 2
        lnEndTo.lnDown.X2 = toX2 - 50
        lnEndTo.lnDown.y2 = toY2
        lnEndTo.lnDown.BorderColor = myColor
        
        lnEndFrom.lnUp.X1 = fromX1
        lnEndFrom.lnUp.y1 = fromY1 - yOffset / 2
        lnEndFrom.lnUp.X2 = fromX1 + 50
        lnEndFrom.lnUp.y2 = fromY1
        lnEndFrom.lnUp.BorderColor = myColor
        
        lnEndFrom.lnDown.X1 = fromX1
        lnEndFrom.lnDown.y1 = fromY1 + yOffset / 2
        lnEndFrom.lnDown.X2 = fromX1 + 50
        lnEndFrom.lnDown.y2 = fromY1
        lnEndFrom.lnDown.BorderColor = myColor
        
    ElseIf fromX1 > toX2 Then 'Start from Right
        lnFrom.X1 = fromX1
        lnFrom.y1 = fromY1
        lnFrom.X2 = fromX1 - RightOffset
        lnFrom.y2 = fromY1
        lnFrom.BorderColor = myColor
        
        lnTo.X1 = toX2 + LeftOffset
        lnTo.y1 = toY2
        lnTo.X2 = toX2
        lnTo.y2 = toY2
        lnTo.BorderColor = myColor
        
        lnDiag.X1 = lnFrom.X2
        lnDiag.y1 = lnFrom.y2
        lnDiag.X2 = lnTo.X1
        lnDiag.y2 = lnTo.y1
        lnDiag.BorderColor = myColor
        
        lnEndTo.lnUp.X1 = toX2
        lnEndTo.lnUp.y1 = toY2 - yOffset / 2
        lnEndTo.lnUp.X2 = toX2 + 50
        lnEndTo.lnUp.y2 = toY2
        lnEndTo.lnUp.BorderColor = myColor
        
        lnEndTo.lnDown.X1 = toX2
        lnEndTo.lnDown.y1 = toY2 + yOffset / 2
        lnEndTo.lnDown.X2 = toX2 + 50
        lnEndTo.lnDown.y2 = toY2
        lnEndTo.lnDown.BorderColor = myColor
       
        lnEndFrom.lnUp.X1 = fromX1
        lnEndFrom.lnUp.y1 = fromY1 - yOffset / 2
        lnEndFrom.lnUp.X2 = fromX1 - 50
        lnEndFrom.lnUp.y2 = fromY1
        lnEndFrom.lnUp.BorderColor = myColor
        
        lnEndFrom.lnDown.X1 = fromX1
        lnEndFrom.lnDown.y1 = fromY1 + yOffset / 2
        lnEndFrom.lnDown.X2 = fromX1 - 50
        lnEndFrom.lnDown.y2 = fromY1
        lnEndFrom.lnDown.BorderColor = myColor
    End If
End Sub

Private Sub Class_Initialize()
    
    Set ParentForm = frmMain 'Set reference to Main Form
    LineNo = LineNo + 1 'Used for unique line number
    'Create a line object
    ParentForm.Controls.Add "VB.Line", "Line" & Trim(Str(LineNo))
    'Make it visible
    ParentForm.Controls("Line" & Trim(Str(LineNo))).Visible = True
    'Get reference to the new line
    Set lnFrom = ParentForm.Controls("Line" & Trim(Str(LineNo)))
    'Save name of the line so that it can be deleted later
    lnFromName = "Line" & Trim(Str(LineNo))
    'Set Borderwidth
    lnFrom.BorderWidth = 2
    
    LineNo = LineNo + 1
    ParentForm.Controls.Add "VB.Line", "Line" & Trim(Str(LineNo))
    ParentForm.Controls("Line" & Trim(Str(LineNo))).Visible = True
    Set lnTo = ParentForm.Controls("Line" & Trim(Str(LineNo)))
    lnToName = "Line" & Trim(Str(LineNo))
    lnTo.BorderWidth = 2
    
    LineNo = LineNo + 1
    ParentForm.Controls.Add "VB.Line", "Line" & Trim(Str(LineNo))
    ParentForm.Controls("Line" & Trim(Str(LineNo))).Visible = True
    Set lnDiag = ParentForm.Controls("Line" & Trim(Str(LineNo)))
    lnDiagName = "Line" & Trim(Str(LineNo))
    lnDiag.BorderWidth = 2
    
    LineNo = LineNo + 1
    ParentForm.Controls.Add "VB.Line", "Line" & Trim(Str(LineNo))
    ParentForm.Controls("Line" & Trim(Str(LineNo))).Visible = True
    Set lnEndTo.lnDown = ParentForm.Controls("Line" & Trim(Str(LineNo)))
    lnEndTo.lnDownName = "Line" & Trim(Str(LineNo))
    lnEndTo.lnDown.BorderWidth = 2
    
    LineNo = LineNo + 1
    ParentForm.Controls.Add "VB.Line", "Line" & Trim(Str(LineNo))
    ParentForm.Controls("Line" & Trim(Str(LineNo))).Visible = True
    Set lnEndTo.lnUp = ParentForm.Controls("Line" & Trim(Str(LineNo)))
    lnEndTo.lnUpName = "Line" & Trim(Str(LineNo))
    lnEndTo.lnUp.BorderWidth = 2
    
    LineNo = LineNo + 1
    ParentForm.Controls.Add "VB.Line", "Line" & Trim(Str(LineNo))
    ParentForm.Controls("Line" & Trim(Str(LineNo))).Visible = True
    Set lnEndFrom.lnDown = ParentForm.Controls("Line" & Trim(Str(LineNo)))
    lnEndFrom.lnDownName = "Line" & Trim(Str(LineNo))
    lnEndFrom.lnDown.BorderWidth = 2
    
    LineNo = LineNo + 1
    ParentForm.Controls.Add "VB.Line", "Line" & Trim(Str(LineNo))
    ParentForm.Controls("Line" & Trim(Str(LineNo))).Visible = True
    Set lnEndFrom.lnUp = ParentForm.Controls("Line" & Trim(Str(LineNo)))
    lnEndFrom.lnUpName = "Line" & Trim(Str(LineNo))
    lnEndFrom.lnUp.BorderWidth = 2
    
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    'Clean up your mess
    ParentForm.Controls.Remove lnDiagName
    ParentForm.Controls.Remove lnFromName
    ParentForm.Controls.Remove lnToName
    ParentForm.Controls.Remove lnEndTo.lnDownName
    ParentForm.Controls.Remove lnEndTo.lnUpName
    ParentForm.Controls.Remove lnEndFrom.lnDownName
    ParentForm.Controls.Remove lnEndFrom.lnUpName
ErrorHandler:
    If Err = 730 Or Err = 365 Then Resume Next 'Happens on last connector that is open at program termination
End Sub
