VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Explorer 
   Caption         =   "TreeView Demo"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMover 
      BackColor       =   &H80000015&
      Height          =   4725
      Left            =   4200
      ScaleHeight     =   4665
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Explorer.frx":0000
            Key             =   "category"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Explorer.frx":0452
            Key             =   "main"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Explorer.frx":08A4
            Key             =   "component"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Explorer.frx":0BBE
            Key             =   "equipment"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVertical 
      Height          =   4725
      Left            =   3480
      ScaleHeight     =   4665
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   0
      Width           =   100
   End
   Begin MSDataGridLib.DataGrid DBGridExplorer 
      Height          =   4725
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8334
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwExplorer 
      Height          =   4725
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   8334
      _Version        =   393217
      Indentation     =   2
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "Expand Tree"
         Index           =   0
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "Search Tree"
         Index           =   1
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frm_Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Action_Type
  Expand = 0 'Expand the tree
  Search = 1 'Search the tree
End Enum

Private bln_search_Mode As Boolean
Private bln_Escape_Mode As Boolean 'If ESC key is pressed, cancel actions
Private bln_dragging As Boolean 'For splitter operation
Private lng_X As Single 'For splitter operation
Private lng_Y As Single 'For splitter operation
Private node_ctl As Node 'for search mode
Private str As String
Private int_Counter  As Integer 'number of successful finds

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ET
  If KeyAscii = 27 Then 'ESC key
    If bln_search_Mode Then
      bln_Escape_Mode = True
    Else
      bln_Escape_Mode = True
    End If
  End If
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo ET
  'Initialize position of splitter
  picMover.Width = 5
  picVertical.Left = tvwExplorer.Width
  picMover.Left = picVertical.Left
  With DBGridExplorer
    .Move picVertical.Left + picVertical.Width, .Top, Me.ScaleWidth - (picVertical.Left + picVertical.Width)
  End With
  'Populate treeview
  PopulateTop_Level tvwExplorer
  'Initiate a node click
  tvwExplorer_NodeClick tvwExplorer.Nodes(1)
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

Private Sub Form_Resize()
On Error Resume Next
  tvwExplorer.Height = ScaleHeight
  picMover.Height = ScaleHeight
  picVertical.Height = ScaleHeight
  With DBGridExplorer
    .Move picVertical.Left + picVertical.Width, .Top, ScaleWidth - (picVertical.Left + picVertical.Width), ScaleHeight
  End With
End Sub

Private Sub DBGridExplorer_KeyPress(KeyAscii As Integer)
On Error GoTo ET
  If KeyAscii = 27 Then 'ESC key
    If bln_search_Mode Then
      bln_Escape_Mode = True
    Else
      bln_Escape_Mode = True
    End If
  End If
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
On Error GoTo ET
  Select Case Index
    Case 0
      Perform_Search Expand 'Expand the tree
    Case 1
      str = InputBox("Type the words you want to search.", "Search Tree")
      If str <> "" Then
        int_Counter = 0
        Perform_Search Search, str 'expand and search the tree
      End If
    Case 3
      Unload Me
  End Select
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

Private Sub picVertical_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bln_dragging = True
End Sub

Private Sub picVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ET
  picVertical.MousePointer = 9
  If Button = vbLeftButton Then
    picMover.Visible = True
    With picMover
      .Move tvwExplorer.Width + X  ' * Screen.TwipsPerPixelX
    End With
    bln_dragging = False
  End If
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

Private Sub picVertical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ET
  picMover.Visible = False
  picVertical.Left = picMover.Left
  'picMover.Left = picVertical.Left
  tvwExplorer.Width = picVertical.Left
  With DBGridExplorer
    .Move tvwExplorer.Width + picVertical.Width, .Top, ScaleWidth - (tvwExplorer.Width + picVertical.Width), ScaleHeight
  End With
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

Private Sub tvwExplorer_BeforeLabelEdit(Cancel As Integer)
    'Cancel the edit operation
    Cancel = True
End Sub

Private Sub tvwExplorer_KeyPress(KeyAscii As Integer)
On Error GoTo ET
  If KeyAscii = 27 Then 'ESC key
    If bln_search_Mode Then
      bln_Escape_Mode = True
    Else
      bln_Escape_Mode = True
    End If
  End If
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

'Sub to perform expand or search of tree
Private Sub Perform_Search(ByVal Actions As Action_Type, Optional ByVal str_Criteria As String)
On Error GoTo ET
Dim str As String
  str = Me.Caption
  Me.Caption = "Expanding Data Tree....Press ESC to terminate."
  bln_search_Mode = True
  Search_tree node_ctl, Actions, str_Criteria
  Me.Caption = str
  bln_search_Mode = False
  bln_Escape_Mode = False
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub

'sub to expand and searching of  the tree
Private Sub Search_tree(ByRef Node As Node, ByVal Action_Type As Action_Type, Optional ByVal str_Criteria As String)
On Error Resume Next
Dim mnode As Node
Dim nodechild As Node

  If bln_Escape_Mode Then Exit Sub

  Set nodechild = Node.Child
  If Action_Type = Search Then
    If Trim(LCase(nodechild.Text)) = LCase(str_Criteria) Then
      nodechild.Bold = True
      nodechild.ForeColor = vbRed
      int_Counter = int_Counter + 1
      If MsgBox(nodechild.Text & " found at " & nodechild.FullPath, vbInformation + vbYesNo, int_Counter & " Item Found!") <> vbYes Then
        bln_Escape_Mode = True
        Exit Sub
      End If
    End If
  End If
  Do While Not (nodechild Is Nothing)
    tvwExplorer_NodeClick nodechild
    DoEvents
    If nodechild.Children <> 0 Then
      nodechild.Expanded = True
      If bln_Escape_Mode Then Exit Sub
      If Action_Type = Search Then
          If Trim(LCase(nodechild.Text)) = LCase(str_Criteria) Then
          nodechild.Bold = True
          nodechild.ForeColor = vbRed
          int_Counter = int_Counter + 1
          If MsgBox(nodechild.Text & " found at " & nodechild.FullPath, vbInformation + vbYesNo, int_Counter & " Item Found!") <> vbYes Then
            bln_Escape_Mode = True
            Exit Sub
          End If
        End If
      End If
      Search_tree nodechild, Action_Type, str_Criteria
    Else
      If bln_Escape_Mode Then Exit Sub
    End If
    Set nodechild = nodechild.Next
  Loop
  
End Sub

Private Sub tvwExplorer_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ET
Dim rst As ADODB.Recordset
  Set node_ctl = Node
  'lock the controls to prevent flicker
  Me.AutoRedraw = False
  'Populate nodes
  If bln_search_Mode Then
    Set rst = New ADODB.Recordset
    Set rst = CheckChildrenAndPopulate(tvwExplorer, Node)
    Set rst = Nothing
    Me.AutoRedraw = True
    Exit Sub
  Else
    With DBGridExplorer
      Set .DataSource = CheckChildrenAndPopulate(tvwExplorer, Node)
      Select Case Node.Tag
        Case "Group"
          .Columns("Category_ID").Visible = False
        Case Else
          .Columns("Equipment_ID").Visible = False
          .Columns("Category_ID").Visible = False
          .Columns("Equipment_Code").Visible = False
      End Select
    End With
  End If
  If Node.Children = 0 Then
    'Get data for the node alone
    With DBGridExplorer
      Set .DataSource = GetDetails_Specific(Node)
      Select Case Node.Tag
        Case "Group"
          'Do nothing
        Case "category"
          .Columns("Category_ID").Visible = False
        Case Else
          .Columns("Equipment_ID").Visible = False
          .Columns("Category_ID").Visible = False
          .Columns("Equipment_Code").Visible = False
      End Select
    End With
  End If
  Node.Expanded = True
  Me.AutoRedraw = True
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, " Error No " & Err.Number
  Exit Sub
End Sub
