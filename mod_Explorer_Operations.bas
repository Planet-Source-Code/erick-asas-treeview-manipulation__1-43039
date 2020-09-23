Attribute VB_Name = "mod_Explorer_Operations"
Option Explicit
Private Conn As ADODB.Connection

Public Function Establish_Connection() As ADODB.Connection

  On Error GoTo ET

  'Connect to database
  Set Conn = New ADODB.Connection
  Conn.CursorLocation = adUseClient
  Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & _
          App.Path & "\TreeView.mdb;"

  Set Establish_Connection = Conn

  Exit Function

ET:
  Err.Raise Err.Number, Err.Description

    Exit Function

End Function
Public Function CheckChildrenAndPopulate(ByRef tvw As TreeView, ByRef mnode As Node) As ADODB.Recordset
On Error GoTo ET
Dim ConnSql As ADODB.Connection
Dim rstSQL As ADODB.Recordset
Dim rst As ADODB.Recordset
      Select Case mnode.Tag
        Case "Group"
          Set rst = Populate_Category(tvw, mnode)
        Case "category"
          Set rst = Populate_Equipment(tvw, mnode)
        Case "equipment", "component"
          Set rst = Populate_Component(tvw, mnode)
      End Select
  Set CheckChildrenAndPopulate = rst
  Set rst = Nothing
  Set rstSQL = Nothing
  Exit Function
ET:
  Err.Raise Err.Number, Err.Description
  Exit Function
End Function

Public Sub PopulateTop_Level(ByRef tvw As TreeView)
On Error GoTo ET
  Dim boolPresent  As Boolean
  Dim m_node As Node
  
    Set m_node = tvw.Nodes.Add
    m_node.Text = "Properties, Assets & Liabilities"
    m_node.Tag = "Group"
    m_node.Image = "main"
    Exit Sub
ET:
  Err.Raise Err.Number, Err.Description
  Exit Sub
End Sub


Public Function Populate_Category(ByRef tvw As TreeView, ByRef mnode As Node) As ADODB.Recordset
On Error GoTo ET
Dim rstSQL As ADODB.Recordset 'recordset for asset Cost Center
Dim rst As ADODB.Recordset
Dim intIndexcostcenters As Integer 'index for cost center
Dim m_node As Node
Dim txtKey As String
Dim boolPresent  As Boolean
Dim nodechild As Node
Dim strsql As String
Dim ConnSql As ADODB.Connection
  
  Set ConnSql = New ADODB.Connection
  Set ConnSql = Establish_Connection()

    'Initialize recordset
    Set rstSQL = New ADODB.Recordset
    'Define SQL string for cost centers
    strsql = "SELECT * FROM Main_Assets_Category"
    'Execute SQL
    Set rstSQL = ConnSql.Execute(strsql)
    'Check for additional children
    strsql = "SELECT count(*) as Recnum FROM Main_Assets_Category"
    Set rst = ConnSql.Execute(strsql)
    If rst("RecNum") <> mnode.Children Then
      If rstSQL.EOF And rstSQL.BOF Then 'do nothing
      Else
        intIndexcostcenters = mnode.Index
        Do While Not rstSQL.EOF
          'Initialize key
          txtKey = rstSQL("Category_ID") & "IDCategory"
          'Check if already present in the node collection
          Set nodechild = mnode.Child
          boolPresent = False
          Do While Not (nodechild Is Nothing)
            If nodechild.Key = txtKey Then
              boolPresent = True
              Exit Do
            End If
            Set nodechild = nodechild.Next
          Loop
          If Not boolPresent Then
            'add node
            Set m_node = tvw.Nodes.Add(intIndexcostcenters, tvwChild)
            m_node.Key = txtKey
            m_node.Text = rstSQL("Description")
            m_node.Tag = "category"
            m_node.Image = "category"
          End If
          rstSQL.MoveNext
        Loop
      End If
    End If
    Set Populate_Category = rstSQL
    Set rstSQL = Nothing
    Set rst = Nothing
    Exit Function
ET:
  Err.Raise Err.Number, Err.Description
  Exit Function
End Function

Public Function Populate_Equipment(ByRef tvw As TreeView, ByRef mnode As Node) As ADODB.Recordset
On Error GoTo ET
Dim rstSQL As ADODB.Recordset 'recordset for asset Cost Center
Dim rst As ADODB.Recordset
Dim intIndexcostcenters As Integer 'index for cost center
Dim m_node As Node
Dim txtKey As String
Dim boolPresent  As Boolean
Dim nodechild As Node
Dim strsql As String
Dim ConnSql As ADODB.Connection
  
  Set ConnSql = New ADODB.Connection
  Set ConnSql = Establish_Connection

    'Initialize recordset
    Set rstSQL = New ADODB.Recordset
    'Define SQL string for cost centers
    'Check if member of the OSP Group
    strsql = "SELECT * FROM Main_Assets_Equipment Where Category_ID = " & val(mnode.Key) & " And Equipment_Code = 0"
    'Execute SQL
    Set rstSQL = ConnSql.Execute(strsql)
    'Check for additional children
    strsql = "SELECT count(*) as Recnum From Main_Assets_Equipment Where Category_ID = " & val(mnode.Key) & " And Equipment_Code = 0"
    Set rst = ConnSql.Execute(strsql)
    If rst("RecNum") <> mnode.Children Then
      If rstSQL.EOF And rstSQL.BOF Then 'do nothing
      Else
        intIndexcostcenters = mnode.Index
        Do While Not rstSQL.EOF
          'Initialize key
          txtKey = rstSQL("Equipment_ID") & "IDEquipment"
          'Check if already present in the node collection
          Set nodechild = mnode.Child
          boolPresent = False
          Do While Not (nodechild Is Nothing)
            If nodechild.Key = txtKey Then
              boolPresent = True
              Exit Do
            End If
            Set nodechild = nodechild.Next
          Loop
          If Not boolPresent Then
            'add node
            Set m_node = tvw.Nodes.Add(intIndexcostcenters, tvwChild)
            m_node.Key = txtKey
            m_node.Text = rstSQL("Component_Code")
            m_node.Tag = "equipment"
            m_node.Image = "equipment"
          End If
          rstSQL.MoveNext
        Loop
      End If
    End If
    Set Populate_Equipment = rstSQL
    Set rstSQL = Nothing
    Set rst = Nothing
    Exit Function
ET:
  Err.Raise Err.Number, Err.Description
  Exit Function
End Function

Public Function Populate_Component(ByRef tvw As TreeView, ByRef mnode As Node) As ADODB.Recordset
On Error GoTo ET
Dim rstSQL As ADODB.Recordset 'recordset for asset Cost Center
Dim rst As ADODB.Recordset
Dim intIndexcostcenters As Integer 'index for cost center
Dim m_node As Node
Dim txtKey As String
Dim boolPresent  As Boolean
Dim nodechild As Node
Dim strsql As String
Dim ConnSql As ADODB.Connection
  
  Set ConnSql = New ADODB.Connection
  Set ConnSql = Establish_Connection

    'Initialize recordset
    Set rstSQL = New ADODB.Recordset
    'Define SQL string for cost centers
    strsql = "SELECT * FROM Main_Assets_Equipment Where Category_ID = " & Category_ID(mnode) & " And Equipment_Code = " & val(mnode.Key)
    'Execute SQL
    Set rstSQL = ConnSql.Execute(strsql)
    'Check for additional children
    strsql = "SELECT count(*) as Recnum FROM Main_Assets_Equipment Where Category_ID = " & Category_ID(mnode) & " And Equipment_Code = " & val(mnode.Key)
    Set rst = ConnSql.Execute(strsql)
    If rst("RecNum") <> mnode.Children Then
      If rstSQL.EOF And rstSQL.BOF Then 'do nothing
      Else
        intIndexcostcenters = mnode.Index
        Do While Not rstSQL.EOF
          'Initialize key
          txtKey = rstSQL("Equipment_ID") & "IDEquipment"
          'Check if already present in the node collection
          Set nodechild = mnode.Child
          boolPresent = False
          Do While Not (nodechild Is Nothing)
            If nodechild.Key = txtKey Then
              boolPresent = True
              Exit Do
            End If
            Set nodechild = nodechild.Next
          Loop
          If Not boolPresent Then
            'add node
            Set m_node = tvw.Nodes.Add(intIndexcostcenters, tvwChild)
            m_node.Key = txtKey
            m_node.Text = rstSQL("Component_Code")
            m_node.Tag = "equipment"
            m_node.Image = "equipment"
          End If
          rstSQL.MoveNext
        Loop
      End If
    End If
    Set Populate_Component = rstSQL
    Set rstSQL = Nothing
    Set rst = Nothing
    Exit Function
ET:
  Err.Raise Err.Number, Err.Description
  Exit Function
End Function

Private Function Category_ID(ByRef mnode As Node) As Long
On Error GoTo ET
  Dim m_node As Node
  Set m_node = mnode
  Dim blnFound As Boolean
    Do While Not blnFound
      If m_node.Tag = "category" Then
        Category_ID = val(m_node.Key)
        Exit Function
      Else
        Set m_node = m_node.Parent
      End If
    Loop
    If Not blnFound Then
      Category_ID = val(mnode.Key)
    End If
  Exit Function
ET:
  Err.Raise Err.Number, Err.Description
  Exit Function
End Function

Public Function GetDetails_Specific(ByRef mnode As Node) As ADODB.Recordset
On Error GoTo ET
Dim ConnSql As ADODB.Connection
Dim rstSQL As ADODB.Recordset
Dim strsql As String
  
  Select Case mnode.Tag
    Case "Group"
      'do nothing
    Case "category"
      strsql = "SELECT * FROM Main_Assets_Category Where Category_ID = " & val(mnode.Key)
    Case "equipment", "component"
      strsql = "SELECT * FROM Main_Assets_Equipment Where Equipment_ID = " & val(mnode.Key)
  End Select
  Set ConnSql = New ADODB.Connection
  Set ConnSql = Establish_Connection
  Set rstSQL = New ADODB.Recordset
  Set rstSQL = ConnSql.Execute(strsql)
  Set GetDetails_Specific = rstSQL
  Set rstSQL = Nothing
  Exit Function
ET:
  Err.Raise Err.Number, Err.Description
  Exit Function
End Function

