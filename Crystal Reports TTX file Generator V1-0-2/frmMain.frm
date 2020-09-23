VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "TTX File Creator for Crystal Reports"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDown 
      Caption         =   "DN"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4620
      TabIndex        =   22
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "UP"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4620
      TabIndex        =   21
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox txtSQL 
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6240
      Width           =   4455
   End
   Begin VB.CommandButton cmdCreateTTX 
      Caption         =   "Create .ttx File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   5760
      Width           =   4215
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "<-"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   16
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "->"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   15
      Top             =   3600
      Width           =   375
   End
   Begin VB.ListBox lstTTX 
      Enabled         =   0   'False
      Height          =   3180
      Left            =   2520
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
   End
   Begin VB.ListBox lstFields 
      Enabled         =   0   'False
      Height          =   3180
      Left            =   120
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelectTable 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   1620
      Width           =   1215
   End
   Begin VB.ComboBox cmbTables 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cmbDatabase 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelectDatabase 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify Server"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtSQLServer 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Fields in .ttx File"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Fields in Table"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4920
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label5 
      Caption         =   "Select Table:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      Caption         =   "Select Dabase:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Login:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Server:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSQLPW As String

Private Sub cmdDown_Click()
    Dim strData As String
    
    With Me
        If .lstTTX.ListIndex < .lstTTX.ListCount - 1 Then
            'move the darn thing up in list...
            strData = .lstTTX.List(.lstTTX.ListIndex + 1)
            .lstTTX.List(.lstTTX.ListIndex + 1) = .lstTTX.List(.lstTTX.ListIndex)
            .lstTTX.List(.lstTTX.ListIndex) = strData
            .lstTTX.ListIndex = .lstTTX.ListIndex + 1
        End If
    End With

End Sub

Private Sub cmdNo_Click()
    Dim indx As Integer
    
    If Me.lstTTX.SelCount <= 0 Then
        MsgBox "Please Select A Field To Remove."
    Else
        indx = 0
        Do While indx < Me.lstTTX.ListCount
            If Me.lstTTX.Selected(indx) Then
                Me.lstFields.AddItem Me.lstTTX.List(indx)
                Me.lstTTX.RemoveItem indx
            Else
                indx = indx + 1
            End If
        Loop
        If Me.lstTTX.ListCount <= 0 Then
            Me.cmdNo.Enabled = False
            Me.cmdDown.Enabled = False
            Me.cmdUp.Enabled = False
            Me.cmdCreateTTX.Enabled = False
        ElseIf Me.lstTTX.ListCount = 1 Then
            Me.cmdDown.Enabled = False
            Me.cmdUp.Enabled = False
        End If
    End If
End Sub

Private Sub cmdUp_Click()
    Dim strData As String
    
    With Me
        If .lstTTX.ListIndex <> 0 Then
            'move the darn thing up in list...
            strData = .lstTTX.List(.lstTTX.ListIndex - 1)
            .lstTTX.List(.lstTTX.ListIndex - 1) = .lstTTX.List(.lstTTX.ListIndex)
            .lstTTX.List(.lstTTX.ListIndex) = strData
            .lstTTX.ListIndex = .lstTTX.ListIndex - 1
        End If
    End With
End Sub

Private Sub cmdYes_Click()
    Dim indx As Integer
    
    If Me.lstFields.SelCount <= 0 Then
        MsgBox "Please Select A Field To Add."
    Else
        indx = 0
        Do While indx < Me.lstFields.ListCount
            If Me.lstFields.Selected(indx) Then
                Me.lstTTX.AddItem Me.lstFields.List(indx)
                Me.lstFields.RemoveItem indx
            Else
                indx = indx + 1
            End If
        Loop
        Me.cmdCreateTTX.Enabled = True
    End If
End Sub

Private Sub cmdSelectDatabase_Click()
    Dim adoCnn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    Dim strCnn As String
    Dim indx As Integer
    
    With Me
        .cmbDatabase.Enabled = False
        .cmdSelectDatabase.Enabled = False
        .cmbTables.Enabled = True
        .cmdSelectTable.Enabled = True
    End With
    'ok.. now load tables and go from there...
    strCnn = "Provider=SQLOLEDB.1;Password=" & Trim$(strSQLPW) & ";Persist Security Info=True;User ID=" & Trim$(Me.txtLogin.Text) & ";Initial Catalog=" & Trim$(Me.cmbDatabase.Text) & ";Data Source=" & Trim$(Me.txtSQLServer.Text)
    With adoCnn
        .ConnectionString = strCnn
        .ConnectionTimeout = 60
        .CursorLocation = adUseClient
        .Open
        Set adoRst = .OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
    End With
    With adoRst
        Do While Not .EOF
            If UCase$(Left$(!Table_Name, 4)) <> "MSYS" Then
                Me.cmbTables.AddItem !Table_Name
            End If
            .MoveNext
        Loop
    End With
    Me.cmbTables.ListIndex = 0
    If IsObject(adoRst) Then
        If adoRst.State = adStateOpen Then
            adoRst.Close
        End If
        Set adoRst = Nothing
    End If
    If IsObject(adoCnn) Then
        If adoCnn.State = adStateOpen Then
            adoCnn.Close
        End If
        Set adoCnn = Nothing
    End If
End Sub

Private Sub cmdSelectTable_Click()
    Dim indx As Integer
    Dim strCnn As String
    Dim adoCnn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    
    With Me
        .cmbTables.Enabled = False
        .cmdSelectTable.Enabled = False
        .lstFields.Enabled = True
        .lstTTX.Enabled = True
        .cmdYes.Enabled = True
    End With
    strCnn = "Provider=SQLOLEDB.1;Password=" & Trim$(strSQLPW) & ";Persist Security Info=True;User ID=" & Trim$(Me.txtLogin.Text) & ";Initial Catalog=" & Trim$(Me.cmbDatabase.Text) & ";Data Source=" & Trim$(Me.txtSQLServer.Text)
    With adoCnn
        .ConnectionString = strCnn
        .ConnectionTimeout = 60
        .CursorLocation = adUseClient
        .Open
        Set adoRst = .OpenSchema(adSchemaColumns, Array(Empty, Empty, UCase(Me.cmbTables.Text), Empty))
    End With
    With adoRst
        Do While Not .EOF
            If UCase$(Left$(!Table_Name, 4)) <> "MSYS" Then
                Me.lstFields.AddItem !Column_Name
            End If
            .MoveNext
        Loop
    End With
    
    If IsObject(adoRst) Then
        If adoRst.State = adStateOpen Then
            adoRst.Close
        End If
        Set adoRst = Nothing
    End If
    If IsObject(adoCnn) Then
        If adoCnn.State = adStateOpen Then
            adoCnn.Close
        End If
        Set adoCnn = Nothing
    End If
End Sub

Private Sub cmdVerify_Click()
    On Error GoTo ErrHandler
    Dim svrSQL As SQLDMO.SQLServer
    Dim indx As Integer
    
    If Trim$(Me.txtSQLServer.Text) <> vbNullString Then
        If Trim$(Me.txtLogin.Text) <> vbNullString Then
            strSQLPW = Trim$(Me.txtPassword.Text)
            Me.txtPassword.Text = vbNullString
            Set svrSQL = New SQLDMO.SQLServer
            With svrSQL
                On Error GoTo BadServer
                .Connect Trim$(Me.txtSQLServer.Text), Trim$(Me.txtLogin.Text), Trim$(strSQLPW)
                On Error GoTo ErrHandler
                For indx = 1 To .Databases.Count
                    Me.cmbDatabase.AddItem (.Databases(indx).Name)
                Next indx
                .DisConnect
            End With
            Me.cmbDatabase.ListIndex = 0
            Set svrSQL = Nothing
            With Me
                .txtLogin.Enabled = False
                .txtPassword.Enabled = False
                .txtSQLServer.Enabled = False
                .cmdVerify.Enabled = False
                .cmbDatabase.Enabled = True
                .cmdSelectDatabase.Enabled = True
            End With
        Else
            MsgBox "Please Enter A Valid User Name. Please Try Again."
            Me.txtLogin.SetFocus
        End If
    Else
        MsgBox "Please Enter A Valid SQL Server. Please Try Again."
        Me.txtSQLServer.SetFocus
    End If
    Exit Sub
BadServer:
    MsgBox "Can NOT Connect to " & Trim$(Me.txtSQLServer.Text) & ". Please Try Again."
    Me.txtSQLServer.SetFocus
    Err.Clear
    Set svrSQL = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description
    Me.txtSQLServer.SetFocus
    Err.Clear
    Set svrSQL = Nothing
End Sub

Private Sub lstFields_DblClick()
    cmdYes_Click
End Sub

Private Sub lstFields_GotFocus()
    Me.cmdNo.Enabled = False
    Me.cmdYes.Enabled = True
End Sub

Private Sub lstTTX_Click()
    Me.cmdDown.Enabled = True
    Me.cmdUp.Enabled = True
End Sub

Private Sub lstTTX_DblClick()
    cmdNo_Click
End Sub

Private Sub lstTTX_GotFocus()
    With Me
        .cmdNo.Enabled = True
        .cmdDown.Enabled = True
        .cmdUp.Enabled = True
        If .lstTTX.ListCount <= 0 Then
            .cmdNo.Enabled = False
            .cmdDown.Enabled = False
            .cmdUp.Enabled = False
            .cmdCreateTTX.Enabled = False
        ElseIf .lstTTX.ListCount = 1 Then
            .cmdDown.Enabled = False
            .cmdUp.Enabled = False
        End If
        .cmdYes.Enabled = False
    End With
End Sub

Private Sub cmdCreateTTX_Click()
    Dim intCh As Integer
    Dim indx As Integer
    Dim strCnn As String, strName As String, strType As String, strLength As String
    Dim adoCnn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    Dim strSQL As String
    
    strCnn = "Provider=SQLOLEDB.1;Password=" & Trim$(strSQLPW) & ";Persist Security Info=True;User ID=" & Trim$(Me.txtLogin.Text) & ";Initial Catalog=" & Trim$(Me.cmbDatabase.Text) & ";Data Source=" & Trim$(Me.txtSQLServer.Text)
    With adoCnn
        .ConnectionString = strCnn
        .ConnectionTimeout = 60
        .CursorLocation = adUseClient
        .Open
        Set adoRst = .OpenSchema(adSchemaColumns, Array(Empty, Empty, UCase(Me.cmbTables.Text), Empty))
    End With
    strSQL = "SELECT"
    intCh = FreeFile
    Open App.Path & "\" & Me.cmbDatabase.Text & "-" & Me.cmbTables.Text & ".ttx" For Output As #intCh
    Do While indx < Me.lstTTX.ListCount
        'process ttx record for this field...
        With adoRst
            Do While Not .EOF
                If UCase$(!Column_Name) = UCase$(Me.lstTTX.List(indx)) Then
                    'found column.. now gell about column....
                    strName = CStr(!Column_Name)
                    strLength = CStr(vbNullString & !CHARACTER_MAXIMUM_LENGTH)
                    strType = ConvType2CR(!Data_Type)
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
        'print information to .ttx file
        Print #intCh, strName & vbTab & strType & vbTab & strLength & vbTab
        If Trim$(strSQL) <> "SELECT" Then
            strSQL = strSQL & ","
        End If
        strSQL = strSQL & " " & Trim$(strName)
        indx = indx + 1
        adoRst.MoveFirst
    Loop
    Close #intCh
    strSQL = strSQL & " FROM " & Trim$(Me.cmbTables.Text)
    If IsObject(adoRst) Then
        If adoRst.State = adStateOpen Then
            adoRst.Close
        End If
        Set adoRst = Nothing
    End If
    If IsObject(adoCnn) Then
        If adoCnn.State = adStateOpen Then
            adoCnn.Close
        End If
        Set adoCnn = Nothing
    End If
    With Me
        .cmdCreateTTX.Enabled = False
        .lstFields.Clear
        .lstFields.Enabled = False
        .lstTTX.Clear
        .lstTTX.Enabled = True
        .cmbTables.Enabled = True
        .cmdSelectTable.Enabled = True
        .cmdSelectTable.SetFocus
        .cmdNo.Enabled = False
        .cmdYes.Enabled = False
        .cmdDown.Enabled = False
        .cmdUp.Enabled = False
        .txtSQL = strSQL
    End With
End Sub

Private Function ConvType2CR(ByVal TypeVal As Long) As String
    Select Case TypeVal
        Case adSmallInt                  ' 2
            ConvType2CR = "Short"
        Case adInteger                   ' 3
            ConvType2CR = "Long"
        Case adSingle                    ' 4
            ConvType2CR = "Number"
        Case adDouble                    ' 5
            ConvType2CR = "Number"
        Case adCurrency                  ' 6
            ConvType2CR = "Currency"
        Case adDate                      ' 7
            ConvType2CR = "Date"
        Case adBoolean                   ' 11
            ConvType2CR = "Boolean"
        Case adTinyInt                   ' 16
            ConvType2CR = "Short"
        Case adChar                      ' 129
            ConvType2CR = "String"
        Case adVarChar                   ' 200
            ConvType2CR = "String"
        Case Else
            ConvType2CR = "Unknown"
   End Select
End Function

