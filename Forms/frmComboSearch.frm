VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComboSearch 
   Caption         =   "MS Access-Style Combo Search Demonstration"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   Icon            =   "frmComboSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Standard Combo:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ADODataCombo:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmComboSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version:   21 July 2001
'
'Purpose:   Demonstrates an "MS Access way" of searching an ADO DataCombo and
'           a standard VB Combo box
'
'Author:    Peter Morgan (firefly_nz24@yahoo.com)
'Credits:   Matt Sendt (sendt@compassnet.com.au) - for the original combo search class
'
'References:
'           Microsoft ActiveX Data Objects 2.6 Library (any 2.x should work just as well)
'           Microsoft Data Binding Collectionn VB 6.0 (SP4)
'Components:
'           Microsoft DataList Controls 6.0 (SP3)(OLEDB)

Option Explicit

Dim objSearch1 As clsComboSearch    'combo search object
Dim objSearch2 As clsComboSearch    'combo search object
Dim ctl_current As Control          'last active combo (so you can tell which one was the last selected)


Private Sub Form_Load()
    Call PopulateBothComboBoxes         'populate the list data for each combo
    
    Set objSearch1 = New clsComboSearch 'set up the first search object
    Set objSearch1.Client = DataCombo1
    
    Set objSearch2 = New clsComboSearch 'set up the second search object
    Set objSearch2.Client = Combo1
End Sub

Private Sub PopulateBothComboBoxes()
'Set up the data to be displayed by both combo boxes
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\database.mdb"
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT UserID, FirstName & ' ' & LastName as UserName FROM tblUser ORDER BY FirstName", cn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        'populate the ADO datacombo by setting its properties
        With DataCombo1
            .Text = rs!UserName
            Set .RowSource = rs
            .DataField = "UserID"
            .ListField = "UserName"
        End With
        'populate the standard combo using this ADO recordset
        Combo1.Text = rs!UserName
        While Not rs.EOF
            Combo1.AddItem rs!UserName
            Combo1.ItemData(Combo1.NewIndex) = rs!UserID
            rs.MoveNext
        Wend
        rs.MoveFirst
    End If
    
Exit_Sub:
    Set rs = Nothing
    Set cn = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub

Private Sub Combo1_GotFocus()
'NB: You do not need this if you have only one searchable combo on the form
    Set ctl_current = Combo1        'Combo1 is the current control
End Sub

Private Sub DataCombo1_GotFocus()
'NB: You do not need this if you have only one searchable combo on the form
    Set ctl_current = DataCombo1    'DataCombo1 is the current control
End Sub

Private Sub cmdSelect_Click()
'Display the combo text (of the current combo box)
    If ctl_current.Name = "DataCombo1" Or ctl_current.Name = "Combo1" Then
        MsgBox "Control: " & ctl_current.Name & " Value: " & ctl_current.Text, vbInformation, "Combo Result"
    End If
End Sub
