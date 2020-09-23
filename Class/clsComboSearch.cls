VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version:   21 July 2001
'
'Purpose:   Class for searching an ADO DataCombo or a standard VB Combo box in
'           an MS Access style
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

Private WithEvents m_DCombo As DataCombo
Attribute m_DCombo.VB_VarHelpID = -1
Private WithEvents m_cbo As ComboBox
Attribute m_cbo.VB_VarHelpID = -1
Private m_RST As ADODB.Recordset    'Holds a clone of the recordset
Private m_FIELD As Field       'Holds the field within the recordset that comprises the ListField of the DataCombo
Private m_STOP As Boolean      'This variable determines whether to start a new search or resume the last search
Private m_FIELD_NAME As String 'Holds the name of the field specified in m_FIELD

Private blnAuto As Boolean     'Keeps the autocomplete functions from triggering the
                               'Change event. Used with ComboBox only
'################################################################################################################################################################
'#  Description                                                                                                                                                 #
'#______________________________________________________________________________________________________________________________________________________________#
'#  This class is designed to give the feel of an Access ComboBox to the New DataCombo control in VB6.                                                          #
'#  Whilst this class was designed by Matt Sendt February 2000, you are free to use this Class module or modify it as long as you do not claim it as your       #
'#  intellectual property. This version is only a beta version and will probably require much more error handling, etc.                                         #
'#  Please contact me at sendt@compassnet.com.au if you have any improvements.                                                                                  #
'################################################################################################################################################################
'#  Example usage                                                                                                                                               #
'#______________________________________________________________________________________________________________________________________________________________#
'#  This class is typically used in a form_load event                                                                                                           #
'#  In the declarations section of the form:                                                                                                                    #
'#  Private objSearch As clsComboSearch                                                                                                                            #
'#  In the form_load event:                                                                                                                                     #
'#  Set objSearch = New clsComboSearch                                                                                                                             #
'#  Set objSearch.Client = DataCombo1   'Replace the word "DataCombo1" with the name of your DataCombo                                                            #
'#                                                                                                                                                              #
'################################################################################################################################################################

Private Sub Class_Initialize()
    'give all objects & variables an initial value
    Set m_DCombo = Nothing
    Set m_RST = Nothing
    Set m_FIELD = Nothing
    m_FIELD_NAME = ""
End Sub

Private Sub Class_Terminate()
    Set m_DCombo = Nothing    'release objects
    Set m_RST = Nothing
    Set m_FIELD = Nothing
End Sub

Public Property Set Client(cbo As Object)
'Set the client to either an ADO DataCombo or a standard ComboBox
    If TypeName(cbo) = "DataCombo" Then   'Check that we are initialising an ADO Datacombo
        Set m_DCombo = cbo
        Set m_RST = cbo.RowSource
        m_FIELD_NAME = cbo.ListField
    ElseIf TypeName(cbo) = "ComboBox" Then
        Set m_cbo = cbo
    Else
        Err.Raise Number:=vbObjectError + 1, _
        Source:="clsDataCombo.Client", _
        Description:="The object must be a DataCombo or a ComboBox."
    End If
End Property

Private Sub Highlight(ByVal lngLength As Long)
    With m_DCombo
        If Len(.Text) > lngLength Then
            .SelStart = lngLength
            .SelLength = Len(.Text) - lngLength
        End If
    End With
End Sub
'#####################################################################
'#       The following events are inherited from the DataCombo       #
'#####################################################################

Private Sub m_DCombo_GotFocus()
    On Error GoTo m_DCombo_GotFocus_Err
    Set m_FIELD = m_RST.Fields(m_FIELD_NAME)
    m_STOP = True
m_DCombo_GotFocus_Exit:
    Exit Sub
m_DCombo_GotFocus_Err:
    Err.Raise Number:=vbObjectError + 1, _
    Source:="clsDataCombo.m_DCombo_GotFocus", _
    Description:="There was an error in m_DCombo_GotFocus. Likely cause is that the Class has not correctly been initialised."
    Resume m_DCombo_GotFocus_Exit
End Sub

Private Sub m_DCombo_KeyPress(KeyAscii As Integer)
    On Error GoTo m_DCombo_KeyPress_Err
    Static Search As String
    Dim Index As Long
    Dim DoSearch As Boolean
    If m_STOP Then Search = ""
    m_STOP = False
    Select Case KeyAscii
        Case vbKeyBack
            If Len(Search) Then
                Search = Left(Search, Len(Search) - 1)
                DoSearch = True
            End If
            KeyAscii = 0
        Case vbKeyReturn, vbKeyTab
            With m_DCombo
                .SelStart = Len(.Text)
                .SelLength = 0
            End With
            DoSearch = False
            KeyAscii = vbKeyReturn
        Case Is >= vbKeySpace
            Search = Search & UCase(Chr(KeyAscii))
            DoSearch = True
    End Select
    If (DoSearch = True) And Len(Search) > 0 Then
        Dim strCriteria As String, BookMk
        strCriteria = m_FIELD.Name & " like '" & Search & "*'"
        With m_RST
            BookMk = .Bookmark
            .Find strCriteria, , , 1
            If .EOF Then
                'There is no match
                .Bookmark = BookMk
                Search = Left(Search, Len(Search) - 1)
                If Len(Search) = 0 Then m_DCombo.Text = ""
            Else
                m_DCombo.Text = m_FIELD
            End If
        End With
        Call Highlight(Len(Search))
        KeyAscii = 0
    Else
        If Len(Search) Then
            KeyAscii = KeyAscii
        Else
            m_DCombo.Text = ""
        End If
    End If
m_DCombo_KeyPress_Exit:
    Exit Sub
m_DCombo_KeyPress_Err:
    MsgBox prompt:="Unexpected Error - " & Err.Number & " " & Err.Description, _
    Buttons:=vbOKOnly + vbCritical, _
    Title:="clsDataComboSearch.m_DCombo_KeyPress"
End Sub

'#####################################################################
'#        The following events are inherited from the Combo          #
'#####################################################################
Private Sub m_cbo_Change()
Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String

    'don't do if no text or if change was made by autocomplete coding
    If Not blnAuto And m_cbo.Text <> "" Then
        'save the selection start point (cursor position)
        iStart = m_cbo.SelStart
        'get the part the user has typed (not selected)
        strPart = Left$(m_cbo.Text, iStart)
        
        For iLoop = 0 To m_cbo.ListCount - 1
            'compare each item to the part the user has typed,
            '"complete" with the first good match
            strItem = UCase$(m_cbo.List(iLoop))
            If strItem Like UCase$(strPart & "*") And _
                    strItem <> UCase$(m_cbo.Text) Then
                'partial match but not the whole thing.
                '(if whole thing, nothing to complete!)
                blnAuto = True
                m_cbo.SelText = Mid$(m_cbo.List(iLoop), iStart + 1) 'add on the new ending
                m_cbo.SelStart = iStart   'reset the selection
                
                m_cbo.SelLength = Len(m_cbo.Text) - iStart
                blnAuto = False
    
                Exit For
            End If
        Next iLoop
    End If
End Sub

Private Sub m_cbo_KeyDown(KeyCode As Integer, Shift As Integer)
    'Unless we watch out for it, backspace or delete will just delete
    'the selected text (the autocomplete part), so we delete it here
    'first so it doesn't interfere with what the user expects
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        blnAuto = True
        m_cbo.SelText = ""
        blnAuto = False
    ElseIf KeyCode = vbKeyReturn Then 'Accept autocomplete on 'Enter' keypress
        m_cbo_LostFocus
        'the following causes the item to be selected and
        'the cursor placed at the end:
        m_cbo.SelStart = Len(m_cbo.Text)
        
        'This would select the whole thing instead:
        m_cbo.SelLength = Len(m_cbo.Text)
    End If
End Sub

Private Sub m_cbo_LostFocus()
Dim iLoop As Integer
'Match capitalization if item entered is one on the list
    If m_cbo.Text <> "" Then
        For iLoop = 0 To m_cbo.ListCount - 1
            If UCase$(m_cbo.List(iLoop)) = UCase$(m_cbo.Text) Then
                blnAuto = True
                m_cbo.Text = m_cbo.List(iLoop)
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
End Sub
