Attribute VB_Name = "mdlIntelliSense"
Option Explicit
Public conn As New ADODB.Connection
Public rs As New Recordset
Dim bExit As Boolean

Public Sub IntelliSense(sTextBox As TextBox, sTable As String, sField As String, Optional sDBPass As String)
    On Error Resume Next
    Dim lLen As Long, strSQL As String
    If bExit = True Or sTextBox = "" Then Exit Sub
    lLen = Len(sTextBox)
    strSQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data.mdb;Persist Security Info=False"
    If conn.State <> 1 Then conn.Open strSQL
    strSQL = "SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextBox & "%'"
    If rs.State = 1 Then rs.Close
    rs.Open strSQL, conn, adOpenKeyset, adLockReadOnly
    If rs.EOF And rs.BOF Then Exit Sub
    sTextBox.Text = rs(sField)
    If sTextBox.SelText = "" Then
        sTextBox.SelStart = lLen
    Else
        sTextBox.SelStart = InStr(sTextBox.Text, sTextBox.SelText)
    End If
    sTextBox.SelLength = Len(sTextBox.Text)
End Sub
                
Public Sub CheckKey(lChar As Integer, sTextBox As TextBox)
    If lChar = 8 Or lChar = 46 Then 'Backspace or Delete
        bExit = True
    ElseIf lChar = 9 Or lChar = 13 Then 'Tab or Enter
        sTextBox.SelStart = Len(sTextBox)
        sTextBox.SelLength = 0
        bExit = True
    ElseIf lChar = 32 Then
        If Len(sTextBox.SelText) <> 0 Then sTextBox = sTextBox '& " "
        sTextBox.SelStart = Len(sTextBox)
        sTextBox.SelLength = 0
        bExit = True
    Else
        bExit = False
    End If
End Sub
