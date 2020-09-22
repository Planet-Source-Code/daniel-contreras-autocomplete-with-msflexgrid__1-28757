Attribute VB_Name = "modAutocomplete"
Public Function AutoComplete(sTextbox As TextBox, sFlexGrid As MSFlexGrid, sDB As Database, sTable As String, sField As String) As Boolean
On Error Resume Next
Dim sCounter As Integer
Dim OldLen As Integer
Dim sTemp As Recordset
'Set AutoComplete function to FALSE
AutoComplete = False
If Not sTextbox.Text = "" And IsDelOrBack = False Then
'Set OldLen as the sTextbox lenght
OldLen = Len(sTextbox.Text)
    Set sTemp = sDB.OpenRecordset("SELECT * FROM " & sTable & " WHERE " & sField & " LIKE '" & sTextbox.Text & "*'", dbOpenDynaset)
If Err = 3075 Then
    'Here we got a bug!!
End If
    If Not sTemp.RecordCount = 0 Then
        If sTemp.EOF = True And sTemp.BOF = True Then
            MsgBox "Not Matching Records", vbInformation, "Error"
        Else
            sTemp.MoveFirst
            sFlexGrid.Clear
            sFlexGrid.FormatString = "Artist                           |          CD Name                                    | Price   | Reference"
            Do While Not sTemp.EOF
                sFlexGrid.AddItem sTemp.Fields(1).Value
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 1) = sTemp.Fields(2).Value
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 2) = sTemp.Fields(5).Value
                    sFlexGrid.TextMatrix(sFlexGrid.Rows - 1, 3) = sTemp.Fields(4).Value
                sTemp.MoveNext
            Loop
        End If
            If sTextbox.SelText = "" Then
                sTextbox.SelStart = OldLen
            Else
                sTextbox.SelStart = InStr(sTextbox.Text, sTextbox.SelText)
            End If
                sTextbox.SelLength = Len(sTextbox.Text)
                AutoComplete = True
    Else
        sFlexGrid.Clear
    End If
End If
End Function
