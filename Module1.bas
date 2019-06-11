Attribute VB_Name = "Module1"
Option Compare Database

Private Sub formatDate()

Dim rs As DAO.Recordset
Dim oldDate As String
Dim newDate As String
Dim month As String
Dim day As String
Dim fullString As String
Set rs = CurrentDb.OpenRecordset("Door Activity Log")

'Check to see if the recordset actually contains rows
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst 'Unnecessary in this case, but still a good habit
    Do Until rs.EOF = True
        Debug.Print (rs![Door Activity Log Date])
        oldDate = rs![Door Activity Log Date]
        month = Left(oldDate, 2)
        day = Mid(oldDate, 4, 2)
        fullString = day & "/" & month & "/2019"
        
        rs.Edit
        rs![Door Activity Log Date] = fullString
        rs.Update

        'Move to the next record. Don't ever forget to do this.
        rs.MoveNext
    Loop
Else
    MsgBox "There are no records in the recordset."
End If

MsgBox "Finished looping through records."

rs.Close 'Close the recordset
Set rs = Nothing 'Clean up

End Sub
