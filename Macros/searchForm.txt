'@Author: Ashok Erva
'@dated: November 21, 2015

Option Compare Database

'@name: Form_Load
'@description: This function truncates the table (search) and
'              refreshes the current form

Private Sub Form_Load()
    TruncateSearch
    refreshForms
End Sub


'@name: searchCmd_Click
'description: This procedure gets the searck key and value
'             from the form(drop-down menu & input box) and
'             checks for the relevant match in the database
'             if found pushes the founded information into
'             the table(search).

Private Sub searchCmd_Click()

    Dim dbs As Database
    Dim rs As Recordset
    Dim strSQL As String
    Set dbs = CurrentDb
    Dim jreference, jprofile, jrequirement As String

    If Not getSearchColumn = "" And Not getSearchKey = "" Then
        TruncateSearch
        strSQL = "select * from applications where " + getSearchColumn + "   =   '" + getSearchKey + "'"
        Set rs = dbs.OpenRecordset(strSQL)
        If Not (rs.EOF And rs.BOF) Then
            Do While Not rs.EOF
                If rs("reference") = "" Or IsNull(rs("reference")) Then
                    jreference = "No reference found"
                Else
                    jreference = rs("reference")
                End If
                If IsNull(rs("jprofile")) Then
                     jprofile = "No Job profile found"
                Else
                    jprofile = rs("jprofile")
                End If
                If IsNull(rs("jrequirement")) Then
                    jrequirement = "No Job requirements found"
                Else
                    jrequirement = rs("jrequirement")
                End If
                On Error GoTo Err_DBError
                   CurrentDb.Execute " INSERT INTO search " _
                     & "(company, jreference, jtitle, jprofile, jrequirements) VALUES " _
                     & "('" + rs("company") + "', '" + jreference + "', '" + rs("jtitle") + "', '" + jprofile + "', '" + jrequirement + "');"
                   rs.MoveNext
            Loop
            refreshForms
        Else
            MsgBox "no Match found!"
            refreshForms
        End If
    End If
Exit_DBError:
        Exit Sub
Err_DBError:
    MsgBox "Database Error!"

End Sub

'@name: getSearchColumn
'@description: This function gets the value from the form's drop down menu
'              and returns the valid application table column name based
'              on the selected drop down.

Public Function getSearchColumn()
    Column = Me.searchBox.Value
    If Not Column = "" Then
        If Column = "Company Name" Then
            getSearchColumn = "company"
        End If
        If Column = "Reference ID" Then
            getSearchColumn = "reference"
        End If
        If Column = "Job Title" Then
            getSearchColumn = "jtitle"
        End If
    Else
        getSearchColumn = ""
    End If
End Function

'@name: getSearchKey
'@description: This function gets the search key from the forms input field

Public Function getSearchKey()
    getSearchKey = Me.searchKey.Value
End Function

'@name: TruncateSearch
'@description: This function truncates/deletes all the information from the table(search)

Public Function TruncateSearch()
    CurrentDb.Execute "DELETE FROM search", dbFailOnError
End Function

'@name: refreshForms
'@description: This function refreshes the current form.

Public Function refreshForms()
    Me.Requery
End Function
