'@Author: Ashok Erva
'@dated: November 21, 2015

'@name: cmdclear_Click
'@description: clears all the fields in the form.
Private Sub cmdclear_Click()
    Me.CmdCompany = ""
    Me.cmdreference = ""
    Me.cmdtitle = ""
    Me.cmdrequirement = ""
    Me.cmdprofile = ""
End Sub

'@name: smdsave_Click
'@description: gets all the informations from all the
'               fields and pushes to the database(applications)

Private Sub cmdsave_Click()
    Dim Cname, RefNr, Title, Jprofile, Jrequirement As String

    Cname = getCName()
    RefNr = getReference()
    Title = getTitle()
    Jprofile = getProfile()
    Jrequirement = getRequirement()

    If Cname = "" Or Title = "" Or Jprofile = "" Or Jrequirement = "" Then
        MsgBox "Please fill the form"
        Exit Sub
    Else
    On Error GoTo Err_DBError


        CurrentDb.Execute " INSERT INTO applications " _
        & "(company, reference, jtitle, jprofile, jrequirement) VALUES " _
         & "(""" + Cname + """, """ + RefNr + """, """ + Title + """, """ + Jprofile + """, """ + Jrequirement + """);"
        MsgBox "Succesfully Saved!! "
        cmdclear_Click

Exit_DBError:
        Exit Sub
Err_DBError:
    MsgBox "Database Error!"
    End If

End Sub

'@name: getCName
'description: gets the company name from the field company.

Public Function getCName()
    If Len(Me.CmdCompany) > 1 And Not Me.CmdCompany = "" Then
        getCName = replaceQuotes(Me.CmdCompany)
    Else
        MsgBox "Enter a valid company name"
        getCName = ""
    End If
End Function

'@name: getReference
'description: gets the job reference id form the field field reference
'              Number

Public Function getReference()
    If Len(Me.cmdreference) > 1 And Not Me.cmdreference = "" Then
        getReference = replaceQuotes(Me.cmdreference)
    Else
        getReference = ""
    End If
End Function

'@name: getTitle
'@description: gets the job title from the filed job title

Public Function getTitle()
    If Len(Me.cmdtitle) > 1 And Not Me.cmdtitle = "" Then
        getTitle = replaceQuotes(Me.cmdtitle)
    Else
        MsgBox "Enter a valid Position Title"
        getTitle = ""
    End If
End Function

'@name: getProfile
'@description: gets the job profile info form the filed profile

Public Function getProfile()
    If Len(Me.cmdprofile) > 1 And Not Me.cmdprofile = "" Then
        getProfile = replaceQuotes(Me.cmdprofile)
    Else
        MsgBox "Enter some valid job profile"
        getProfile = ""
    End If
End Function

'@name: getRequirement
'@description: gets the job requirements info form the field requirements

Public Function getRequirement()
    If Len(Me.cmdrequirement) > 1 And Not Me.cmdrequirement = "" Then
        getRequirement = replaceQuotes(Me.cmdrequirement)
    Else
        MsgBox "Enter some valid job requirements"
        getRequirement = ""
    End If
End Function

'@name: replaceQuotes
'@param: actualText string (a string with "'s)
'@description: replaces all the double quotes(")
'               With ""

Public Function replaceQuotes(actualText As String)
    DQ = Chr(34)
    replaceQuotes = Replace(actualText, DQ, "")
End Function
