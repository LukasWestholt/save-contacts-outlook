VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SendSelectedMailToPython()

    Const MAX_SUBJECT_LEN As Long = 120

    Dim sel As Selection
    Dim mail As Outlook.MailItem
    Dim path As String
    Dim shell As Object
    Dim subjectPart As String

    basePath = Environ("LOCALAPPDATA") & "\save-contacts-outlook\"

    If Dir(basePath, vbDirectory) = "" Then
      MkDir basePath
    End If

    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Then
        MsgBox "Please select an email first."
        Exit Sub
    End If

    If TypeName(sel.Item(1)) <> "MailItem" Then
        MsgBox "Selected item is not an email."
        Exit Sub
    End If

    Set mail = sel.Item(1)

    Function AlphaNumOnly(ByVal txt As String) As String
        Dim i As Long
        Dim ch As String
        Dim result As String

        For i = 1 To Len(txt)
            ch = Mid$(txt, i, 1)
            If ch Like "[A-Za-z0-9 ]" Then
                If ch = " " Then ch = "_"
                result = result & ch
            End If
        Next i

        AlphaNumOnly = result
    End Function

    subjectPart = AlphaNumOnly(mail.Subject)

    If Len(subjectPart) > MAX_SUBJECT_LEN Then
        subjectPart = Left(subjectPart, MAX_SUBJECT_LEN)
    End If

    If Len(subjectPart) = 0 Then
        subjectPart = "NoSubject"
    End If

    path = basePath & _
           Format(Now, "yyyymmdd_hhnnss_") & _
           subjectPart & ".msg"

    mail.SaveAs path, olMSG

    Set shell = CreateObject("WScript.Shell")
    shell.Run """" & basePath & "run.bat" & """ """ & path & """", 1, False

End Sub
