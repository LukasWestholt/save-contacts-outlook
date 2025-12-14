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

    Dim sel As Selection
    Dim mail As Outlook.MailItem
    Dim path As String
    Dim shell As Object

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

    path = basePath & _
           Format(Now, "yyyymmdd_hhnnss_") & _
           Replace(mail.Subject, ":", "_") & ".msg"

    mail.SaveAs path, olMSG

    Set shell = CreateObject("WScript.Shell")
    shell.Run """" & basePath & "run.bat" & """ """ & path & """", 1, False

End Sub


