Attribute VB_Name = "SupplierMailMerge"
'==============================================================================
' Supplier Mail Merge
' Sends one personalized email per supplier:
'   Excel (data)  ->  Word Mail Merge (letter body)  ->  Outlook (email + attachment)
'
' Setup: import this module into Supplier_Mailing.xlsm (see README).
' Uses late binding, so no library references are needed.
' Default mode is "Display": emails open as drafts for review, nothing is sent.
'==============================================================================
Option Explicit

' Word constants (late binding)
Private Const wdFormLetters As Long = 0
Private Const wdSendToNewDocument As Long = 0
Private Const wdDoNotSaveChanges As Long = 0
' Outlook constants
Private Const olMailItem As Long = 0

' Column positions on the Suppliers sheet
Private Const COL_ID As Long = 1
Private Const COL_NAME As Long = 2
Private Const COL_CONTACT As Long = 3
Private Const COL_EMAIL As Long = 4
Private Const COL_ATTACH As Long = 7
Private Const COL_SEND As Long = 8
Private Const COL_STATUS As Long = 9
Private Const COL_SENT_AT As Long = 10

Public Sub SendSupplierMails()
    Dim wsData As Worksheet, wsSet As Worksheet
    Set wsData = ThisWorkbook.Worksheets("Suppliers")
    Set wsSet = ThisWorkbook.Worksheets("Settings")

    Dim templatePath As String, attachFolder As String, subjectText As String, mode As String
    templatePath = ResolvePath(CStr(wsSet.Range("TemplatePath").Value))
    attachFolder = ResolvePath(CStr(wsSet.Range("AttachmentFolder").Value))
    subjectText = CStr(wsSet.Range("Subject").Value)
    mode = LCase$(Trim$(CStr(wsSet.Range("Mode").Value)))   ' "display" or "send"

    If Dir(templatePath) = "" Then
        MsgBox "Word template not found:" & vbCrLf & templatePath, vbExclamation
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first. Word reads the data from the saved file.", vbExclamation
        Exit Sub
    End If
    If mode = "send" Then
        If MsgBox("Mode is SEND. Emails go out immediately. Continue?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    ThisWorkbook.Save    ' Word reads the saved file, so pending edits must be on disk

    Dim wdApp As Object, olApp As Object, mainDoc As Object
    On Error GoTo CleanFail
    Set wdApp = CreateObject("Word.Application")
    Set olApp = CreateObject("Outlook.Application")
    wdApp.Visible = False

    ' Connect the Word template to this workbook as mail merge data source
    Set mainDoc = wdApp.Documents.Open(Filename:=templatePath, ReadOnly:=True, AddToRecentFiles:=False)
    mainDoc.MailMerge.MainDocumentType = wdFormLetters
    mainDoc.MailMerge.OpenDataSource Name:=ThisWorkbook.FullName, _
        ConfirmConversions:=False, ReadOnly:=True, LinkToSource:=True, _
        SQLStatement:="SELECT * FROM `Suppliers$`"

    Dim lastRow As Long, r As Long, recordNo As Long, sent As Long, skipped As Long
    lastRow = wsData.Cells(wsData.Rows.Count, COL_ID).End(xlUp).Row

    Application.ScreenUpdating = False
    For r = 2 To lastRow
        recordNo = r - 1    ' data source record = sheet row minus header
        If UCase$(Trim$(CStr(wsData.Cells(r, COL_SEND).Value))) <> "YES" Then
            wsData.Cells(r, COL_STATUS).Value = "Skipped (Send <> Yes)"
            skipped = skipped + 1
        ElseIf Not IsValidEmail(CStr(wsData.Cells(r, COL_EMAIL).Value)) Then
            wsData.Cells(r, COL_STATUS).Value = "Error: invalid email"
            skipped = skipped + 1
        Else
            Dim attachPath As String
            attachPath = ""
            If Len(Trim$(CStr(wsData.Cells(r, COL_ATTACH).Value))) > 0 Then
                attachPath = attachFolder & Application.PathSeparator & Trim$(CStr(wsData.Cells(r, COL_ATTACH).Value))
            End If

            wsData.Cells(r, COL_STATUS).Value = SendOne(wdApp, olApp, mainDoc, recordNo, _
                CStr(wsData.Cells(r, COL_EMAIL).Value), _
                Replace(subjectText, "{SupplierName}", CStr(wsData.Cells(r, COL_NAME).Value)), _
                attachPath, mode)
            wsData.Cells(r, COL_SENT_AT).Value = Now
            If Left$(wsData.Cells(r, COL_STATUS).Value, 5) <> "Error" Then sent = sent + 1
        End If
    Next r

    MsgBox sent & " emails " & IIf(mode = "send", "sent", "created as drafts") & ", " & skipped & " skipped.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not mainDoc Is Nothing Then mainDoc.Close SaveChanges:=wdDoNotSaveChanges
    If Not wdApp Is Nothing Then wdApp.Quit SaveChanges:=wdDoNotSaveChanges
    Set mainDoc = Nothing: Set wdApp = Nothing: Set olApp = Nothing
    Exit Sub

CleanFail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Merges one record into a new Word document and puts it into an Outlook email.
Private Function SendOne(wdApp As Object, olApp As Object, mainDoc As Object, recordNo As Long, _
                         toAddress As String, subjectText As String, attachPath As String, mode As String) As String
    On Error GoTo Fail
    Dim merged As Object, mail As Object, editor As Object

    With mainDoc.MailMerge
        .Destination = wdSendToNewDocument
        .DataSource.FirstRecord = recordNo
        .DataSource.LastRecord = recordNo
        .Execute Pause:=False
    End With
    Set merged = wdApp.ActiveDocument

    Set mail = olApp.CreateItem(olMailItem)
    mail.To = toAddress
    mail.Subject = subjectText
    mail.Display                                   ' needed to get the Word editor of the mail
    Set editor = mail.GetInspector.WordEditor
    editor.Content.FormattedText = merged.Content.FormattedText

    If attachPath <> "" Then
        If Dir(attachPath) <> "" Then
            mail.Attachments.Add attachPath
        Else
            SendOne = "Warning: attachment missing, "
        End If
    End If

    If mode = "send" Then
        mail.Send
        SendOne = SendOne & "Sent"
    Else
        SendOne = SendOne & "Draft created"
    End If
    merged.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function

Fail:
    SendOne = "Error: " & Err.Description
    On Error Resume Next
    If Not merged Is Nothing Then merged.Close SaveChanges:=wdDoNotSaveChanges
End Function

' Relative paths in Settings are resolved against the workbook folder.
Private Function ResolvePath(p As String) As String
    p = Trim$(p)
    If p = "" Then
        ResolvePath = ThisWorkbook.Path
    ElseIf InStr(p, ":") > 0 Or Left$(p, 2) = "\\" Or Left$(p, 1) = "/" Then
        ResolvePath = p
    Else
        ResolvePath = ThisWorkbook.Path & Application.PathSeparator & p
    End If
End Function

Private Function IsValidEmail(s As String) As Boolean
    s = Trim$(s)
    IsValidEmail = (s Like "*?@?*.?*") And InStr(s, " ") = 0
End Function

' Clears the status columns before a new run.
Public Sub ResetStatus()
    With ThisWorkbook.Worksheets("Suppliers")
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, COL_ID).End(xlUp).Row
        If lastRow >= 2 Then .Range(.Cells(2, COL_STATUS), .Cells(lastRow, COL_SENT_AT)).ClearContents
    End With
End Sub
