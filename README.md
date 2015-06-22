# VB-Script
Automated Email attachment download from a specific Date and Sender from Outlook to local disk or cloud
Module SR_Html


    Dim isAttachment As Boolean
    Dim mailBox As Object
    Dim olFolder As Object
    Dim destFolder As Object
    Dim olFolder1 As Object
    Dim fsSaveFolder, sSavePathFS, ssender As String
    Dim objNamespace As Object
    'Dim Msg As Object
    Dim sysDate As Date
    Dim colItems As Object
    Dim colFilteredItems As Object
    Dim intMsgCount As Integer

    Dim fileName As Object
    ' Dim tokens As Object
    'Dim counter As Integer



    Dim objMsg1 As Object
    Dim Msg1 As Object
    Dim intSize As Object


    
    Private Property objOutlook As Object

    Sub Main()

        fsSaveFolder = "C:\Users\path\"

        isAttachment = False

        objOutlook = CreateObject("Outlook.Application")
        objNamespace = objOutlook.GetNamespace("MAPI")
        mailBox = objNamespace.Folders("UR Email Address")
        olFolder = mailBox.Folders("Inbox")

        destFolder = olFolder.Folders("")

        colItems = olFolder.Items
        colFilteredItems = colItems.Restrict("[Unread] =  True")

        If olFolder Is Nothing Then Exit Sub

        sysDate = Date.Today()

        For Each msg In colItems
            If (msg.Subject = "" Or msg.Subject = "") And msg.Unread = True And (DatePart("yyyy", msg.ReceivedTime) = DatePart("yyyy", sysDate) And DatePart("m", msg.ReceivedTime) = DatePart("m", sysDate) And DatePart("d", msg.ReceivedTime) = DatePart("d", sysDate)) Then
                intSize = intSize + 1
            End If

        Next
        'MsgBox(intSize)
        For Each Msg In colItems
            If (Msg.Subject = "" Or Msg.Subject = "") And Msg.Unread = True And (DatePart("yyyy", Msg.ReceivedTime) = DatePart("yyyy", sysDate) And DatePart("m", Msg.ReceivedTime) = DatePart("m", sysDate) And DatePart("d", Msg.ReceivedTime) = DatePart("d", sysDate)) Then

                intMsgCount = Msg.Attachments.Count
                If intMsgCount > 0 Then
                    'If Msg.attachments() <> "" And Msg.attachments().filename <> "" Then Continue For
                    'MsgBox("here")
                    For mt As Integer = 1 To intMsgCount
                        'MsgBox("move attachment")
                        sSavePathFS = fsSaveFolder & Msg.Attachments(mt).FileName
                        Msg.Attachments(mt).SaveAsFile(sSavePathFS)
                    Next mt
                    Msg.Unread = False

                End If


            End If

        Next
        For Each msg In colItems
            If (msg.Subject = "" Or msg.Subject = "") And (DatePart("yyyy", msg.ReceivedTime) = DatePart("yyyy", sysDate) And DatePart("m", msg.ReceivedTime) = DatePart("m", sysDate) And DatePart("d", msg.ReceivedTime) = DatePart("d", sysDate)) Then

                msg.move(destFolder)
                ' msg.Unread = True

            End If
        Next
        ' MsgBox("inside")
                '  MsgBox(DatePart("d", Msg.receivedtime))
        ' End If

        ' Next

    End Sub

End Module

