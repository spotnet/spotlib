Imports System.ComponentModel

Imports Spotlib

Namespace Spotlib
    Public Class Spotview

        'Private LastTime As Date
        'Private LastBody As String = ""
        'Private DatabaseFile As String = ""
        'Private CacheXoverID As Long = -1
        'Private AskUnload As Boolean = False

        'Private Const XGR As String = "Geen reacties"

        'Private WithEvents CommentStarter As BackgroundWorker
        'Private WithEvents ImageStarter As BackgroundWorker
        'Private WithEvents ExternalStarter As BackgroundWorker

        'Private WithEvents CommentLoader As Comments
        'Private WithEvents CommentUpdater As Comments

        'Private CommentIDCache As New HashSet(Of Long)
        'Private UniqueCache As New Dictionary(Of String, String)
        'Private CommentProgressCache As String
        'Private FetchNewComments As Boolean = True

        'Private MenuFrom As String
        'Private MenuQuery As String
        'Private MenuQueryName As String
        'Private MenuModulus As String

        'Private FC As New FetchCache()
        'Private SkipMessages As New HashSet(Of String)

        'Private Sub xDoStart()

        '    If AskUnload Then Exit Sub

        '    Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.DoStart), DispatcherPriority.Background)

        'End Sub

        'Private Sub StartImage()

        '    If AskUnload Then Exit Sub

        '    ImageStarter = New BackgroundWorker
        '    ImageStarter.RunWorkerAsync(GetSpot.ImageID)

        'End Sub

        'Private Sub DoStart()

        '    Dim sErr As String = ""

        '    If AskUnload Then Exit Sub

        '    Try

        '        If (Not My.Settings.ShowComments) Then
        '            CommentsDone("", True)
        '            Exit Sub
        '        End If

        '        If StartUpdate(GetSpot.MessageID, sErr) Then Exit Sub

        '        CommentsDone(sErr)
        '        Exit Sub

        '    Catch ex As Exception

        '        CommentsDone("DoStart: " & ex.Message)
        '        Exit Sub

        '    End Try

        'End Sub

        'Public Function CancelComments() As Boolean

        '    If Not CommentLoader Is Nothing Then
        '        CommentLoader.Cancel()
        '    End If

        '    If Not CommentUpdater Is Nothing Then
        '        CommentUpdater.Cancel()
        '    End If

        '    Return True

        'End Function

        'Public Function Unload() As Boolean

        '    If AskUnload Then Return True

        '    Try

        '        AskUnload = True
        '        CancelComments()

        '        CommentLoader = Nothing
        '        CommentUpdater = Nothing

        '        Return True

        '    Catch ex As Exception

        '        Foutje("HTMLView_Unload: " & ex.Message)
        '        Return False

        '    End Try

        'End Function

        'Private Function StartUpdate(ByVal sMsgID As String, ByRef zErr As String) As Boolean

        '    If AskUnload Then
        '        zErr = "Exiting"
        '        Return False
        '    End If

        '    If (CommentProgress Is Nothing) Then
        '        zErr = "CommentProgress Is Nothing"
        '        Return False
        '    End If

        '    If CommentProgress.InnerHtml <> CommentProgressCache Then
        '        CommentProgress.InnerHtml = CommentProgressCache
        '    End If

        '    CommentsStatus = _document.GetElementById("CommentsStatus")

        '    ProgressChanged("Reacties laden...", -1)

        '    CommentStarter = New BackgroundWorker

        '    CommentStarter.WorkerReportsProgress = False
        '    CommentStarter.WorkerSupportsCancellation = False

        '    CommentStarter.RunWorkerAsync(sMsgID)

        '    Return True

        'End Function

        'Private Sub CommentLoader_Completed(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles CommentLoader.Completed

        '    CommentLoader = Nothing

        '    If AskUnload Then Exit Sub

        '    Dim xReturn As CommentsCompletedEventArgs = CType(e, CommentsCompletedEventArgs)

        '    If xReturn Is Nothing Then
        '        CommentsDone("xReturn Is Nothing")
        '        Exit Sub
        '    End If

        '    If xReturn.Cancelled Or (Not xReturn.Error Is Nothing) Then
        '        If Len(xReturn.Error.Message) = 0 Then
        '            CommentsDone("Message Is Nothing")
        '        Else
        '            CommentsDone(xReturn.Error.Message)
        '        End If
        '        Exit Sub
        '    End If

        '    If Not FetchNewComments Then
        '        CommentsDone("")
        '        Exit Sub
        '    End If

        '    CheckNewComments()

        'End Sub

        'Private Sub CommentsDone(ByVal sError As String, Optional ByVal NotFetched As Boolean = False)

        '    Try

        '        If AskUnload Then Exit Sub

        '        Dim xReload As String = "<p><A onfocus='this.blur()' HREF='spotnet:reload'><IMG id='reload' onfocus='this.blur()' title='Vernieuwen' style='border: 0px; cursor:hand; width: 32px; height:32px;' SRC=" & Chr(34) & SettingsFolder() & "\Images\refresh.png" & Chr(34) & "></A>"

        '        If Not CommentProgress Is Nothing Then
        '            If Len(sError) > 0 Then
        '                CommentProgress.InnerHtml = "<center>" & HtmlEncode(sError) & "<br></center>" & xReload
        '            Else
        '                If CommentIDCache.Count = 0 Then
        '                    If Not NotFetched Then
        '                        CommentProgress.InnerHtml = "<center>" & "Geen reacties gevonden" & "<br></center>" & xReload
        '                    Else
        '                        CommentProgress.InnerHtml = "<center>" & "Reacties niet opgehaald" & "<br></center>" & xReload
        '                    End If
        '                Else
        '                    CommentProgress.InnerHtml = xReload
        '                End If
        '            End If
        '        End If

        '    Catch ex As Exception

        '        Foutje("CommentsDone: " & ex.Message)

        '    End Try

        'End Sub

        'Private Sub CommentLoader_NewComment(ByVal e As SpotnetNewCommentEventArgs) Handles CommentLoader.NewComment

        '    If Not AskUnload Then NewComment(e.cComment, False)

        'End Sub

        'Private Sub CommentLoader_ProgressChanged(ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles CommentLoader.ProgressChanged

        '    If AskUnload Then Exit Sub

        '    Dim zx As SpotnetProgressChangedEventArgs = CType(e, SpotnetProgressChangedEventArgs)
        '    ProgressChanged(zx.ProgressMessage, zx.ProgressPercentage)

        'End Sub

        'Public Sub ProgressChanged(ByVal sMessage As String, ByVal sValue As Long)

        '    Try

        '        If AskUnload Then Exit Sub

        '        Dim sText As String = ""

        '        If CommentsStatus Is Nothing Then Exit Sub

        '        If (Len(sMessage) > 0) Then

        '            sText = sMessage

        '        Else

        '            If sValue > 100 Then sValue = 100

        '            If sValue > 0 Then
        '                sText = "Reacties laden (" & CStr(sValue) & "%)"
        '            Else
        '                sText = "Reacties laden..."
        '            End If

        '        End If

        '        If CommentsStatus.InnerText <> sText Then CommentsStatus.InnerText = sText

        '    Catch ex As Exception

        '    End Try

        'End Sub

        'Private Sub CommentUpdater_Completed(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles CommentUpdater.Completed

        '    Dim sErr As String = ""
        '    Dim xMessageID As String = ""

        '    CommentUpdater = Nothing
        '    If AskUnload Then Exit Sub

        '    Dim xReturn As CommentsCompletedEventArgs = CType(e, CommentsCompletedEventArgs)

        '    If xReturn Is Nothing Then
        '        CommentsDone("xReturn Is Nothing")
        '        Exit Sub
        '    End If

        '    If xReturn.Cancelled Or (Not xReturn.Error Is Nothing) Then
        '        If Len(xReturn.Error.Message) = 0 Then
        '            CommentsDone("Message Is Nothing")
        '        Else
        '            CommentsDone(xReturn.Error.Message)
        '        End If
        '        Exit Sub
        '    End If

        '    If xReturn.Comments Is Nothing Then
        '        CommentsDone("xReturn.Comments Is Nothing")
        '    End If

        '    If xReturn.Comments.Count = 0 Then
        '        CommentsDone("")
        '        Exit Sub
        '    End If

        '    xMessageID = MakeMsg(GetSpot.MessageID, False)

        '    Dim bDidSome As Boolean = False

        '    For Each xComment As Comment In xReturn.Comments

        '        If xComment.MessageID = xMessageID Then

        '            If Not FC.CacheHash.Contains(xComment.Article) Then
        '                bDidSome = True
        '                FC.Cache.Add(xComment.Article)
        '                FC.CacheHash.Add(xComment.Article)
        '            End If

        '        End If

        '    Next

        '    CacheXoverID = xReturn.Comments.Item(xReturn.Comments.Count - 1).Article

        '    If Not bDidSome Then
        '        CommentsDone("")
        '        Exit Sub
        '    End If

        '    ShowComments(False)

        'End Sub

        'Private Sub CommentUpdater_ProgressChanged(ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles CommentUpdater.ProgressChanged

        '    If AskUnload Then Exit Sub

        '    Dim zx As SpotnetProgressChangedEventArgs = CType(e, SpotnetProgressChangedEventArgs)

        '    If zx.ProgressMessage.Contains("gevonden") Then
        '        ProgressChanged("Reacties bijwerken (" & zx.ProgressPercentage & "%)", zx.ProgressPercentage)
        '    Else
        '        ProgressChanged(zx.ProgressMessage, zx.ProgressPercentage)
        '    End If

        'End Sub

        'Private Sub CommentStarter_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles CommentStarter.DoWork

        '    Try

        '        e.Result = Nothing
        '        If AskUnload Then Exit Sub

        '        Dim zErr As String = ""
        '        Dim Db As Database = New Database

        '        Dim TheList As List(Of Long) = Db.CommentStarter_DoWork(e.Argument.ToString(), DatabaseFile, New Parameters(), FC)

        '        If TheList Is Nothing Then Throw New Exception(zErr)

        '        e.Result = TheList

        '    Catch ex As Exception

        '        Foutje("CommentStarter: " & ex.Message)

        '    End Try

        'End Sub

        'Private Sub CommentStarter_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles CommentStarter.RunWorkerCompleted

        '    Try

        '        CommentStarter = Nothing
        '        If AskUnload Then Exit Sub

        '        If Not e.Result Is Nothing Then

        '            ShowComments(True)
        '            Exit Sub

        '        End If

        '        CommentsDone("") ' Is al een messagebox geweest 

        '    Catch ex As Exception

        '        CommentsDone("CommentStarter_Complete: " & ex.Message)

        '    End Try

        'End Sub

        'Private Sub ImageStarter_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles ImageStarter.DoWork

        '    e.Result = Nothing

        '    If AskUnload Then Exit Sub

        '    Try
        '        e.Result = OpenImage(CType(e.Argument, String))
        '    Catch
        '    End Try

        'End Sub

        'Private Sub ImageStarter_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles ImageStarter.RunWorkerCompleted

        '    ImageStarter = Nothing

        '    If AskUnload Then Exit Sub

        '    Try

        '        Dim sImg As String = CType(e.Result, String)

        '        If Not sImg Is Nothing Then
        '            If sImg.Length > 0 Then

        '                If Len(GetSpot.Web) > 0 Then
        '                    SpotImage.Style = "cursor:hand;" & SpotImage.Style
        '                End If

        '                SpotImage.SetAttribute("SRC", sImg)

        '            Else
        '                SpotImage.OuterHtml = ""
        '            End If
        '        Else
        '            SpotImage.OuterHtml = ""
        '        End If

        '    Catch ex As Exception

        '        Foutje("ImageStarter: " & ex.Message)
        '        Exit Sub

        '    End Try

        '    DoStart()

        'End Sub

        'Private Sub ExternalStarter_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles ExternalStarter.DoWork

        '    Dim p As New System.Diagnostics.Process
        '    Dim s As New System.Diagnostics.ProcessStartInfo(CType(e.Argument, String))

        '    Try

        '        s.UseShellExecute = True
        '        s.WindowStyle = ProcessWindowStyle.Normal
        '        p.StartInfo = s
        '        p.Start()

        '    Catch ex As Exception
        '    End Try

        'End Sub

        'Private Sub ExternalStarter_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles ExternalStarter.RunWorkerCompleted

        '    ExternalStarter = Nothing

        '    If AskUnload Then Exit Sub

        '    Try

        '    Catch
        '        EnableDownload()
        '    End Try

        'End Sub

    End Class

End Namespace
