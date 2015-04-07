Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Diagnostics
Imports System.Globalization
Imports System.ComponentModel
Imports System.Collections.Specialized

Imports Spotlib.Common
Imports Fusenet

Public Delegate Sub NewCommentEventHandler(ByVal e As SpotnetNewCommentEventArgs)
Public Delegate Sub CommentProgressChangedEventHandler(ByVal e As ProgressChangedEventArgs)
Public Delegate Sub CommentCompletedEventHandler(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)

Namespace Spotlib.Core

    Public Class Comments

        Private userStateToLifetime As New HybridDictionary
        Private onCompletedDelegate As SendOrPostCallback
        Private onProgressReportDelegate As SendOrPostCallback
        Private onNewCommentDelegate As SendOrPostCallback

        Public Event NewComment As NewCommentEventHandler
        Public Event Completed As CommentCompletedEventHandler
        Public Event ProgressChanged As ProgressChangedEventHandler

        Private Delegate Sub eFindComments(ByVal tPhuse As Phuse.Engine, ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation, ByVal TheTask As Guid)
        Private Delegate Sub eGetComments(ByVal tPhuse As Phuse.Engine, ByVal zList As List(Of Long), ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation, ByVal TheTask As Guid)

        Friend Sub New()

            MyBase.New()

            onCompletedDelegate = New SendOrPostCallback(AddressOf WorkCompleted)
            onNewCommentDelegate = New SendOrPostCallback(AddressOf ReportComment)
            onProgressReportDelegate = New SendOrPostCallback(AddressOf ReportProgress)

        End Sub

        Friend Sub GetComments(ByVal tPhuse As Phuse.Engine, ByVal ArticleIDs As List(Of Long), ByVal xParam As NNTPSettings)

            Dim TaskID As Guid = Guid.NewGuid()
            Dim asyncOp As AsyncOperation = AsyncOperationManager.CreateOperation(TaskID)

            SyncLock userStateToLifetime.SyncRoot
                If userStateToLifetime.Count <> 0 Then
                    Throw New ArgumentException("Task already running!", "taskId")
                End If
                userStateToLifetime(TaskID) = asyncOp
            End SyncLock

            Dim workerDelegate As New eGetComments(AddressOf iGetComments)
            workerDelegate.BeginInvoke(tPhuse, ArticleIDs, xParam, asyncOp, TaskID, Nothing, Nothing)

        End Sub

        Friend Sub FindComments(ByVal tPhuse As Phuse.Engine, ByVal xParam As NNTPSettings)

            Dim TaskID As Guid = Guid.NewGuid()
            Dim asyncOp As AsyncOperation = AsyncOperationManager.CreateOperation(TaskID)

            SyncLock userStateToLifetime.SyncRoot
                If userStateToLifetime.Count <> 0 Then
                    Throw New ArgumentException("Task already running!", "taskId")
                End If
                userStateToLifetime(TaskID) = asyncOp
            End SyncLock

            Dim workerDelegate As New eFindComments(AddressOf iFindComments)
            workerDelegate.BeginInvoke(tPhuse, xParam, asyncOp, TaskID, Nothing, Nothing)

        End Sub

        Private Sub ReportProgress(ByVal state As Object)

            Dim e As ProgressChangedEventArgs = CType(state, ProgressChangedEventArgs)

            OnProgressChanged(e)

        End Sub

        Private Sub ReportComment(ByVal state As Object)

            Dim e As SpotnetNewCommentEventArgs = CType(state, SpotnetNewCommentEventArgs)

            OnNewComment(e)

        End Sub

        Private Sub WorkCompleted(ByVal operationState As Object)

            Dim e As AsyncCompletedEventArgs = CType(operationState, AsyncCompletedEventArgs)

            OnWorkCompleted(e)

        End Sub

        Protected Sub OnWorkCompleted(ByVal e As AsyncCompletedEventArgs)

            RaiseEvent Completed(Me, e)

        End Sub

        Protected Sub OnNewComment(ByVal e As SpotnetNewCommentEventArgs)

            RaiseEvent NewComment(e)

        End Sub

        Protected Sub OnProgressChanged(ByVal e As ProgressChangedEventArgs)

            RaiseEvent ProgressChanged(e)

        End Sub

        Private Function TaskCanceled(ByVal taskId As Object) As Boolean

            Return (userStateToLifetime(taskId) Is Nothing)

        End Function

        Public Sub Cancel()

            SyncLock userStateToLifetime.SyncRoot

                userStateToLifetime.Clear()

            End SyncLock

        End Sub

        Private Sub SetP(ByVal lProg As Integer, ByVal sInfo As String, ByVal asyncOp As AsyncOperation)

            Dim e As ProgressChangedEventArgs = New SpotnetProgressChangedEventArgs(lProg, sInfo, asyncOp.UserSuppliedState)

            asyncOp.Post(Me.onProgressReportDelegate, e)

        End Sub

        Private Sub SendComment(ByRef zCom As Comment, ByVal asyncOp As AsyncOperation)

            Dim e As SpotnetNewCommentEventArgs = New SpotnetNewCommentEventArgs(zCom, asyncOp.UserSuppliedState)

            asyncOp.Post(Me.onNewCommentDelegate, e)

        End Sub

        Private Function MustCancel(ByVal asyncOp As AsyncOperation) As Boolean

            Return Me.TaskCanceled(asyncOp.UserSuppliedState)

        End Function

        Private Sub iFindComments(ByVal tPhuse As Phuse.Engine, ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation, ByVal TheTask As Guid)

            Dim lLast As Long
            Dim lFirst As Long
            Dim zRes As New List(Of Comment)

            Dim UK As Long
            Dim xWork As NNTPWork
            Dim WorkLoad As List(Of NNTPWork)
            Dim HeaderX As String
            Dim zResp As String = ""
            Dim sError As String = ""
            Dim LastP As Integer
            Dim WorkTotal As Integer
            Dim LastSpeed As String = ""
            Dim ProgressValue As Integer = 1

            Dim NN As New cNNTP(tPhuse)

            Try

                SetP(0, "Verbinding maken...", asyncOp)

                If Not NN.SelectGroup(xParam.GroupName, lFirst, lLast, 0, 0, sError) Then
                    iCompleted(zRes, True, sError, asyncOp)
                    Exit Sub
                End If

                Dim bNew As Boolean = (xParam.Position < 1)
                Dim xProg As String = sIIF(bNew, "Zoeken naar reacties", "Zoeken naar nieuwe reacties")

                SetP(0, xProg & "...", asyncOp)

                Dim TotalSize As Long = 0
                Dim SpeedWatch As New Stopwatch()

                If ((lLast - lFirst) < 0) Then
                    iCompleted(zRes, False, vbNullString, asyncOp)
                    Exit Sub
                End If

                If (xParam.Position > 0) Then

                    If lLast <= xParam.Position Then

                        iCompleted(zRes, False, vbNullString, asyncOp)
                        Exit Sub

                    Else

                        If (lFirst <= xParam.Position) Then
                            lFirst = xParam.Position + 1 ' XOVER Fix
                        End If

                    End If

                End If

                WorkLoad = CreateWork(lFirst, lLast)
                WorkTotal = WorkLoad.Count

                If WorkTotal < 1 Then
                    iCompleted(zRes, False, vbNullString, asyncOp)
                    Exit Sub
                End If

                Dim WorkXMLS(WorkTotal - 1) As List(Of Comment)

                For Each xWork In WorkLoad

                    UK += 1
                    HeaderX = Nothing

                    If MustCancel(asyncOp) Then
                        iCompleted(zRes, True, CancelMSG, asyncOp)
                        Exit Sub
                    End If

                    sError = vbNullString

                    SpeedWatch.Start()

                    HeaderX = NN.GetField(xParam.GroupName, "References", xWork.xStart, xWork.xEnd, 0, sError)

                    If HeaderX Is Nothing Then
                        If Len(sError) = 0 Then sError = "Er is een time-out opgetreden."
                        iCompleted(zRes, True, "Fout tijdens het ophalen van de headers:" & vbCrLf & vbCrLf & sError, asyncOp)
                        Exit Sub
                    End If

                    If Len(HeaderX) = 0 Then
                        iCompleted(zRes, True, "Fout tijdens het ophalen van de headers:" & vbCrLf & vbCrLf & sError, asyncOp)
                        Exit Sub
                    End If

                    TotalSize += Len(HeaderX)
                    SpeedWatch.Stop()

                    Dim DatDiff As Single = CSng((TotalSize / 1000) / (SpeedWatch.ElapsedMilliseconds / 1000))

                    If DatDiff < 1000 Then
                        If DatDiff >= 50 Then
                            LastSpeed = "  (" & Math.Round(DatDiff) & " KB/s)"
                        Else
                            LastSpeed = ""
                        End If
                    Else
                        LastSpeed = "  (" & Replace$(CStr(Math.Round((DatDiff / 1000), 1)), ".", ",") & " MB/s)"
                    End If

                    If SpeedWatch.ElapsedMilliseconds > 1000 Then
                        TotalSize = 0
                        SpeedWatch.Reset()
                    End If

                    ProgressValue = CInt(((100 / WorkTotal) * UK) + 1)

                    If ProgressValue <> LastP Then
                        If zRes.Count = 0 Then
                            SetP(0, sIIF(bNew, "Zoeken naar ", "Zoeken naar nieuwe ") & "reacties..." & LastSpeed, asyncOp)
                        Else
                            SetP(ProgressValue, FormatLong(zRes.Count) & " " & sIIF(Not bNew, "nieuwe ", "") & "reacties gevonden" & LastSpeed, asyncOp)
                        End If
                        LastP = ProgressValue
                    End If

                    If MustCancel(asyncOp) Then
                        iCompleted(zRes, True, CancelMSG, asyncOp)
                        Exit Sub
                    End If

                    If (Not HeaderX Is Nothing) And Len(HeaderX) > 0 Then

                        Dim HeaderData() As String = Split(HeaderX, vbCrLf)

                        If UBound(HeaderData) < 2 Then
                            iCompleted(zRes, True, "Fout tijdens het ophalen van de headers:" & vbCrLf & vbCrLf & "Code 620", asyncOp)
                            Exit Sub
                        End If

                        If Len(HeaderData(UBound(HeaderData))) > 0 Then
                            iCompleted(zRes, True, "Fout tijdens het ophalen van de headers:" & vbCrLf & vbCrLf & "Code 621", asyncOp)
                            Exit Sub
                        End If

                        If HeaderData(UBound(HeaderData) - 1) <> "." Then
                            iCompleted(zRes, True, "Fout tijdens het ophalen van de headers:" & vbCrLf & vbCrLf & "Code 631", asyncOp)
                            Exit Sub
                        End If

                        If (UBound(HeaderData) > 2) Then

                            For xLoop As Integer = UBound(HeaderData) - 2 To 1 Step -1 '' Van nieuw naar oud  (ivm CHeckMessageID)

                                Dim LineData() As String = Split(HeaderData(xLoop), " ")
                                If UBound(LineData) < 1 Then Continue For

                                Dim ZK As New Comment

                                ZK.Article = CLng(LineData(0))
                                If ZK.Article < 1 Then Continue For

                                If LineData(1) Is Nothing Then Continue For
                                If LineData(1).Length < 4 Then Continue For

                                ZK.MessageID = LineData(1).Substring(1, LineData(1).Length - 2)

                                zRes.Add(ZK)

                            Next

                        End If

                    End If

                Next

                If MustCancel(asyncOp) Then
                    iCompleted(zRes, True, CancelMSG, asyncOp)
                    Exit Sub
                End If

                zRes.Reverse()

                SetP(100, FormatLong(zRes.Count) & " " & sIIF(Not bNew, "nieuwe ", "") & "reacties gevonden" & LastSpeed, asyncOp)

                iCompleted(zRes, False, vbNullString, asyncOp)

            Catch ex As Exception

                iCompleted(zRes, True, ex.Message, asyncOp)
                Exit Sub

            End Try

        End Sub

        Private Sub iGetComments(ByVal tPhuse As Phuse.Engine, ByVal ArticleIDs As List(Of Long), ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation, ByVal TheTask As Guid)

            Dim sError As String = ""
            Dim zRes As New List(Of Comment)

            Dim xProg As String = "Reacties ophalen"
            SetP(0, xProg & "...", asyncOp)

            Dim TC As Comment
            Dim zRet As String = ""
            Dim LastP As Integer = -1
            Dim LoopCount As Long = 0
            Dim ProgressValue As Integer = 0
            Dim NN As New cNNTP(tPhuse)

            Try

                For Each xArticle As Long In ArticleIDs

                    LoopCount += 1

                    ProgressValue = CInt((100 / ArticleIDs.Count) * LoopCount)

                    If ProgressValue <> LastP Then
                        SetP(ProgressValue, xProg & " (" & ProgressValue & "%)", asyncOp)
                        LastP = ProgressValue
                    End If

                    If MustCancel(asyncOp) Then
                        iCompleted(zRes, True, CancelMSG, asyncOp)
                        Exit Sub
                    End If

                    zRet = vbNullString
                    sError = vbNullString

                    Dim xRet As Integer = -1

                    If Not NN.GetArticle(xParam.GroupName, CStr(xArticle), zRet, xRet, sError) Then

                        If (xRet = 423) Then Continue For

                        iCompleted(zRes, True, sError, asyncOp)
                        Exit Sub

                    End If

                    If zRet.Substring(zRet.Length - 3) <> ("." & vbCrLf) Then

                        iCompleted(zRes, True, "Invalid ending", asyncOp)
                        Exit Sub

                    End If

                    Dim sErr As String = ""
                    TC = Article2Comment(xArticle, zRet, xParam, sErr)

                    If TC Is Nothing Then Continue For

                    zRes.Add(TC)
                    SendComment(TC, asyncOp)

                Next

                If MustCancel(asyncOp) Then
                    iCompleted(zRes, True, CancelMSG, asyncOp)
                    Exit Sub
                End If

                SetP(100, xProg & "...", asyncOp)

                iCompleted(zRes, False, "", asyncOp)
                Exit Sub

            Catch ex As Exception

                iCompleted(zRes, True, ex.Message, asyncOp)
                Exit Sub

            End Try

        End Sub

        Private Function Article2Comment(ByVal xArticleID As Long, ByVal sArt As String, ByVal xParam As NNTPSettings, ByRef sError As String) As Comment

            Try

                Dim XL() As String
                Dim KL As New Comment
                Dim zAvatar As String = ""

                KL.Article = xArticleID

                KL.User = New UserInfo

                Dim Z() As String = Split(sArt, vbCrLf & vbCrLf)

                Dim sHead As String = Z(0)
                Dim sBody As String = sArt.Substring(Z(0).Length + 4)

                XL = Split(sHead, vbCrLf)

                For iXL = 1 To UBound(XL) - 1

                    If UCase(XL(iXL)).StartsWith("FROM: ") Then
                        KL.From = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                        For RR = iXL + 1 To UBound(XL) - 1
                            If XL(RR).IndexOf(":") = -1 Or XL(RR).StartsWith(" ") Or XL(RR).StartsWith(Chr(9)) Then
                                KL.From += XL(RR) ' Multine from
                            Else
                                Exit For
                            End If
                        Next
                    End If

                    If UCase(XL(iXL)).StartsWith("DATE: ") Then
                        KL.Created = ConvertDate(Mid(XL(iXL), XL(iXL).IndexOf(":") + 3))
                    End If

                    If UCase(XL(iXL)).StartsWith("MESSAGE-ID: ") Then
                        KL.MessageID = MakeMsg(Mid(XL(iXL), XL(iXL).IndexOf(":") + 3))
                    End If

                    If UCase(XL(iXL)).StartsWith("X-USER-AVATAR: ") Then
                        zAvatar += Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                    End If

                    If UCase(XL(iXL)).StartsWith("X-USER-KEY: ") Then

                        With KL.User

                            .Modulus = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)

                            If .Modulus.ToLower.Contains("<modulus>") Then

                                .Modulus = .Modulus.Substring(.Modulus.ToLower.IndexOf("<modulus>") + 9)
                                If .Modulus.Contains("<") Then .Modulus = .Modulus.Substring(0, .Modulus.IndexOf("<"))

                            Else

                                .Modulus = FixPadding(UnSpecialString(.Modulus))

                            End If

                        End With

                    End If

                    If UCase(XL(iXL)).StartsWith("X-USER-SIGNATURE: ") Then
                        KL.User.Signature = UnSpecialString(Mid(XL(iXL), XL(iXL).IndexOf(":") + 3))
                    End If

                    If UCase(XL(iXL)).StartsWith("ORGANIZATION: ") Then
                        KL.User.Organisation = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                        KL.User.Organisation = KL.User.Organisation.Substring(0, 1).ToUpper & KL.User.Organisation.Substring(1) ' Capitalize first char
                    End If

                    If UCase(XL(iXL)).StartsWith("X-TRACE: ") Then
                        KL.User.Trace += vbCrLf & Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                    End If

                    If UCase(XL(iXL)).StartsWith("NNTP-POSTING-HOST: ") Then
                        Dim sHost As String = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                        If sHost.IndexOf(" (") > 0 Then
                            sHost = sHost.Replace(")", "")
                            sHost = sHost.Substring(0, sHost.IndexOf(" ("))
                        End If
                        KL.User.Trace += vbCrLf & sHost
                    End If

                    If UCase(XL(iXL)).StartsWith("X-ORIGINATING-IP: ") Then
                        KL.User.Trace += vbCrLf & Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                    End If

                Next

                If xParam.BlackList.Contains(KL.User.Modulus) Then
                    sError = "Blacklist"
                    Return Nothing
                End If

                KL.User.Avatar = FixPadding(UnSpecialString(zAvatar))
                KL.User.Trace = KL.User.Trace.Replace(vbCrLf & KL.User.Organisation, "").Trim

                If KL.User.Trace = vbCrLf Then KL.User.Trace = ""

                KL.From = Trim(Split(KL.From, "<")(0))
                KL.Body = sBody.Substring(0, sBody.Length - 5)
                KL.Body = KL.Body.Replace(vbCrLf & "..", ".")

                If Len(KL.From) = 0 Then sError = "Sip1" : Return Nothing
                If Len(KL.Body) = 0 Then sError = "Sip3" : Return Nothing
                If Len(KL.MessageID) = 0 Then sError = "Sip2" : Return Nothing

                KL.User.ValidSignature = False

                If Not xParam.CheckSignatures Then Return KL

                KL.User.ValidSignature = CheckUserSignature(KL.MessageID, KL.User.Signature, KL.User.Modulus)

                If Not KL.User.ValidSignature Then
                    KL.User.ValidSignature = CheckUserSignature(KL.MessageID & KL.Body & vbCrLf & KL.From, KL.User.Signature, KL.User.Modulus)
                End If

                If Not KL.User.ValidSignature Then sError = "Invalid signature" : Return Nothing

                Return KL

            Catch ex As Exception

                sError = Err.Description
                Return Nothing

            End Try

        End Function

        Private Sub iCompleted(ByRef Comments As List(Of Comment), ByVal bError As Boolean, ByVal sError As String, ByVal asyncOp As AsyncOperation, Optional ByVal LastMessage As Long = 0)

            If bError Then
                Me.CompletionMethod(New Exception(sError), asyncOp, Nothing)
            Else
                Me.CompletionMethod(Nothing, asyncOp, Comments)
            End If

        End Sub

        Private Sub CompletionMethod(ByVal exc As Exception, ByVal asyncOp As AsyncOperation, ByRef tResults As List(Of Comment))

            SyncLock userStateToLifetime.SyncRoot
                userStateToLifetime.Clear()
            End SyncLock

            Dim e As New CommentsCompletedEventArgs(tResults, exc, Not exc Is Nothing, asyncOp.UserSuppliedState)

            asyncOp.PostOperationCompleted(onCompletedDelegate, e)

        End Sub

    End Class

End Namespace
