Imports System.Threading
Imports System.Diagnostics
Imports System.ComponentModel
Imports System.Collections.Specialized
Imports System.Security.Cryptography

Public Delegate Sub ProgressChangedEventHandler(ByVal e As ProgressChangedEventArgs)
Public Delegate Sub CompletedEventHandler(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)

Namespace Spotlib

    Public Class Headers

        Private WorkStop As Integer
        Private WorkTotal As Integer
        Private WorkCounter As Integer

        Private WorkCnt As Long = 0
        Private WorkXMLS() As List(Of Spot)

        Private SUbTotal As Long
        Private LastSpeed As String
        Private ProgressValue As Integer = 1
        Private bNew As Boolean = True

        Private StopFlag As Boolean = False
        Private userStateToLifetime As New HybridDictionary

        Public Event Completed As CompletedEventHandler
        Public Event ProgressChanged As ProgressChangedEventHandler

        Private onCompletedDelegate As SendOrPostCallback
        Private onProgressReportDelegate As SendOrPostCallback

        Private Delegate Sub WorkerEventHandler(ByVal tPhuse As Phuse.Engine, ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation, ByVal TheTask As Guid)

        Friend Sub FindHeaders(ByVal tPhuse As Phuse.Engine, ByVal xParam As NNTPSettings)

            Dim TaskID As Guid = Guid.NewGuid()
            Dim asyncOp As AsyncOperation = AsyncOperationManager.CreateOperation(TaskID)

            SyncLock userStateToLifetime.SyncRoot
                If userStateToLifetime.Count <> 0 Then
                    Throw New ArgumentException("Task already running!", "taskId")
                End If
                userStateToLifetime(TaskID) = asyncOp
            End SyncLock

            Dim workerDelegate As New WorkerEventHandler(AddressOf iFindHeaders)
            workerDelegate.BeginInvoke(tPhuse, xParam, asyncOp, TaskID, Nothing, Nothing)

        End Sub

        Friend Sub New()

            MyBase.New()

            onProgressReportDelegate = New SendOrPostCallback(AddressOf ReportProgress)
            onCompletedDelegate = New SendOrPostCallback(AddressOf WorkCompleted)

        End Sub

        Private Sub WorkCompleted(ByVal operationState As Object)

            Dim e As AsyncCompletedEventArgs = CType(operationState, AsyncCompletedEventArgs)
            OnWorkCompleted(e)

        End Sub

        Private Sub ReportProgress(ByVal state As Object)

            Dim e As ProgressChangedEventArgs = CType(state, ProgressChangedEventArgs)
            OnProgressChanged(e)

        End Sub

        Protected Sub OnWorkCompleted(ByVal e As AsyncCompletedEventArgs)

            RaiseEvent Completed(Me, e)

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

            Dim e As ProgressChangedEventArgs = Nothing
            e = New SpotnetProgressChangedEventArgs(lProg, sInfo, asyncOp.UserSuppliedState)
            asyncOp.Post(Me.onProgressReportDelegate, e)

        End Sub

        Private Function MustCancel(ByVal asyncOp As AsyncOperation) As Boolean

            Return Me.TaskCanceled(asyncOp.UserSuppliedState)

        End Function

        Private Sub iFindHeaders(ByVal hPhuse As Phuse.Engine, ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation, ByVal TheTask As Guid)

            Dim lLast As Long
            Dim lFirst As Long

            Dim UK As Integer
            Dim WK As Worker
            Dim xWork As NNTPWork
            Dim HeaderX As String
            Dim zResp As String = ""
            Dim sError As String = ""
            Dim WorkLoad As List(Of NNTPWork)

            SetP(0, "Verbinding maken...", asyncOp)

            Dim NN As New cNNTP(hPhuse)

            Try

                If Not NN.SelectGroup(xParam.GroupName, lFirst, lLast, 0, 0, sError) Then
                    iCompleted(True, xParam, sError, asyncOp)
                    Exit Sub
                End If

                bNew = (xParam.Position < 1)
                SetP(0, "Zoeken naar " & Utils.sIIF(bNew, "nieuwe ", "") & "spots...", asyncOp)

                Dim TotalSize As Long = 0
                Dim SpeedWatch As New Stopwatch

                If ((lLast - lFirst) < 0) Then
                    iCompleted(False, xParam, vbNullString, asyncOp)
                    Exit Sub
                End If

                If (xParam.Position > 0) Then

                    If lLast <= xParam.Position Then

                        iCompleted(False, xParam, vbNullString, asyncOp)
                        Exit Sub

                    Else

                        If (lFirst <= xParam.Position) Then
                            lFirst = xParam.Position + 1 ' XOVER Fix
                        End If

                    End If

                End If

                WorkCnt = 0
                SUbTotal = 0

                WorkLoad = Utils.CreateWork(lFirst, lLast)

                WorkTotal = WorkLoad.Count
                WorkCounter = WorkTotal

                If WorkTotal < 1 Then
                    iCompleted(False, xParam, vbNullString, asyncOp)
                    Exit Sub
                End If

                ReDim WorkXMLS(WorkTotal - 1)
                Dim RSA() As RSACryptoServiceProvider = Utils.GetRSA(xParam.TrustedKeys)

                For Each xWork In WorkLoad

                    UK += 1
                    HeaderX = Nothing

                    If StopFlag Then
                        Exit Sub
                    End If

                    If MustCancel(asyncOp) Then
                        iCompleted(True, xParam, Utils.CancelMSG, asyncOp)
                        Exit Sub
                    End If

                    If WorkStop = 0 Then

                        sError = vbNullString
                        SpeedWatch.Start()

                        HeaderX = NN.GetHeaders(xParam.GroupName, xWork.xStart, xWork.xEnd, 0, sError)

                        If HeaderX Is Nothing Then
                            If Len(sError) = 0 Then sError = "Er is een time-out opgetreden."
                            iCompleted(True, xParam, "Fout tijdens het ophalen van de headers:" & vbCrLf & vbCrLf & sError, asyncOp)
                            Exit Sub
                        End If

                        If Len(HeaderX) = 0 Then
                            iCompleted(True, xParam, "Fout tijdens het ophalen van de headers:" & vbCrLf & vbCrLf & sError, asyncOp)
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

                    End If

                    If (Not HeaderX Is Nothing) And (WorkStop = 0) Then

                        WK = New Worker

                        WK.RSA = RSA
                        WK.Operation = asyncOp
                        WK.InstanceCount = UK
                        WK.xSettings = xParam
                        WK.HeaderData = Split(HeaderX, vbCrLf)

                        If UBound(WK.HeaderData) < 2 Then
                            iCompleted(True, xParam, "Code 720", asyncOp)
                            Exit Sub
                        End If

                        If Len(WK.HeaderData(UBound(WK.HeaderData))) > 0 Then
                            iCompleted(True, xParam, "Code 721", asyncOp)
                            Exit Sub
                        End If

                        If WK.HeaderData(UBound(WK.HeaderData) - 1) <> "." Then
                            iCompleted(True, xParam, "Code 731", asyncOp)
                            Exit Sub
                        End If

                        WK.ParseHeaders(Me)

                    Else

                        Dim zOut As New List(Of Spot)
                        xWorkDone(False, 0, UK, "", zOut, 0, 0, True, xParam, asyncOp)

                    End If

                Next

            Catch ex As Exception
                iCompleted(True, xParam, ex.Message, asyncOp)
                Exit Sub
            End Try

        End Sub

        Friend Sub xWorkDone(ByVal Errors As Boolean, ByVal WorkDone As Integer, ByVal InStanceCount As Integer, ByVal sError As String, ByRef zxOut As List(Of Spot), ByVal xTotal As Long, ByVal xCount As Long, ByVal NoProgress As Boolean, ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation)

            If StopFlag Then Exit Sub

            If MustCancel(asyncOp) Then
                iCompleted(True, xParam, Utils.CancelMSG, asyncOp)
                Exit Sub
            End If

            If Errors Then
                iCompleted(True, xParam, "Fout tijdens het verwerken van spots: " & sError, asyncOp)
                Exit Sub
            End If

            If WorkDone > 0 Then WorkStop = WorkDone

            If xCount > 0 Then

                WorkCnt += xCount
                WorkXMLS(InStanceCount - 1) = zxOut

            End If

            SUbTotal += xTotal
            WorkCounter = WorkCounter - 1

            If (Not NoProgress) Then

                ProgressValue = CInt((100 / WorkTotal) * (WorkTotal - WorkCounter))

                If WorkCnt < 1 Then
                    SetP(0, "Zoeken naar " & Utils.sIIF(bNew, "nieuwe ", "") & "spots..." & LastSpeed, asyncOp) '' (totaal " & FormatLong2(SUbTotal) & ").")
                Else
                    SetP(ProgressValue, Utils.FormatLong(WorkCnt) & " " & Utils.sIIF(bNew, "nieuwe ", "") & "spots gevonden" & LastSpeed, asyncOp) '' (totaal " & FormatLong2(SUbTotal) & ").")
                End If

            End If

            If WorkCounter = 0 Then

                iCompleted(False, xParam, vbNullString, asyncOp)

            End If

        End Sub

        Private Sub CompletionMethod(ByVal exc As Exception, ByVal asyncOp As AsyncOperation, ByVal tResults As List(Of Spot), ByVal tDeletes As HashSet(Of String))

            SyncLock userStateToLifetime.SyncRoot
                userStateToLifetime.Clear()
            End SyncLock

            Dim e As New SpotsCompletedEventArgs(tResults, tDeletes, exc, Not exc Is Nothing, asyncOp.UserSuppliedState)
            asyncOp.PostOperationCompleted(onCompletedDelegate, e)

        End Sub

        Friend Sub iCompleted(ByVal bError As Boolean, ByVal xParam As NNTPSettings, ByVal sError As String, ByVal asyncOp As AsyncOperation, Optional ByVal LastMessage As Long = 0)

            If Not StopFlag Then

                StopFlag = True

                If bError Then
                    Me.CompletionMethod(New Exception(sError), asyncOp, Nothing, Nothing)
                    Exit Sub
                End If

                Dim TheReturn As New List(Of Spot)
                Dim TheDeletes As New HashSet(Of String)

                Try

                    SetP(0, "Spots verzamelen...", asyncOp)

                    If Not WorkXMLS Is Nothing Then
                        For Zx = UBound(WorkXMLS) To 0 Step -1
                            If Not WorkXMLS(Zx) Is Nothing Then
                                If WorkXMLS(Zx).Count > 0 Then

                                    If Zx <= WorkStop Or (WorkStop = 0) Then

                                        For Each TheSpot As Spot In WorkXMLS(Zx)

                                            If TheSpot.KeyID <> 2 Then

                                                TheReturn.Add(TheSpot)

                                            Else

                                                If Len(TheSpot.Title) < 3 Then Continue For

                                                Dim sCmd() As String = Split(TheSpot.Title, " ")

                                                If UBound(sCmd) < 1 Then Continue For

                                                Select Case sCmd(0).ToLower

                                                    Case "dispose"

                                                        If UBound(sCmd) < 1 Then Continue For
                                                        If Len(sCmd(1)) < 1 Then Continue For

                                                        TheDeletes.Add(Utils.MakeMsg(sCmd(1), False))

                                                End Select

                                            End If

                                        Next

                                    End If

                                End If
                            End If
                        Next
                    End If

                    Me.CompletionMethod(Nothing, asyncOp, TheReturn, TheDeletes)

                Catch ex As Exception

                    Me.CompletionMethod(New Exception(ex.Message), asyncOp, Nothing, Nothing)

                End Try

            End If

        End Sub

    End Class

End Namespace
