Imports System.IO
Imports System.Text
Imports System.Net.Sockets
Imports System.Net.Security
Imports System.Diagnostics.Debug
Imports System.Security.Cryptography.X509Certificates

Imports Fusenet
Imports Spotlib.Common

Namespace Spotlib

    Public Class cNNTP

        Private tPhuse As Phuse.Engine = Nothing

        Friend Sub New(ByVal hPhuse As Phuse.Engine)

            tPhuse = hPhuse

        End Sub

        Private Function GetResponse(ByVal sGroup As String, ByVal sCommands As List(Of String), ByRef sReturn As String, ByRef sError As String) As Integer

            Try

                sError = ""
                sReturn = ""

                sReturn = tPhuse.Slots.Send(sGroup, sCommands)

                Return GetCode(sReturn)

            Catch ex As Exception

                sError = ex.Message
                Return -1

            End Try

        End Function

        Private Function TranslateError(ByVal lRet As Integer, ByVal OriginalError As String) As String

            Select Case lRet

                Case 381, 450, 452, 480, 481, 482

                    Return "Gebruikersnaam en/of wachtwoord is onjuist."

                Case 400

                    Return "Maximaal aantal verbindingen bereikt."

                Case 411

                    Return "Groep niet gevonden."

                Case 931
                    Return "Er is een time-out opgetreden."

                Case 941
                    Return "Hostnaam niet gevonden."

            End Select

            Return OriginalError

        End Function

        Public Function SelectGroup(ByVal sGroup As String, ByRef lFirst As Long, ByRef lLast As Long, ByRef lCount As Long, ByRef lRet As Integer, ByRef sError As String) As Boolean

            Dim sRet As String = ""
            Dim bContinue As Boolean = False

            lRet = GetResponse(sGroup, New List(Of String), sRet, sError)

            Select Case lRet

                Case 211

                    bContinue = True

                Case -1

                    lRet = GetCode(sError)

                Case Else

                    sError = Left(sRet, 500)
                    lRet = GetCode(sError)

            End Select

            If Not bContinue Then

                sError = TranslateError(lRet, sError)
                Return False

            End If

            Dim sArt() As String

            Try

                sArt = sRet.Split(" "c)
                lLast = CLng(sArt(3))
                lFirst = CLng(sArt(2))
                lCount = CLng(sArt(1))

                Return True

            Catch ex As Exception

                lRet = 560
                sError = ex.Message
                Return False

            End Try

        End Function

        Public Function GetHeaders(ByVal sGroup As String, ByVal xStart As Long, ByVal xEnd As Long, ByRef lRet As Integer, ByRef sError As String) As String

            Dim sRet As String = ""

            lRet = GetResponse(sGroup, StringToList("XOVER " & xStart & "-" & xEnd), sRet, sError)

            Select Case lRet

                Case 224

                    Return sRet

                Case -1

                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return Nothing

                Case Else

                    sError = Left(sRet, 500)
                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return Nothing

            End Select

        End Function

        Public Function GetField(ByVal sGroup As String, ByVal sField As String, ByVal xStart As Long, ByVal xEnd As Long, ByRef lRet As Integer, ByRef sError As String) As String

            Dim sRet As String = ""

            lRet = GetResponse(sGroup, StringToList("XHDR " & sField & " " & xStart & "-" & xEnd), sRet, sError)

            Select Case lRet

                Case 221

                    Return sRet

                Case -1

                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return Nothing

                Case Else

                    sError = Left(sRet, 500)
                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return Nothing

            End Select

        End Function

        Public Function PostData(ByVal sGroup As String, ByVal sData As String, ByRef sResp As String, ByRef lRet As Integer, ByRef sError As String) As Boolean

            Dim zOut As New List(Of String)

            zOut.Add("POST")
            zOut.Add(sData)

            lRet = GetResponse(sGroup, zOut, sResp, sError)

            Select Case lRet

                Case 240

                    Return True

                Case -1

                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

                Case Else

                    sError = Left(sResp, 500)
                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

            End Select

        End Function

        Public Function GetHeader(ByVal sGroup As String, ByVal sArticleID As String, ByRef sResp As String, ByRef lRet As Integer, ByRef sError As String) As Boolean

            If Len(sArticleID) = 0 Then
                sError = "No article"
                Return False
            End If

            lRet = GetResponse(sGroup, StringToList("HEAD " & sArticleID), sResp, sError)

            Select Case lRet

                Case 221

                    Return True

                Case -1

                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

                Case Else

                    sError = Left(sResp, 500)
                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

            End Select

        End Function

        Public Function GetArticle(ByVal sGroup As String, ByVal sArticleID As String, ByRef sResp As String, ByRef lRet As Integer, ByRef sError As String) As Boolean

            If Len(sArticleID) = 0 Then
                sError = "No article"
                Return False
            End If

            lRet = GetResponse(sGroup, StringToList("ARTICLE " & sArticleID), sResp, sError)

            Select Case lRet

                Case 220

                    Return True

                Case -1

                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

                Case Else

                    sError = Left(sResp, 500)
                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

            End Select

        End Function

        Public Function GetBody(ByVal sGroup As String, ByVal sArticleID As String, ByRef sResp As String, ByRef lRet As Integer, ByRef sError As String) As Boolean

            If Len(sArticleID) = 0 Then
                sError = "No article"
                Return False
            End If

            lRet = GetResponse(sGroup, StringToList("BODY " & sArticleID), sResp, sError)

            Select Case lRet

                Case 222

                    Return True

                Case -1

                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

                Case Else

                    sError = Left(sResp, 500)
                    lRet = GetCode(sError)
                    sError = TranslateError(lRet, sError)

                    Return False

            End Select

        End Function

        Private Function StringToList(ByVal sIn As String) As List(Of String)

            Dim zOut As New List(Of String)

            zOut.Add(sIn)

            Return zOut

        End Function

    End Class

End Namespace
