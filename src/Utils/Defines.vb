Imports System.ComponentModel
Imports DataVirtualization

Namespace Spotlib
    Public Class Spot

        Public KeyID As Byte
        Public Stamp As Integer
        Public Category As Byte
        Public SubCat As Byte
        Public SubCats As String
        Public Title As String
        Public Poster As String
        Public Tag As String
        Public Filesize As Long
        Public MessageID As String
        Public Article As Long
        Public Modulus As String

    End Class

    Public Class SpotEx

        Inherits Spot

        Public Body As String = ""
        Public Web As String = ""
        Public Image As String = ""

        Public NZB As String = ""
        Public ImageID As String = ""
        Public ImageWidth As Integer = 0
        Public ImageHeight As Integer = 0

        Public User As New UserInfo
        Public OldInfo As New FTDInfo

    End Class

    Public Class FTDInfo

        Public Groups As String = ""
        Public FileName As String = ""

    End Class

    Public Class UserInfo

        Public Trace As String = ""
        Public Avatar As String = ""
        Public Modulus As String = ""
        Public Signature As String = ""
        Public Organisation As String = ""
        Public ValidSignature As Boolean

    End Class

    Public Class Comment

        Public From As String
        Public Created As Date
        Public Body As String
        Public User As UserInfo
        Public MessageID As String
        Public Article As Long

    End Class

    Friend Class NNTPWork

        Public xStart As Long
        Public xEnd As Long
        Public xDone As Boolean

    End Class

    Public Class NNTPSettings

        Public Position As Long
        Public GroupName As String
        Public TrustedKeys() As String
        Public CheckSignatures As Boolean

        Public BlackList As HashSet(Of String)
        Public WhiteList As HashSet(Of String)

    End Class

    Public Class SpotnetProgressChangedEventArgs
        Inherits ProgressChangedEventArgs

        Private ProgMsg As String = ""

        Public Sub New(
        ByVal progressPercentage As Integer,
        ByVal progressMessage As String,
        ByVal UserState As Object)

            MyBase.New(progressPercentage, UserState)

            Me.ProgMsg = progressMessage

        End Sub

        Public ReadOnly Property ProgressMessage() As String
            Get
                Return ProgMsg
            End Get
        End Property

    End Class

    Public Class SpotnetNewCommentEventArgs
        Inherits ProgressChangedEventArgs

        Private ProgComment As Comment

        Public Sub New(
        ByVal PC As Comment,
        ByVal UserState As Object)

            MyBase.New(0, UserState)

            Me.ProgComment = PC

        End Sub

        Public ReadOnly Property cComment() As Comment
            Get
                Return ProgComment
            End Get
        End Property

    End Class

    Public Class SpotsCompletedEventArgs
        Inherits AsyncCompletedEventArgs

        Public DbFile As String = ""
        Public Spots As List(Of Spot)
        Public Deletes As HashSet(Of String)

        Public Sub New(
    ByRef tSpots As List(Of Spot),
    ByRef tDeletes As HashSet(Of String),
    ByVal e As Exception,
    ByVal Cancelled As Boolean,
    ByVal State As Object)

            MyBase.New(e, Cancelled, State)

            If Not tSpots Is Nothing Then Me.Spots = tSpots
            If Not tDeletes Is Nothing Then Me.Deletes = tDeletes

        End Sub

    End Class

    Public Class CommentsCompletedEventArgs
        Inherits AsyncCompletedEventArgs

        Public DbFile As String = ""
        Public Comments As List(Of Comment)

        Public Sub New(
            ByRef tComments As List(Of Comment),
            ByVal e As Exception,
            ByVal Cancelled As Boolean,
            ByVal State As Object)

            MyBase.New(e, Cancelled, State)
            If Not tComments Is Nothing Then Me.Comments = tComments

        End Sub

    End Class


    Public Class rSaveSpots

        Public NewCats(10) As Integer
        Public SpotsAdded As Integer = 0
        Public SpotsDeleted As Integer = 0

    End Class

    Public Class rSaveComments

        Public CommentsAdded As Integer = 0
        Public CommentsDeleted As Integer = 0

    End Class

    Public Class SpotCat

        Public Name As String
        Public Tag As String
        Public Children As New Collection

        Public Function AddChild(ByVal sName As String) As Boolean

            Children.Add(sName)
            Return True

        End Function

    End Class

    Public Class SpotRowChild

        Public ID As Long
        Public Stamp As Integer
        Public SubCat As Integer
        Public ExtCat As Integer
        Public Title As String
        Public Poster As String
        Public Tag As String
        Public Filesize As Long
        Public Modulus As String

    End Class

    Public Class iSpotRow

        Implements SpotRow

        Private _Spot As SpotRowChild

        Public ReadOnly Property ID As Long Implements SpotRow.ID
            Get
                If Len(Spot.Title) = 0 Then Return 0
                Return Spot.ID
            End Get
        End Property

        Private ReadOnly Property Spot As SpotRowChild
            Get

                If _Spot Is Nothing Then
                    Return New SpotRowChild
                End If

                Return _Spot

            End Get
        End Property

        Private Function Naam(ByVal lInt As Long) As String

            Select Case lInt
                Case 1
                    Return "maandag"
                Case 2
                    Return "dinsdag"
                Case 3
                    Return "woensdag"
                Case 4
                    Return "donderdag"
                Case 5
                    Return "vrijdag"
                Case 6
                    Return "zaterdag"
                Case Else
                    Return "zondag"
            End Select

        End Function

        Public ReadOnly Property Leeftijd As String Implements SpotRow.Leeftijd

            Get

                If Len(Spot.Title) = 0 Then Return ""

                Dim xD As Date = Utils.EPOCH.AddSeconds(Spot.Stamp).ToLocalTime
                Dim xCom As Date = CDate(Now.ToString("yyyy-MM-dd"))

                Dim sRet As String = ""

                Dim sDiff2 As Long = DateDiff("s", xCom, xD)
                Dim sDiff As Long = DateDiff(DateInterval.Day, xD, xCom)

                If sDiff < 7 Then

                    If sDiff2 > 0 Then
                        sRet = "vandaag (" & xD.ToString("HH:mm") & ")"
                    Else
                        If sDiff2 > -(60 * 60 * 24) Then
                            sRet = "gisteren (" & xD.ToString("HH:mm") & ")"
                        Else
                            sRet = Naam(xD.DayOfWeek) & " (" & xD.ToString("HH:mm") & ")"
                        End If
                    End If

                Else

                    sRet = CStr(sDiff + 1) & " dagen (" & xD.ToString("HH:mm") & ")"

                End If

                Return sRet

            End Get

        End Property

        Public ReadOnly Property Datum As String Implements SpotRow.Datum

            Get
                If Len(Spot.Title) = 0 Then Return ""
                Dim xD As Date = Utils.EPOCH.AddSeconds(Spot.Stamp).ToLocalTime
                Return xD.ToString("dd-MM-yy (HH:mm)")

            End Get

        End Property

        Public ReadOnly Property Formaat As String Implements SpotRow.Formaat
            Get

                If Len(Spot.Title) = 0 Then Return ""
                Dim lCat As Byte = CByte(CStr(Spot.SubCat).Substring(0, 1))

                Try
                    Return Utils.TranslateCatShort(lCat, CByte(CStr(CLng(Spot.SubCat)).Substring(1)))
                Catch
                    Return ""
                End Try

            End Get
        End Property

        Public ReadOnly Property Genre As String Implements SpotRow.Genre

            Get
                If Len(Spot.Title) = 0 Then Return ""
                Dim sVal As String = Utils.ReturnInfo(Spot.ExtCat)
                If sVal Is Nothing Then sVal = ""
                Return sVal
            End Get

        End Property

        Public ReadOnly Property Afzender As String Implements SpotRow.Afzender
            Get
                If Len(Spot.Title) = 0 Then Return ""
                Return Utils.StripNonAlphaNumericCharacters(Spot.Poster)
            End Get
        End Property

        Public ReadOnly Property Tag As String Implements SpotRow.Tag

            Get
                If Len(Spot.Title) = 0 Then Return ""
                Return Utils.StripNonAlphaNumericCharacters(Spot.Tag)
            End Get

        End Property

        Public ReadOnly Property Omvang As String Implements SpotRow.Omvang

            Get
                If Len(Spot.Title) = 0 Then Return ""
                If Spot.Filesize > 0 Then
                    Return Utils.ConvertSize(Spot.Filesize)
                Else
                    Return ""
                End If
            End Get

        End Property

        Public ReadOnly Property Titel As String Implements SpotRow.Titel
            Get
                If Len(Spot.Title) = 0 Then Return ""

                If Spot.Title.Contains("&") Then
                    Return System.Net.WebUtility.HtmlDecode(Spot.Title)
                End If

                Return Spot.Title

            End Get
        End Property

        Public ReadOnly Property Modulus As String Implements SpotRow.Modulus

            Get
                Return Spot.Modulus
            End Get

        End Property

        Public Sub New(ByVal xSpot As SpotRowChild)
            MyBase.New()

            _Spot = xSpot

        End Sub

    End Class

    Public Class FetchCache

        Public Cache As New List(Of Long)
        Public CacheHash As New HashSet(Of Long)

    End Class

    Public Class Parameters

        Public Property MaxResults() As Integer
            Get
                Return 5000
            End Get
            Set
                Throw New NotImplementedException()
            End Set
        End Property

        'Public Property SortOrder() As String
        '    Get
        '        Throw New NotImplementedException()
        '    End Get
        '    Set
        '        Throw New NotImplementedException()
        '    End Set
        'End Property

        'Public Property SortColumn() As String
        '    Get
        '        Return "rowid"
        '    End Get
        '    Set
        '        Throw New NotImplementedException()
        '    End Set
        'End Property

        Public Property DatabaseMax() As Long
            Get
                Return 0
            End Get
            Set
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property DatabaseCount() As Long
            Get
                Return 0
            End Get
            Set
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property DatabaseCache() As Integer
            Get
                Return 72500
            End Get
            Set
                Throw New NotImplementedException()
            End Set
        End Property

        Public Sub Save()

            Throw New NotImplementedException()

        End Sub

        Public Property DatabaseFilter() As Long
            Get
                Return 0
            End Get
            Set
                Throw New NotImplementedException()
            End Set
        End Property

    End Class

End Namespace
