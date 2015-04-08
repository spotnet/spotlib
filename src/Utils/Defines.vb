Imports System.ComponentModel

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

End Namespace
