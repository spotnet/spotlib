Imports System.Xml
Imports System.Text
Imports System.ComponentModel
Imports System.Globalization
Imports System.Security.Cryptography

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
                    Return Spotz.TranslateCatShort(lCat, CByte(CStr(CLng(Spot.SubCat)).Substring(1)))
                Catch
                    Return ""
                End Try

            End Get
        End Property

        Public ReadOnly Property Genre As String Implements SpotRow.Genre

            Get
                If Len(Spot.Title) = 0 Then Return ""
                Dim sVal As String = Spotz.ReturnInfo(Spot.ExtCat)
                If sVal Is Nothing Then sVal = ""
                Return sVal
            End Get

        End Property

        Public ReadOnly Property Afzender As String Implements SpotRow.Afzender
            Get
                If Len(Spot.Title) = 0 Then Return ""
                Return Utils.StripChars(Spot.Poster)
            End Get
        End Property

        Public ReadOnly Property Tag As String Implements SpotRow.Tag

            Get
                If Len(Spot.Title) = 0 Then Return ""
                Return Utils.StripChars(Spot.Tag)
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

    Public Enum ServerType
        Headers
        Upload
        Download
    End Enum

    Public Class ServerInfo

        Public Port As Integer = 119
        Public SSL As Boolean = False
        Public Username As String = ""
        Public Password As String = ""
        Public Server As String = ""
        Public Connections As Integer = 1

    End Class

    Public Class ProviderItem

        Public Name As String = ""
        Public Address As String = ""
        Public Port As Long

        Public Overrides Function ToString() As String
            Return Name
        End Function

    End Class

    Public Interface iWorkParams

        Sub Save()
        Function MaxResults() As Integer
        Sub SetMaxResults(vRes As Integer)
        Function DatabaseMax() As Long
        Sub SetDatabaseMax(vRes As Long)
        Function DatabaseCount() As Long
        Sub SetDatabaseCount(vRes As Long)
        Function DatabaseCache() As Integer
        Sub SetDatabaseCache(vRes As Integer)
        Function DatabaseFilter() As Long
        Sub SetDatabaseFilter(vRes As Long)

    End Interface

    Public Class Spotz

        Public Const SearchM As String = "Zoeken: "
        Public Const CancelMSG As String = "Geannuleerd"
        Public Const Spotname As String = "Spotnet"
        Public Const DefaultFilter As String = "cat < 9"
        Public Const MsgDomain As String = "spot.net"

        Public Shared Function IsNZB(ByVal _FileLnk As String, ByRef sSize As Long) As Boolean

            Dim m_xmld As XmlDocument

            Try

                If _FileLnk Is Nothing Then Return False
                If Len(_FileLnk) = 0 Then Return False

                m_xmld = New XmlDocument()
                m_xmld.XmlResolver = Nothing
                m_xmld.LoadXml(_FileLnk)

                Dim namespaceManager As XmlNamespaceManager = New XmlNamespaceManager(m_xmld.NameTable)
                namespaceManager.AddNamespace("pf", "http://www.newzbin.com/DTD/2003/nzb")

                If m_xmld.DocumentElement.Name = "nzb" Then
                    For Each dx As XmlNode In m_xmld.SelectNodes("/pf:nzb/pf:file/pf:segments/pf:segment", namespaceManager)
                        sSize += CLng(dx.Attributes("bytes").InnerXml)
                    Next
                    Return (sSize > 0)
                End If

                Return False

            Catch ex As Exception
                Return False
            End Try

        End Function

        Public Shared Function MakeUnique(ByVal sModulus As String) As String

            If Len(sModulus) = 0 Then Return "Onbekend"

            Try
                Return Utils.StripChars(Convert.ToBase64String(BitConverter.GetBytes(Utils.GetCrc.Calculate(Convert.FromBase64String(sModulus)))))
            Catch
                Return "Onbekend"
            End Try

        End Function

        Friend Shared Function CheckHash(ByVal sMsg As String) As Boolean

            Dim ShA As New SHA1Managed
            Dim ShABytes() As Byte = ShA.ComputeHash(Utils.MakeLatin(sMsg))

            If ShABytes(0) = 0 Then
                If ShABytes(1) = 0 Then

                    Return True

                End If
            End If

            Return False

        End Function

        Private Shared Function GetBinsearch() As String

            Throw New NotImplementedException()

        End Function

        Private Shared Function GetNZB(ByVal zPostData As String, ByVal OrgUrl As String, ByRef zError As String) As String

            Dim sReturn As String
            Dim bytArguments As Byte()

            bytArguments = Utils.MakeLatin(zPostData)

            Try

                Dim myRequest As Net.HttpWebRequest = CType(Net.WebRequest.Create("http://" & GetBinsearch() & "/fcgi/nzb.fcgi?" & OrgUrl), Net.HttpWebRequest)

                myRequest.Proxy = Nothing
                myRequest.ServicePoint.Expect100Continue = False
                myRequest.Method = "POST"
                myRequest.ContentType = "application/x-www-form-urlencoded"
                myRequest.ContentLength = bytArguments.Length

                Dim newStream As IO.Stream = myRequest.GetRequestStream()
                newStream.Write(bytArguments, 0, bytArguments.Length)
                newStream.Close()

                Dim reader As New IO.StreamReader(myRequest.GetResponse().GetResponseStream, Utils.LatinEnc)

                sReturn = Utils.UnGZIP(Utils.MakeLatin(reader.ReadToEnd))

                If IsNZB(sReturn, 0) Then Return sReturn

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Binsearch GetNZB Error")
                zError = ex.Message
            End Try

            Return vbNullString

        End Function

        Public Shared Function CheckUserSignature(ByVal sOrg As String, ByVal sSignature As String, ByVal sUserKey As String) As Boolean

            Try

                Dim RSAalg As RSACryptoServiceProvider = MakeRSA(sUserKey)

                If RSAalg Is Nothing Then Return False

                Return RSAalg.VerifyHash((New SHA1Managed).ComputeHash(Utils.MakeLatin(sOrg)), Nothing, Convert.FromBase64String(Utils.FixPadding(sSignature)))

            Catch
            End Try

            Return False

        End Function

        Friend Shared Function CreateWork(ByVal sFirst As Long, ByVal sLast As Long) As List(Of NNTPWork)

            Dim xWork As NNTPWork
            Dim WorkCol As New List(Of NNTPWork)

            Dim TotalMSG As Long
            Dim MaxHeaders As Long

            Dim FastPos As Long = 0
            Dim MinPacket As Long = 100
            Dim MaxPacket As Long = 9500
            Dim MaxProgress As Long = 100
            Dim FastPacket As Long = MinPacket

            TotalMSG = (sLast - sFirst) + 1

            If TotalMSG < 1 Then Return New List(Of NNTPWork)

            If TotalMSG > FastPacket Then
                FastPos = sFirst + (TotalMSG - FastPacket)
            Else
                FastPacket = 0
            End If

            MaxHeaders = CLng(Math.Ceiling((TotalMSG - FastPacket) / MaxProgress))

            If MaxHeaders > MaxPacket Then MaxHeaders = MaxPacket
            If MaxHeaders < MinPacket Then MaxHeaders = MinPacket
            If MaxHeaders > (TotalMSG - FastPacket) Then MaxHeaders = (TotalMSG - FastPacket)

            Dim TF As Long = sFirst
            Dim NextWork As Long = MaxHeaders
            Dim DoSplit As Long = CLng(Math.Ceiling((TotalMSG - FastPacket) / MaxHeaders)) + 1

            For lx = 1 To DoSplit

                If TF > sLast Then Continue For

                If (TF + NextWork) > sLast Then
                    NextWork = (sLast - TF) + 1
                End If

                If FastPacket > 0 Then

                    If TF >= FastPos Then Continue For

                    If (TF + NextWork) >= FastPos Then
                        NextWork = (FastPos - TF)
                    End If

                End If

                If NextWork > 0 Then

                    xWork = New NNTPWork

                    xWork.xDone = False
                    xWork.xStart = TF
                    xWork.xEnd = xWork.xStart + (NextWork - 1)

                    WorkCol.Add(xWork)

                    TF = TF + NextWork
                    NextWork = MaxHeaders

                End If

            Next

            If FastPacket > 0 Then

                xWork = New NNTPWork

                xWork.xDone = False
                xWork.xStart = FastPos
                xWork.xEnd = xWork.xStart + (FastPacket - 1)

                WorkCol.Add(xWork)

            End If

            WorkCol.Reverse()

            Dim cTotal As Long = 0

            For Each xW As NNTPWork In WorkCol
                cTotal += (xW.xEnd - xW.xStart) + 1
            Next

            If TotalMSG <> cTotal Then
                Throw New Exception("TotalMSG (" & CStr(TotalMSG) & ") <> cTotal (" & CStr(cTotal) & ")")
            End If

            Return WorkCol

        End Function

        Friend Shared Function FormatLong(ByVal zLong As Long) As String

            Dim sT As String
            sT = FormatLong2(zLong)
            If sT = "0" Then Return "Geen"
            Return sT

        End Function

        Friend Shared Function FormatLong2(ByVal zLong As Long) As String

            If zLong = 0 Then Return "0"
            Return zLong.ToString("#,#", CultureInfo.InvariantCulture).Replace(",", ".")

        End Function

        Friend Shared Function GetHash(ByVal zPostData As String, ByRef SignServer As String, ByVal SignCmd As String, ByRef zError As String) As String

            Dim sLoc As String = ""
            Dim sReturn As String = ""
            Dim bytArguments As Byte()
            Dim oWeb As System.Net.WebClient

            Try

                oWeb = New System.Net.WebClient
                oWeb.Headers.Add("Content-Type", "application/x-www-form-urlencoded")

                Dim iPos As Integer

                If Not Utils.IsAscii(zPostData, iPos) Then
                    zError = "Postdata niet ASCII: " & Mid(zPostData, iPos, 10) & ".."
                    Return vbNullString
                End If

                bytArguments = Utils.MakeLatin(zPostData)

                Dim sDot As String = "."

                If Len(SignServer) = 0 Then
                    zError = "Geen server opgegeven."
                    Return vbNullString
                End If

                sLoc = Utils.GetLocation("http://" & SignServer.Replace("http:", "").Replace("/", "") & "/")

                If Len(sLoc) = 0 Then

                    zError = "Kan de server niet bereiken. Probeer het later nog eens."
                    Return vbNullString

                End If

                Try

                    sReturn = Utils.GetLatin(oWeb.UploadData("http://" & sLoc & "/" & SignCmd, "POST", bytArguments))

                Catch ex As Exception

                    zError = "Kan de server niet bereiken (" & ex.Message & "). Probeer het later nog eens."
                    Return vbNullString

                End Try

                oWeb.Dispose()
                oWeb = Nothing

                If Not Utils.IsAscii(sReturn, iPos) Then
                    zError = "Response niet ASCII: " & Mid(sReturn, iPos, 10) & ".."
                    Return vbNullString
                End If

                Return sReturn

            Catch ex As Exception
                zError = "GetHash:: " & ex.Message
                Return vbNullString
            End Try

        End Function

        Friend Shared Function CreateUserSignature(ByVal sDataIn As String, ByVal cRSA As RSACryptoServiceProvider) As String

            Return Utils.SpecialString(Convert.ToBase64String(cRSA.SignHash((New SHA1Managed).ComputeHash(Utils.MakeLatin(sDataIn)), Nothing)))

        End Function

        Friend Shared Function CheckFrom(ByVal sFrom As String) As Boolean

            If sFrom.Trim.ToLower = "god" Then Return False
            If sFrom.Trim.ToLower = "mod" Then Return False
            If sFrom.Trim.ToLower = "modje" Then Return False
            If sFrom.Trim.ToLower = "spot" Then Return False
            If sFrom.Trim.ToLower = "spotje" Then Return False
            If sFrom.Trim.ToLower = "spotmod" Then Return False
            If sFrom.Trim.ToLower = "admin" Then Return False
            If sFrom.Trim.ToLower = "drazix" Then Return False
            If sFrom.Trim.ToLower = "moderator" Then Return False
            If sFrom.Trim.ToLower = "superuser" Then Return False
            If sFrom.Trim.ToLower = "supervisor" Then Return False
            If sFrom.Trim.ToLower = "spotnet" Then Return False
            If sFrom.Trim.ToLower = "spotned" Then Return False
            If sFrom.Trim.ToLower = "spotnetmod" Then Return False
            If sFrom.Trim.ToLower = "administrator" Then Return False

            Return True

        End Function

        Friend Shared Function PostData(ByVal tPhuse As Phuse.Engine, ByRef zInput As List(Of String), ByVal zSub As String, ByVal zFrom As String, ByVal zGroup As String, ByVal zExtra As String, ByRef xOutID As String, ByVal MsgID As String, ByRef sError As String) As Boolean

            Dim cNNTP As cNNTP
            Dim XC As Long = 0

            Dim sRet As Boolean = False
            Dim zNZB As String = vbNullString
            Dim sTmp As String = vbNullString
            Dim xOutList As String = vbNullString
            Dim TheData As String = vbNullString

            Dim TempId As String = ""

            cNNTP = New cNNTP(tPhuse)

            For Each NX As String In zInput

                XC += 1
                sError = vbNullString

                If Len(MsgID) = 0 Then

                    TempId = CreateMsgID()

                Else

                    TempId = MakeMsg(MsgID, True)

                End If

                Dim sHeader As String = "From: " & zFrom & vbCrLf & "Subject: " & zSub & CStr(Utils.sIIF((zInput.Count > 1), " [" & XC & "/" & zInput.Count & "]", "")) & vbCrLf & "Newsgroups: " & zGroup & vbCrLf & "Message-ID: " & TempId & vbCrLf & zExtra & "Content-Type: text/plain; charset=ISO-8859-1" & vbCrLf & "Content-Transfer-Encoding: 8bit"

                sRet = cNNTP.PostData(zGroup, sHeader & vbCrLf & vbCrLf & NX & ".", "", 0, sError)

                If sRet Then

                    xOutList += MakeMsg(TempId, False) & " "

                Else

                    cNNTP = Nothing
                    Return False

                End If

            Next

            cNNTP = Nothing
            xOutID = xOutList.Trim

            Return True

        End Function

        Friend Shared Function MakeP(ByVal sIn As String) As String

            Return Convert.ToBase64String(Utils.MakeLatin(sIn)).Replace("=", "%3d").Replace("+", "%2b").Replace("&", "%26").Replace$("/", "%2f").Trim

        End Function

        Friend Shared Function GetBinary(ByVal tPhuse As Phuse.Engine, ByVal Newsgroup As String, ByVal xMsgID As List(Of String), ByRef sxOut() As Byte, ByRef sError As String) As Boolean

            Dim cNNTP As cNNTP
            Dim Retr As New StringBuilder

            If xMsgID Is Nothing Then Return False
            If xMsgID.Count < 1 Then Return False

            cNNTP = New cNNTP(tPhuse)

            For Each ID As String In xMsgID

                Dim xRes As String = vbNullString
                sError = vbNullString

                Dim xRet As Integer = -1

                If Not cNNTP.GetBody(Newsgroup, ID, xRes, xRet, sError) Then

                    If (xRet = 430) Then
                        sError = "Kan de NZB (nog) niet vinden, als deze spot nog niet zo oud is, probeer het dan later nog eens. Soms heeft je provider wat vertraging."
                    End If

                    Return False

                End If

                If xRes.Substring(xRes.Length - 3) <> "." & vbCrLf Then
                    sError = "Kan artikel niet downloaden: Code 4"
                    Return False
                End If

                Dim xPos As Integer = xRes.IndexOf(vbCrLf) + 2

                xRes = xRes.Substring(xPos, xRes.Length - xPos - 5)
                xRes = xRes.Replace(vbCrLf & "..", ".").Replace(vbCrLf, vbNullString)

                Retr.Append(xRes)

            Next

            sxOut = Utils.MakeLatin(Retr.ToString.Replace("=C", vbLf).Replace("=B", vbCr).Replace("=A", Chr(0)).Replace("=D", "="))

            Return True

        End Function

        Friend Shared Function GetRSA(ByVal TrustedKeys() As String) As RSACryptoServiceProvider()

            Dim BoundTrust As Integer = UBound(TrustedKeys)
            Dim RSAalg(BoundTrust) As RSACryptoServiceProvider

            For cI As Integer = 0 To BoundTrust
                RSAalg(cI) = Nothing
                If Len(TrustedKeys(cI)) > 0 Then
                    Try
                        RSAalg(cI) = MakeRSA(TrustedKeys(cI))
                    Catch
                        ''
                    End Try
                End If
            Next

            Return RSAalg

        End Function

        Friend Shared Function MakeRSA(ByVal sModulus As String) As RSACryptoServiceProvider

            Dim cRSA As Security.Cryptography.RSACryptoServiceProvider = Nothing

            If Len(sModulus) = 0 Then Return Nothing

            Try

                Dim Exp(2) As Byte
                Dim pRsa As New RSAParameters()

                Exp(0) = 1
                Exp(2) = 1

                pRsa.Exponent = Exp
                pRsa.Modulus = Convert.FromBase64String(sModulus)

                cRSA = New RSACryptoServiceProvider()
                cRSA.ImportParameters(pRsa)

                Return cRSA

            Catch ex As Exception

                Return Nothing

            End Try

        End Function

        Friend Shared Function IsEro(ByVal sVal As String) As Boolean

            If sVal.Contains("d2") Then
                If sVal.Contains("d23|") Then Return True
                If sVal.Contains("d24|") Then Return True
                If sVal.Contains("d25|") Then Return True
                If sVal.Contains("d26|") Then Return True
            End If

            If sVal.Contains("d7") Then
                If sVal.Contains("d72|") Then Return True
                If sVal.Contains("d73|") Then Return True
                If sVal.Contains("d74|") Then Return True
                If sVal.Contains("d75|") Then Return True
            End If

            If sVal.Contains("z3|") Then Return True

            Return False

        End Function

        Friend Shared Function IsEbook(ByVal sVal As String) As Boolean

            If sVal.Contains("a5|") Then Return True
            If sVal.Contains("z2|") Then Return True

            Return False

        End Function

        Friend Shared Function IsTV(ByVal sVal As String) As Boolean

            If sVal.Contains("b4|") Then Return True
            If sVal.Contains("d11|") Then Return True
            If sVal.Contains("z1|") Then Return True

            Return False

        End Function
        Public Shared Function MakeMsg(ByVal sMes As String, Optional ByVal Tag As Boolean = True) As String

            If sMes.Substring(0, 1) = "<" Then
                If Tag Then Return sMes
            Else
                If Not Tag Then Return sMes
            End If

            If Tag Then
                Return "<" & sMes & ">"
            Else
                Return sMes.Substring(1, sMes.Length - 2)
            End If

        End Function
        Public Shared Function CatDesc(ByVal hCat As Byte, Optional ByVal zCat As Byte = 0) As String

            Select Case hCat

                Case 1

                    If zCat < 2 Then Return "Films"

                    Select Case zCat

                        Case 2
                            Return "Series"

                        Case 3
                            Return "Boeken"

                        Case 4
                            Return "Erotiek"

                    End Select

                Case 2

                    If zCat < 2 Then Return "Muziek"

                    Select Case zCat

                        Case 2
                            Return "Liveset"

                        Case 3
                            Return "Podcast"

                        Case 4
                            Return "Audiobook"

                    End Select

                Case 3
                    Return "Spellen"

                Case 4
                    Return "Applicaties"

                Case 5
                    Return "Boeken"

                Case 6
                    Return "Series"

                Case 9
                    Return "Erotiek"

            End Select

            Return "Fout"

        End Function

        Public Shared Function TranslateCat(ByVal hCat As Long, ByVal sCat As String, Optional ByVal bStrict As Boolean = False) As String

            If sCat.Length < 2 Then Return ""

            Select Case hCat

                Case 2

                    Select Case sCat.Substring(0, 1)

                        Case "a"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "MP3"
                                Case 1 : Return "WMA"
                                Case 2 : Return "WAV"
                                Case 3 : Return "OGG"
                                Case 4 : Return "EAC"
                                Case 5 : Return "DTS"
                                Case 6 : Return "AAC"
                                Case 7 : Return "APE"
                                Case 8 : Return "FLAC"
                                Case Else : Return ""
                            End Select

                        Case "b"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "CD"
                                Case 1 : Return "Radio"
                                Case 3 : Return "DVD"
                                Case 5 : Return "Vinyl"
                                Case 2 : Return Utils.sIIF(bStrict, "", "Compilatie")
                                Case 4 : Return "" ''"Anders"
                                Case 6 : Return "Stream"
                                Case Else : Return ""
                            End Select

                        Case "c"

                            Select Case CInt(sCat.Substring(1))
                                Case 1 : Return "< 96kbit"
                                Case 2 : Return "96kbit"
                                Case 3 : Return "128kbit"
                                Case 4 : Return "160kbit"
                                Case 5 : Return "192kbit"
                                Case 6 : Return "256kbit"
                                Case 7 : Return "320kbit"
                                Case 8 : Return "Lossless"
                                Case 0 : Return "Variabel"
                                Case 9 : Return "" ''"Anders"
                                Case Else : Return ""
                            End Select

                        Case "d"
                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Blues"
                                Case 1 : Return "Compilatie"
                                Case 2 : Return "Cabaret"
                                Case 3 : Return "Dance"
                                Case 4 : Return "Diversen"
                                Case 5 : Return "Hardstyle"
                                Case 6 : Return "Wereld"
                                Case 7 : Return "Jazz"
                                Case 8 : Return "Jeugd"
                                Case 9 : Return "Klassiek"
                                Case 10 : Return Utils.sIIF(bStrict, "", "Kleinkunst")
                                Case 11 : Return "Hollands"
                                Case 12 : Return Utils.sIIF(bStrict, "", "New Age")
                                Case 13 : Return "Pop"
                                Case 14 : Return "RnB"
                                Case 15 : Return "Hiphop"
                                Case 16 : Return "Reggae"
                                Case 17 : Return "Religieus"
                                Case 18 : Return "Rock"
                                Case 19 : Return "Soundtrack"
                                Case 20 : Return "" ''"Anders"
                                Case 21 : Return Utils.sIIF(bStrict, "", "Hardstyle")
                                Case 22 : Return Utils.sIIF(bStrict, "", "Aziatisch")
                                Case 23 : Return "Disco"
                                Case 24 : Return "Classics"
                                Case 25 : Return "Metal"
                                Case 26 : Return "Country"
                                Case 27 : Return "Dubstep"
                                Case 28 : Return Utils.sIIF(bStrict, "", "Nederhop")
                                Case 29 : Return "DnB"
                                Case 30 : Return "Electro"
                                Case 31 : Return "Folk"
                                Case 32 : Return "Soul"
                                Case 33 : Return "Trance"
                                Case 34 : Return "Balkan"
                                Case 35 : Return "Techno"
                                Case 36 : Return "Ambient"
                                Case 37 : Return "Latin"
                                Case 38 : Return "Live"
                                Case Else : Return ""
                            End Select

                        Case "z"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Album"
                                Case 1 : Return "Liveset"
                                Case 2 : Return "Podcast"
                                Case 3 : Return "Luisterboek"
                                Case Else : Return ""
                            End Select

                        Case Else

                            Return ""

                    End Select

                Case 3

                    Select Case sCat.Substring(0, 1)

                        Case "a"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Windows"
                                Case 1 : Return "Macintosh"
                                Case 2 : Return "Linux"
                                Case 3 : Return "Playstation"
                                Case 4 : Return "Playstation 2"
                                Case 5 : Return "PSP"
                                Case 6 : Return "XBox"
                                Case 7 : Return "XBox 360"
                                Case 8 : Return "Gameboy Advance"
                                Case 9 : Return "Gamecube"
                                Case 10 : Return "Nintendo DS"
                                Case 11 : Return "Nintendo Wii"
                                Case 12 : Return "Playstation 3"
                                Case 13 : Return "Windows Phone"
                                Case 14 : Return "iOs"
                                Case 15 : Return "Android"
                                Case 16 : Return Utils.sIIF(bStrict, "", "Nintendo 3DS")
                                Case Else : Return ""
                            End Select

                        Case "b"

                            Select Case CInt(sCat.Substring(1))
                                Case 1 : Return "Rip"
                                Case 0 : Return Utils.sIIF(bStrict, "", "ISO")
                                Case 2 : Return "Retail"
                                Case 3 : Return "DLC"
                                Case 4 : Return "" ''"Anders"
                                Case 5 : Return "Patch"
                                Case 6 : Return "Crack"
                                Case Else : Return ""
                            End Select

                        Case "c"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Actie"
                                Case 1 : Return "Avontuur"
                                Case 2 : Return "Strategie"
                                Case 3 : Return "Rollenspel"
                                Case 4 : Return "Simulatie"
                                Case 5 : Return "Race"
                                Case 6 : Return "Vliegen"
                                Case 7 : Return "Shooter"
                                Case 8 : Return "Platform"
                                Case 9 : Return "Sport"
                                Case 10 : Return "Jeugd"
                                Case 11 : Return "Puzzel"
                                Case 12 : Return "" ''"Anders"
                                Case 13 : Return "Bordspel"
                                Case 14 : Return "Kaarten"
                                Case 15 : Return "Educatie"
                                Case 16 : Return "Muziek"
                                Case 17 : Return "Party"
                                Case Else : Return ""
                            End Select

                        Case Else

                            Return ""

                    End Select

                Case 4

                    Select Case sCat.Substring(0, 1)

                        Case "a"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Windows"
                                Case 1 : Return "Macintosh"
                                Case 2 : Return "Linux"
                                Case 3 : Return "OS/2"
                                Case 4 : Return "Windows Phone"
                                Case 5 : Return "Navigatie"
                                Case 6 : Return "iOs"
                                Case 7 : Return "Android"
                                Case Else : Return ""
                            End Select

                        Case "b"
                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Audio"
                                Case 1 : Return "Video"
                                Case 2 : Return "Grafisch"
                                Case 3 : Return Utils.sIIF(bStrict, "", "CD/DVD Tools")
                                Case 4 : Return Utils.sIIF(bStrict, "", "Media spelers")
                                Case 5 : Return Utils.sIIF(bStrict, "", "Rippers & Encoders")
                                Case 6 : Return Utils.sIIF(bStrict, "", "Plugins")
                                Case 7 : Return Utils.sIIF(bStrict, "", "Database tools")
                                Case 8 : Return Utils.sIIF(bStrict, "", "Email software")
                                Case 9 : Return "Foto"
                                Case 10 : Return Utils.sIIF(bStrict, "", "Screensavers")
                                Case 11 : Return Utils.sIIF(bStrict, "", "Skin software")
                                Case 12 : Return Utils.sIIF(bStrict, "", "Drivers")
                                Case 13 : Return Utils.sIIF(bStrict, "", "Browsers")
                                Case 14 : Return Utils.sIIF(bStrict, "", "Download managers")
                                Case 15 : Return "Download"
                                Case 16 : Return Utils.sIIF(bStrict, "", "Usenet software")
                                Case 17 : Return Utils.sIIF(bStrict, "", "RSS Readers")
                                Case 18 : Return Utils.sIIF(bStrict, "", "FTP software")
                                Case 19 : Return Utils.sIIF(bStrict, "", "Firewalls")
                                Case 20 : Return Utils.sIIF(bStrict, "", "Antivirus software")
                                Case 21 : Return Utils.sIIF(bStrict, "", "Antispyware software")
                                Case 22 : Return Utils.sIIF(bStrict, "", "Optimalisatiesoftware")
                                Case 23 : Return "Beveiliging"
                                Case 24 : Return "Systeem"
                                Case 25 : Return "" ''"Anders"
                                Case 26 : Return "Educatief"
                                Case 27 : Return "Kantoor"
                                Case 28 : Return "Internet"
                                Case 29 : Return "Communicatie"
                                Case 30 : Return "Ontwikkel"
                                Case 31 : Return "Spotnet"
                                Case Else : Return ""
                            End Select

                        Case Else

                            Return ""

                    End Select

                Case Else

                    Select Case sCat.Substring(0, 1)

                        Case "a"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "DivX"
                                Case 1 : Return "WMV"
                                Case 2 : Return "MPG"
                                Case 3 : Return "DVD5"
                                Case 4 : Return Utils.sIIF(bStrict, "", "HD Overig")
                                Case 5 : Return "ePub"
                                Case 6 : Return "Bluray"
                                Case 7 : Return Utils.sIIF(bStrict, "", "HD-DVD")
                                Case 8 : Return Utils.sIIF(bStrict, "", "WMV HD")
                                Case 9 : Return "x264"
                                Case 10 : Return "DVD9"
                                Case Else : Return ""
                            End Select

                        Case "b"

                            Select Case CInt(sCat.Substring(1))
                                Case 4 : Return Utils.sIIF(bStrict, "", "TV")
                                Case 1 : Return Utils.sIIF(bStrict, "", "(S)VCD")
                                Case 6 : Return Utils.sIIF(bStrict, "", "Satelliet")
                                Case 2 : Return Utils.sIIF(bStrict, "", "Promo")
                                Case 3 : Return "Retail"
                                Case 7 : Return "R5"
                                Case 0 : Return "Cam"
                                Case 8 : Return Utils.sIIF(bStrict, "", "Telecine")
                                Case 9 : Return "Telesync"
                                Case 5 : Return "" '' "Anders"
                                Case 10 : Return "Scan"
                                Case Else : Return ""
                            End Select

                        Case "c"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Geen ondertitels"
                                Case 3 : Return "Engels ondertiteld (extern)"
                                Case 4 : Return Utils.sIIF(hCat <> 5, "Engels ondertiteld (ingebakken)", "Engels geschreven")
                                Case 7 : Return "Engels ondertiteld (instelbaar)"
                                Case 1 : Return "Nederlands ondertiteld (extern)"
                                Case 2 : Return Utils.sIIF(hCat <> 5, "Nederlands ondertiteld (ingebakken)", "Nederlands geschreven")
                                Case 6 : Return "Nederlands ondertiteld (instelbaar)"
                                Case 10 : Return "Engels gesproken"
                                Case 11 : Return "Nederlands gesproken"
                                Case 12 : Return Utils.sIIF(hCat <> 5, "Duits gesproken", "Duits geschreven")
                                Case 13 : Return Utils.sIIF(hCat <> 5, "Frans gesproken", "Frans geschreven")
                                Case 14 : Return Utils.sIIF(hCat <> 5, "Spaans gesproken", "Spaans geschreven")
                                Case 5 : Return ""  '' "Anders"
                                Case Else : Return ""
                            End Select

                        Case "d"

                            Select Case CInt(sCat.Substring(1))

                                Case 0 : Return "Actie"
                                Case 29 : Return "Anime"
                                Case 2 : Return "Animatie"
                                Case 28 : Return "Aziatisch"
                                Case 1 : Return "Avontuur"
                                Case 3 : Return "Cabaret"
                                Case 32 : Return "Cartoon"
                                Case 6 : Return "Documentaire"
                                Case 7 : Return "Drama"
                                Case 8 : Return "Familie"
                                Case 9 : Return "Fantasie"
                                Case 10 : Return "Filmhuis"
                                Case 12 : Return "Horror"
                                Case 33 : Return "Jeugd"
                                Case 4 : Return "Komedie"
                                Case 19 : Return "Kort"
                                Case 5 : Return "Misdaad"
                                Case 13 : Return "Muziek"
                                Case 14 : Return "Musical"
                                Case 15 : Return "Mysterie"
                                Case 21 : Return "Oorlog"
                                Case 16 : Return "Romantiek"
                                Case 17 : Return "Science Fiction"
                                Case 18 : Return "Sport"
                                Case 11 : Return Utils.sIIF(bStrict, "", "Televisie")
                                Case 20 : Return "Thriller"
                                Case 22 : Return "Western"

                                Case 23 : Return "Hetero"
                                Case 24 : Return "Homo"
                                Case 25 : Return "Lesbo"
                                Case 26 : Return "Bi"

                                Case 27 : Return "" '' "Anders"

                                Case 30 : Return "Cover"
                                Case 43 : Return "Dagblad"
                                Case 44 : Return "Tijdschrift"
                                Case 31 : Return "Stripboek"

                                Case 34 : Return "Zakelijk"
                                Case 35 : Return "Computer"
                                Case 36 : Return "Hobby"
                                Case 37 : Return "Koken"
                                Case 38 : Return "Knutselen"
                                Case 39 : Return "Handwerk"
                                Case 40 : Return "Gezondheid"
                                Case 41 : Return "Historie"
                                Case 42 : Return "Psychologie"
                                Case 45 : Return "Wetenschap"
                                Case 46 : Return "Vrouw"
                                Case 47 : Return "Religie"
                                Case 48 : Return "Roman"
                                Case 49 : Return "Biografie"
                                Case 50 : Return "Detective"
                                Case 51 : Return "Dieren"
                                Case 52 : Return ""
                                Case 53 : Return "Reizen"
                                Case 54 : Return "Waargebeurd"
                                Case 55 : Return "Non-fictie"
                                Case 57 : Return "Poezie"
                                Case 58 : Return "Sprookje"

                                Case 75 : Return Utils.sIIF(bStrict, "", "Hetero")
                                Case 74 : Return Utils.sIIF(bStrict, "", "Homo")
                                Case 73 : Return Utils.sIIF(bStrict, "", "Lesbo")
                                Case 72 : Return Utils.sIIF(bStrict, "", "Bi")

                                Case 76 : Return "Amateur"
                                Case 77 : Return "Groep"
                                Case 78 : Return "POV"
                                Case 79 : Return "Solo"
                                Case 80 : Return "Jong"
                                Case 81 : Return "Soft"
                                Case 82 : Return "Fetisj"
                                Case 83 : Return "Oud"
                                Case 84 : Return "BBW"
                                Case 85 : Return "SM"
                                Case 86 : Return "Hard"
                                Case 87 : Return "Donker"
                                Case 88 : Return "Hentai"
                                Case 89 : Return "Buiten"

                                Case Else : Return ""

                            End Select

                        Case "z"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Film"
                                Case 1 : Return "Serie"
                                Case 2 : Return "Boek"
                                Case 3 : Return "Erotiek"
                                Case Else : Return ""
                            End Select

                        Case Else

                            Return ""

                    End Select

            End Select

        End Function

        Public Shared Function TranslateCatDesc(ByVal hCat As Long, ByVal sCat As String) As String

            Select Case hCat
                Case 2
                    Select Case sCat.Substring(0, 1)
                        Case "a" : Return "Formaat"
                        Case "b" : Return "Bron"
                        Case "c" : Return "Bitrate"
                        Case "d" : Return "Genre"
                        Case "z" : Return "Categorie"
                        Case Else : Return ""
                    End Select
                Case 3
                    Select Case sCat.Substring(0, 1)
                        Case "a" : Return "Platform"
                        Case "b" : Return "Formaat"
                        Case "c" : Return "Genre"
                        Case "z" : Return "Categorie"
                        Case Else : Return ""
                    End Select
                Case 4
                    Select Case sCat.Substring(0, 1)
                        Case "a" : Return "Platform"
                        Case "b" : Return "Genre"
                        Case "z" : Return "Categorie"
                        Case Else : Return ""
                    End Select
                Case Else
                    Select Case sCat.Substring(0, 1)
                        Case "a" : Return "Formaat"
                        Case "b" : Return "Bron"
                        Case "c" : Return "Taal"
                        Case "d" : Return "Genre"
                        Case "z" : Return "Categorie"
                        Case Else : Return ""
                    End Select
            End Select

        End Function

        Public Shared Function TranslateCatShort(ByVal hCat As Integer, ByVal sCat As Integer) As String

            Select Case hCat

                Case 2

                    Select Case sCat
                        Case 0
                            Return "MP3"
                        Case 1
                            Return "WMA"
                        Case 2
                            Return "WAV"
                        Case 3
                            Return "OGG"
                        Case 4
                            Return "EAC"
                        Case 5
                            Return "DTS"
                        Case 6
                            Return "AAC"
                        Case 7
                            Return "APE"
                        Case 8
                            Return "FLAC"
                        Case Else
                            Return ""
                    End Select


                Case 3

                    Select Case sCat
                        Case 0
                            Return "Win"
                        Case 1
                            Return "Mac"
                        Case 2
                            Return "Linux"
                        Case 3
                            Return "PSX"
                        Case 4
                            Return "PS2"
                        Case 5
                            Return "PSP"
                        Case 6
                            Return "XBox"
                        Case 7
                            Return "360"
                        Case 8
                            Return "GBA"
                        Case 9
                            Return "GC"
                        Case 10
                            Return "NDS"
                        Case 11
                            Return "Wii"
                        Case 12
                            Return "PS3"
                        Case 13
                            Return "WP7"
                        Case 14
                            Return "iOs"
                        Case 15
                            Return "Android"
                        Case Else
                            Return ""
                    End Select

                Case 4

                    Select Case sCat
                        Case 0
                            Return "Win"
                        Case 1
                            Return "Mac"
                        Case 2
                            Return "Linux"
                        Case 3
                            Return "OS2"
                        Case 4
                            Return "WP7"
                        Case 5
                            Return "Navi"
                        Case 6
                            Return "iOs"
                        Case 7
                            Return "Android"
                        Case Else
                            Return ""
                    End Select

                Case Else

                    Select Case sCat
                        Case 0
                            Return "DivX"
                        Case 1
                            Return "WMV"
                        Case 2
                            Return "MPG"
                        Case 3
                            Return "DVD5"
                        Case 4
                            Return "HD"
                        Case 5
                            Return "ePub"
                        Case 6
                            Return "Bluray"
                        Case 7
                            Return "HD"
                        Case 8
                            Return "HD"
                        Case 9
                            Return "x264"
                        Case 10
                            Return "DVD9"
                        Case Else
                            Return ""
                    End Select

            End Select

        End Function

        Public Shared Sub Foutje(ByVal sMsg As String, Optional ByVal sCaption As String = "Error")

            Throw New Exception(sCaption + ": " + sMsg)

        End Sub

        Public Shared Function GetKey() As RSACryptoServiceProvider

            Static cKey As RSACryptoServiceProvider

            If cKey Is Nothing Then

                Dim CP As New CspParameters

                CP.KeyContainerName = Spotname & " User Key"
                CP.Flags = CType(CspProviderFlags.NoPrompt + CspProviderFlags.UseArchivableKey, CspProviderFlags)

                cKey = New RSACryptoServiceProvider(384, CP)

            End If

            Return cKey

        End Function

        Public Shared Function GetModulus() As String

            Return Convert.ToBase64String(GetKey.ExportParameters(False).Modulus)

        End Function

        Public Shared Function ReturnInfo(ByVal ExtCat As Integer) As String

            Dim FindCat As Char
            Dim sCat As String = CStr(CInt(ExtCat))

            Dim lCat As Byte = CByte(sCat.Substring(0, 1))
            Dim zCat As Byte = CByte(sCat.Substring(1))

            If zCat > 98 Then Return ""

            Select Case lCat
                Case 3
                    FindCat = "c"c
                Case 4
                    FindCat = "b"c
                Case Else
                    FindCat = "d"c
            End Select

            Return TranslateCat(lCat, FindCat & zCat)

        End Function

        Public Shared Function TranslateInfo(ByVal hCat As Integer, ByVal sCats As String) As Byte

            Dim FindCat As Char
            Dim xInd As Byte = 0
            Dim SubCat As Long = 0
            Dim lPos As Integer = -1
            Dim sFind As String = ""
            Dim sRet As String = ""
            Dim MaxCat As Byte = 100

            Static bCat0(MaxCat) As Boolean
            Static bCat1(MaxCat) As Boolean
            Static bCat2(MaxCat) As Boolean
            Static bCat3(MaxCat) As Boolean

            Static DidOnce As Boolean = False

            If Not DidOnce Then

                For i As Byte = 0 To MaxCat
                    sRet = TranslateCat(1, "d" & i)
                    bCat0(i) = (Len(sRet) > 0)
                Next

                For i As Byte = 0 To MaxCat
                    sRet = TranslateCat(2, "d" & i)
                    bCat1(i) = (Len(sRet) > 0)
                Next

                For i As Byte = 0 To MaxCat
                    sRet = TranslateCat(3, "c" & i)
                    bCat2(i) = (Len(sRet) > 0)
                Next

                For i As Byte = 0 To MaxCat
                    sRet = TranslateCat(4, "b" & i)
                    bCat3(i) = (Len(sRet) > 0)
                Next

                DidOnce = True

            End If

            Try

                Select Case hCat
                    Case 3
                        FindCat = "c"c
                    Case 4
                        FindCat = "b"c
                    Case Else
                        FindCat = "d"c
                End Select

                Do

                    lPos = sCats.IndexOf(FindCat, lPos + 1)

                    If lPos = -1 Then Exit Do

                    xInd = CByte(Val(sCats.Substring(lPos + 1, 2)))

                    If xInd > MaxCat Then Continue Do

                    If hCat = 6 And xInd = 11 Then Continue Do

                    If hCat = 9 Then

                        Select Case xInd

                            Case 23, 24, 25, 26, 72, 73, 74, 75

                                If sCats.IndexOf(FindCat, lPos + 1) > -1 Then Continue Do

                        End Select

                    End If

                    Select Case hCat

                        Case 1
                            If Not bCat1(xInd) Then Continue Do

                        Case 2
                            If Not bCat2(xInd) Then Continue Do

                        Case 3
                            If Not bCat3(xInd) Then Continue Do

                        Case Else
                            If Not bCat0(xInd) Then Continue Do

                    End Select

                    Return xInd

                Loop

            Catch ex As Exception
            End Try

            Return 99

        End Function

        Public Shared Function CreateMsgID(Optional ByVal sPrefix As String = "") As String

            Dim ZL(7) As Byte
            Dim ZK As New Random
            Dim sDomain As String = "spot.net"

            ZK.NextBytes(ZL)

            Dim Span As TimeSpan = (DateTime.UtcNow - Utils.EPOCH)
            Dim CurDate As Integer = CInt(Span.TotalSeconds)

            Dim sRandom As String = Convert.ToBase64String(ZL) & Convert.ToBase64String(BitConverter.GetBytes(CurDate))
            sRandom = sRandom.Replace("/", "s").Replace("+", "p").Replace("=", "")

            If Len(sPrefix) = 0 Then
                Return Utils.CreateHash("<" & sRandom, "@" & sDomain & ">")
            Else
                Return Utils.CreateHash("<" & sPrefix.Replace(".", "") & ".0." & sRandom & ".", "@" & sDomain & ">")
            End If

        End Function

        Friend Shared Sub ShowOnce(ByVal sMsg As String, ByVal sTitle As String)

            Static DidList As New HashSet(Of String)

            If DidList.Add(sMsg) Then
                MsgBox(sMsg, MsgBoxStyle.Information, sTitle)
            End If

        End Sub

    End Class

End Namespace
