Imports System.Xml
Imports System.Diagnostics
Imports System.ComponentModel
Imports System.Security.Cryptography

Friend Class Worker

    Public HeaderData() As String
    Public InstanceCount As Integer
    Public xSettings As NNTPSettings
    Public Operation As AsyncOperation

    Friend RSA() As RSACryptoServiceProvider

    Private xOutputData As List(Of Spot)

    Public Event WorkDone(ByVal Errors As Boolean, ByVal WorkDone As Integer, ByVal InstanceCount As Integer, ByVal sError As String, ByRef zOut As List(Of Spot), ByVal xTotal As Long, ByVal xCnt As Long, ByVal NoProgress As Boolean, ByVal xParam As NNTPSettings, ByVal asyncOp As AsyncOperation)

    Friend Sub ParseHeaders(ByVal Hw As Headers)

        Dim bRet As Boolean
        Dim sRet As String = ""
        Dim WorkDone As Integer = 0
        Dim xTotal As Long
        Dim xCnt As Long

        Try
            bRet = DoWork(WorkDone, sRet, xTotal, xCnt)
        Catch ex As Exception
            bRet = False
            sRet = ex.Message
        End Try

        Hw.xWorkDone(Not bRet, WorkDone, InstanceCount, sRet, xOutputData, xTotal, xCnt, False, xSettings, Operation)

    End Sub

    Friend Function ParseSpotXML(ByRef Lz As SpotEx, ByVal TheDoc As XmlDocument, ByVal XmlSignature As String, ByVal bCheck As Boolean, ByRef sError As String) As SpotEx

        Dim bNative As Boolean = False

        Try

            Dim RootNode As XmlNode = TheDoc.SelectSingleNode("Spotnet")
            If RootNode Is Nothing Then RootNode = TheDoc.SelectSingleNode("SpotNet")

            With RootNode.SelectSingleNode("Posting")

                Try
                    Dim nodeList As XmlNodeList = TheDoc.GetElementsByTagName("ID")
                    bNative = (nodeList.Count = 0)
                Catch
                    bNative = True
                End Try

                If Not bNative Then

                    Try
                        Dim nodeList As XmlNodeList = TheDoc.GetElementsByTagName("Key")
                        Dim Lx As XmlElement = CType(nodeList(0), XmlElement)
                        bNative = CByte(Lx.InnerText) <> 1
                    Catch
                    End Try

                End If

                If bNative Then Lz.OldInfo = Nothing

                If .SelectSingleNode("Description") Is Nothing Then sError = "E1" : Return Nothing

                Lz.Body = HtmlDecode(Trim(.SelectSingleNode("Description").InnerText))

                If Not .SelectSingleNode("Image") Is Nothing Then

                    Lz.Image = HtmlDecode(.SelectSingleNode("Image").InnerText)

                    If bNative Then
                        With .SelectSingleNode("Image")
                            If Not .Attributes("Width") Is Nothing Then
                                Lz.ImageWidth = CInt(Val(.Attributes("Width").InnerText))
                            End If
                            If Not .Attributes("Height") Is Nothing Then
                                Lz.ImageHeight = CInt(Val(.Attributes("Height").InnerText))
                            End If
                            For Each GN As XmlNode In .SelectNodes("Segment")
                                Lz.ImageID = MakeMsg(Trim(GN.InnerText), True)
                                Lz.Image = ""
                                Exit For
                            Next
                        End With
                    End If

                End If

                If Not .SelectSingleNode("Website") Is Nothing Then Lz.Web = HtmlDecode(.SelectSingleNode("Website").InnerText)

                If Not bNative Then

                    If Not .SelectSingleNode("Filename") Is Nothing Then Lz.OldInfo.FileName = .SelectSingleNode("Filename").InnerText

                    For Each GN As XmlNode In .SelectNodes("Newsgroup")
                        If GN.InnerText <> "Other" Then
                            Lz.OldInfo.Groups += GN.InnerText & "|"
                        End If
                    Next

                Else

                    If .SelectSingleNode("NZB") Is Nothing Then sError = "E3" : Return Nothing

                    For Each GN As XmlNode In .SelectSingleNode("NZB").SelectNodes("Segment")
                        Lz.NZB += GN.InnerText & " "
                    Next

                    Lz.NZB = Lz.NZB.Trim.Replace(Chr(34), "")

                    If Len(Lz.NZB) = 0 Then sError = "E4" : Return Nothing

                End If

                Lz.User.ValidSignature = False

                If (Not bCheck) Or (Lz.KeyID = 1) Then Return Lz

                If (Len(Lz.User.Signature) > 0) And (Len(Lz.User.Modulus) > 0) And (Len(Lz.MessageID) > 0) Then
                    Lz.User.ValidSignature = CheckUserSignature(MakeMsg(Lz.MessageID), Lz.User.Signature, Lz.User.Modulus)
                End If

                If Not Lz.User.ValidSignature Then
                    If (Len(Lz.User.Signature) > 0) And (Len(Lz.User.Modulus) > 0) And (Len(XmlSignature) > 0) Then
                        Lz.User.ValidSignature = CheckUserSignature(XmlSignature, Lz.User.Signature, Lz.User.Modulus)
                    End If
                End If

                If Not Lz.User.ValidSignature Then

                    If (Lz.KeyID <> 3) Or (Lz.Stamp > 1317617563) Or (Len(Lz.User.Modulus) > 0) Then

                        sError = "Ongeldige handtekening"
                        Return Nothing

                    End If

                End If

                Return Lz

            End With

        Catch ex As Exception

            sError = ex.Message
            Return Nothing

        End Try

    End Function

    Private Function CopyToEx(ByVal sX As Spot) As SpotEx

        Dim tX As New SpotEx

        With tX
            .Category = sX.Category
            .Filesize = sX.Filesize
            .KeyID = sX.KeyID
            .MessageID = sX.MessageID
            .Poster = sX.Poster
            .Stamp = sX.Stamp
            .SubCat = sX.SubCat
            .SubCats = sX.SubCats
            .Tag = sX.Tag
            .Title = sX.Title
            .Article = sX.Article
            .Modulus = sX.Modulus
        End With

        Return tX

    End Function

    Friend Function ParseSpot(ByVal sSubject As String, ByVal sFrom As String, ByVal sMessageID As String, ByVal xParam As NNTPSettings, ByRef sError As String) As SpotEx

        Dim hData(3) As String
        Dim tData(4) As String

        Try

            tData(0) = "1234" 
            tData(1) = sSubject
            tData(2) = sFrom
            tData(3) = "2010" 
            tData(4) = sMessageID

            hData(0) = "" 
            hData(1) = Join(tData, vbTab)
            hData(2) = "."
            hData(3) = ""

            InstanceCount = 1
            HeaderData = hData

            xSettings = xParam
            xSettings.CheckSignatures = True

            RSA = sModule.GetRSA(xParam.TrustedKeys)

            If Not DoWork(1, sError, 1, 1) Then Return Nothing
            If xOutputData.Count = 0 Then sError = "Ongeldige handtekening." : Return Nothing

            Return CopyToEx(xOutputData.Item(0))

        Catch ex As Exception

            sError = ex.Message
            Return Nothing

        End Try

    End Function

    Private Function DoWork(ByRef WorkDone As Integer, ByRef sError As String, ByRef sTotal As Long, ByRef xCnt As Long) As Boolean

        If UBound(HeaderData) = 2 Then Return True ' Leeg blok

        Dim xB As Integer
        Dim Segs() As String
        Dim xOD As Spot
        Dim lID As Integer
        Dim zzT() As String
        Dim HeaderInfo() As String
        Dim Catz As String
        Dim sCheck As String = ""
        Dim TempCat As String = ""
        Dim NativeSpot As Boolean = True
        Dim DateOffset As Integer = 25000

        Dim dPos As Integer = -1
        Dim AtPos As Integer = -1
        Dim GTtPos As Integer = -1
        Dim LTtPos As Integer = -1
        Dim sModulus As String = ""

        Dim iFound As Integer = 0
        Dim sSecondByte As Byte
        Dim iFirstChar As Integer
        Dim ShABytes() As Byte = Nothing

        Dim sHeader As String
        Dim sHeaderSign As String

        Dim ShA As New SHA1Managed
        Dim CatzCol As New List(Of String)
        Dim Span As TimeSpan = (DateTime.UtcNow - EPOCH)
        Dim MaxDate As Integer = CInt(Span.TotalSeconds) + DateOffset
        Dim cRSA As Security.Cryptography.RSACryptoServiceProvider = Nothing

        xOutputData = New List(Of Spot)
        xB = UBound(HeaderData) - 2

        sTotal = xB

        For xLoop As Integer = xB To 1 Step -1

            Segs = Split(HeaderData(xLoop), vbTab)
            ' article number, subject, author, date, message-id

            If UBound(Segs) < 4 Then Continue For
            If Len(Segs(1)) = 0 Then Continue For

            xOD = New Spot

            Try

                With xOD

                    .MessageID = Segs(4).Substring(1, Segs(4).Length - 2)

                    AtPos = Segs(2).IndexOf("@")
                    If AtPos < 1 Then Continue For

                    GTtPos = Segs(2).IndexOf(">")
                    LTtPos = Segs(2).IndexOf("<")

                    If LTtPos < 1 Then Continue For
                    If LTtPos > AtPos Then Continue For
                    If GTtPos < AtPos Then Continue For

                    sHeader = Segs(2).Substring(AtPos + 1, GTtPos - AtPos - 1)

                    HeaderInfo = sHeader.Split("."c)
                    If UBound(HeaderInfo) < 6 Then Continue For

                    If Not Byte.TryParse(HeaderInfo(0).Substring(1, 1), .KeyID) Then Continue For
                    If .KeyID < 1 Then Continue For

                    NativeSpot = .KeyID <> 1

                    If Not Long.TryParse(Segs(0), .Article) Then Continue For
                    If .Article < 1 Then Continue For

                    If Not Integer.TryParse(HeaderInfo(3), .Stamp) Then Continue For
                    If (.Stamp < 1218171600) Then Continue For

                    If Not NativeSpot Then

                        If Not Integer.TryParse(HeaderInfo(3), lID) Then Continue For

                        If lID < 1000000 Then Continue For
                        If (.Stamp > 1317617563) Then Continue For ' SPAM fix
                        If Not Segs(3).Contains("2010") Then Continue For ' SPAM fix

                    End If

                    If .Stamp > MaxDate Then .Stamp = MaxDate - DateOffset

                    If Not Long.TryParse(HeaderInfo(1), .Filesize) Then
                        .Filesize = 0
                    Else
                        If .Filesize < 0 Then .Filesize = 0
                        If .Filesize = 94165742 Then Continue For ' SPAM fix
                    End If

                    .Poster = Segs(2).Substring(0, LTtPos).Trim

                    If Not Byte.TryParse(HeaderInfo(0).Substring(0, 1), .Category) Then Continue For

                    If .Category < 1 Then Continue For

                    .SubCat = 100
                    TempCat = vbNullString
                    Catz = HeaderInfo(0).Substring(2).ToLower

                    If NativeSpot Then

                        If Catz.Length < 3 Then Continue For
                        If Catz.Length Mod 3 <> 0 Then Continue For

                        zzT = SplitBySizEx(Catz, 3)

                        For Each rt As String In zzT
                            If rt.Length > 0 Then

                                iFirstChar = Asc(rt)

                                If Not (iFirstChar > 96 And iFirstChar < 123) Then Continue For
                                If Not Byte.TryParse(rt.Substring(1), sSecondByte) Then Continue For

                                TempCat += Chr(iFirstChar) & CStr(sSecondByte) & "|"
                                If iFirstChar = 97 Then .SubCat = sSecondByte

                            End If
                        Next

                    Else

                        CatzCol.Clear()
                        Dim cCatz() As Char = Catz.ToCharArray

                        For i = 0 To Catz.Length - 1
                            If Val(cCatz(i)) = 0 Then
                                If cCatz(i) <> "0" Then
                                    If Len(TempCat) > 0 Then
                                        CatzCol.Add(TempCat)
                                        TempCat = vbNullString
                                    End If
                                End If
                            End If
                            TempCat += Catz.Chars(i)
                        Next

                        CatzCol.Add(TempCat)
                        TempCat = vbNullString

                        For Each rt As String In CatzCol
                            TempCat += rt.Substring(0, 1) & CByte(rt.Substring(1)) & "|"
                        Next

                        .SubCat = CByte(CatzCol(0).Substring(1))

                    End If

                    If .SubCat > 99 Then Continue For
                    If TempCat.Length < 3 Then Continue For

                    .SubCats = TempCat

                    If .Category = 1 Then
                        If IsTV(.SubCats) Then
                            .Category = 6
                        Else
                            If IsEro(.SubCats) Then
                                .Category = 9
                            Else
                                If IsEbook(.SubCats) Then
                                    .Category = 5
                                End If
                            End If
                        End If
                    End If

                    If Segs(1).Contains("=?") Then
                        If Segs(1).Contains("?=") Then
                            Segs(1) = Parse(Replace$(Trim$(Segs(1)), "?= =?", "?==?")).Replace(vbCr, "").Replace(vbLf, "")
                        End If
                    End If

                    If Segs(1).Contains("|") Then
                        zzT = Split(Segs(1), "|")
                        .Title = zzT(0).Trim
                        .Tag = zzT(UBound(zzT)).Trim
                    Else
                        .Title = Segs(1).Trim
                        .Tag = vbNullString
                    End If

                    If .Title.Contains(Chr(194)) Or .Title.Contains(Chr(195)) Then
                        If NativeSpot Then
                            .Title = .Title.Replace(Chr(194), "?").Replace(Chr(195), "?")
                        Else
                            '' Unicode error (alleen bij oude spots)
                            .Title = Parse(UTFEncode(.Title)).Trim
                        End If
                    End If

                    If .Title.Length = 0 Then Continue For
                    If .Poster.Length = 0 Then Continue For

                    sCheck = ""
                    sModulus = Segs(2).Substring(LTtPos + 1, AtPos - LTtPos - 1)

                    If sModulus.Length > 50 Then

                        dPos = sModulus.IndexOf(".")

                        If dPos = -1 Then
                            sModulus = FixPadding(UnSpecialString(sModulus))
                        Else
                            sCheck = FixPadding(UnSpecialString(sModulus.Substring(dPos + 1)))
                            sModulus = FixPadding(UnSpecialString(sModulus.Substring(0, dPos)))
                        End If

                    End If

                    If xSettings.BlackList.Count > 0 Then
                        If xSettings.BlackList.Contains(sModulus) Then Continue For
                    End If

                    If NativeSpot Then

                        If (xSettings.CheckSignatures Or (.KeyID = 2)) Then

                            sHeaderSign = HeaderInfo(UBound(HeaderInfo))

                            If sHeaderSign Is Nothing Then Continue For
                            If sHeaderSign.Length = 0 Then Continue For

                            If (.KeyID <> 7) Then

                                If RSA(.KeyID) Is Nothing Then Continue For

                                If sCheck.Length > 0 Then

                                    ShABytes = ShA.ComputeHash(MakeLatin(Segs(4)))

                                    If ShABytes(0) <> 0 Then Continue For
                                    If ShABytes(1) <> 0 Then Continue For

                                End If

                                If Not RSA(.KeyID).VerifyHash(ShA.ComputeHash(MakeLatin(.Title & sHeader.Substring(0, sHeader.Length - sHeaderSign.Length - 1) & .Poster)), Nothing, Convert.FromBase64String(FixPadding(UnSpecialString(sHeaderSign)))) Then

                                    Continue For

                                End If

                                If sCheck.Length > 0 Then

                                    cRSA = MakeRSA(sModulus)

                                    If cRSA Is Nothing Then Continue For

                                    If Not cRSA.VerifyHash(ShABytes, Nothing, Convert.FromBase64String(sCheck)) Then Continue For

                                    .Modulus = sModulus

                                End If

                            Else

                                If Not RSA(7) Is Nothing Then Continue For

                                ShABytes = ShA.ComputeHash(MakeLatin(Segs(4)))

                                If ShABytes(0) <> 0 Then Continue For
                                If ShABytes(1) <> 0 Then Continue For

                                If sModulus.Length < 50 Then

                                    .KeyID = 9

                                Else

                                    cRSA = MakeRSA(sModulus)

                                    If cRSA Is Nothing Then Continue For

                                    If sCheck.Length > 0 Then

                                        If Not cRSA.VerifyHash(ShABytes, Nothing, Convert.FromBase64String(sCheck)) Then Continue For

                                    Else

                                        If Not cRSA.VerifyHash(ShA.ComputeHash(MakeLatin(.Title & sHeader.Substring(0, sHeader.Length - sHeaderSign.Length - 1) & .Poster)), Nothing, Convert.FromBase64String(FixPadding(UnSpecialString(sHeaderSign)))) Then Continue For

                                    End If

                                    .Modulus = sModulus

                                End If

                            End If

                        End If

                    End If

                    xOutputData.Add(xOD)

                End With

            Catch ex As Exception

                Continue For

            End Try
        Next

        xCnt += xOutputData.Count
        xOutputData.Reverse()

        Return True

    End Function

End Class

