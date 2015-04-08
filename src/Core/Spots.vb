Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography

Imports Phuse
Imports System.Data.Common

Namespace Spotlib

    Public Class Spots

        Friend Sub New()

            MyBase.New()

        End Sub

        Public Shared Function CreateSpot(ByVal tPhuse As Engine, ByVal Newsgroup As String, ByVal sTitle As String, ByVal sDesc As String, ByVal bCat As Byte, ByVal sSubCats As String, ByVal sUrl As String, ByVal sLanguage As String, ByVal SizeX As Long, ByVal SizeY As Long, ByVal sNZB As String, ByVal sFrom As String, ByVal sTag As String, ByVal sNZBGroup As String, ByVal cRSA As RSACryptoServiceProvider, ByVal sHashMsgID As String, ByVal bImage() As Byte, ByVal bAvatar() As Byte, ByVal SignLocal As Boolean, ByRef Settings As NNTPSettings, ByRef zErr As String) As Boolean

            sUrl = sUrl.Trim
            sNZB = sNZB.Trim
            sDesc = sDesc.Trim
            sNZBGroup = sNZBGroup.Trim

            sTag = Utils.StripNonAlphaNumericCharacters(sTag.Trim)
            sSubCats = Utils.StripNonAlphaNumericCharacters(sSubCats.Trim.ToLower)
            sTitle = sTitle.Replace("!!", "").Replace("**", "").Replace("_", " ").Trim

            If bCat < 1 Then zErr = "Vul een categorie in!" : Return False

            If Len(sTitle) < 2 Then zErr = "Vul een titel in!" : Return False
            If Len(sDesc) = 0 Then zErr = "Vul een beschrijving in!" : Return False
            If Len(sDesc) < 30 Then zErr = "Beschrijving is te kort!" : Return False

            If Len(sSubCats) < 3 Then zErr = "Selecteer een subcategorie!!" : Return False
            If sSubCats.Length Mod 3 <> 0 Then zErr = "Ongeldige subcategorien!!" : Return False
            If Not sSubCats.StartsWith("a") Then zErr = "Ongeldige subcategorien!!" : Return False

            If Len(sFrom) < 3 Then zErr = "Afzender niet ingevuld!" : Return False
            If Not Utils.CheckFrom(sFrom) Then zErr = "Afzender ongeldig!" : Return False
            If Len(sFrom) > 22 Then zErr = "Afzender is te lang!" : Return False
            If sFrom <> Utils.StripNonAlphaNumericCharacters(Trim(sFrom)) Then zErr = "Afzender bevat ongeldige tekens!" : Return False

            If sTitle.ToLower.Contains("www.") Then zErr = "URL's mogen niet in de titel voorkomen!" : Return False
            If sTitle.ToLower.Contains("http:/") Then zErr = "URL's mogen niet in de titel voorkomen!" : Return False
            If sTitle.ToLower.Contains(sFrom.ToLower.Trim) Then zErr = "Je eigen naam mag niet in de titel voorkomen!" : Return False
            If (Len(sTitle) > 6) And (sTitle.ToUpper = sTitle) Then zErr = "Titel mag niet alleen uit hoofdletters bestaan!" : Return False

            If Len(sNZB) = 0 Then zErr = ("Voeg eerst een NZB bestand toe!") : Return False
            If Len(sNZB) = 0 Then zErr = ("NZB niet opgegeven!") : Return False
            If Len(sNZBGroup) = 0 Then zErr = ("Specificeer de groep waar de NZB geplaats moet worden!") : Return False

            If Len(sHashMsgID) < 7 Then zErr = ("Message-ID ongeldig!") : Return False
            If Not Utils.CheckHash(sHashMsgID) Then zErr = ("Message-ID ongeldig!") : Return False
            If Not sHashMsgID.StartsWith("<") Then zErr = ("Message-ID ongeldig!") : Return False

            If bImage Is Nothing Then zErr = ("Geen plaatje toegevoegd!") : Return False
            If UBound(bImage) < 10 Then zErr = ("Geen plaatje toegevoegd!") : Return False
            If UBound(bImage) > 1048576 Then zErr = ("Plaatje is groter dan 1 MB!") : Return False

            If Not bAvatar Is Nothing Then
                If UBound(bAvatar) > 4000 Then zErr = ("Avatar is te groot!") : Return False
            End If

            If cRSA Is Nothing Then zErr = ("RSA sleutel ongeldig!") : Return False
            If cRSA.PublicOnly Then zErr = ("RSA sleutel ongeldig!") : Return False

            Dim NZBXML As String

            If sNZB.ToLower.Contains("<nzb") And sNZB.ToLower.Contains("<?xml") Then
                NZBXML = sNZB
            Else
                If Not Utils.FileExists(sNZB) Then zErr = ("Kan de NZB niet vinden? Controleer of het bestand bestaat!") : Return False
                NZBXML = Utils.GetFileContents(sNZB)
            End If

            Dim NzbSize As Long = 0

            If Not Utils.IsNZB(NZBXML, NzbSize) Then zErr = ("Ongeldig NZB bestand!") : Return False

            Dim GoodDesc As String = Microsoft.VisualBasic.Strings.Left(sDesc, 9000).Trim
            Dim zDesc As String = GoodDesc.Replace("[b]", "").Replace("[/b]", "").Replace("[i]", "").Replace("[/i]", "").Replace("[u]", "").Replace("[/u]", "")
            GoodDesc = GoodDesc.Replace(vbCrLf, "[br]").Replace(vbTab, "[tab]")

            Dim GoodPoster As String = Microsoft.VisualBasic.Strings.Left(Utils.StripNonAlphaNumericCharacters(sFrom), 50).Trim.Replace(" ", "")
            Dim GoodTag As String = Microsoft.VisualBasic.Strings.Left(Utils.StripNonAlphaNumericCharacters(sTag), 100).Trim.Replace("|", "").Replace(";", "").Replace(" ", "")

            Dim GoodURL As String = Utils.AddHttp(Microsoft.VisualBasic.Strings.Left(sUrl, 450).Trim.Replace(vbTab, "").Replace(vbCr, "").Replace(vbLf, ""))
            Dim GoodTitle As String = Microsoft.VisualBasic.Strings.Left(sTitle, 450).Trim.Replace(vbTab, "").Replace(vbCr, "").Replace(vbLf, "").Replace("|", "")

            Do While GoodTitle.Contains("  ")
                GoodTitle = GoodTitle.Replace("  ", " ")
            Loop

            Dim NzbMsgId As String = vbNullString
            Dim sMod As String = Utils.SpecialString(Convert.ToBase64String(cRSA.ExportParameters(False).Modulus))

            GoodTitle = GoodTitle.Substring(0, 1).ToUpper & GoodTitle.Substring(1) ' Capitalize first char

            If Not Utils.PostData(tPhuse, Utils.SplitLinesGZIP(Utils.ZipStr(Utils.MakeLatin(NZBXML))), Guid.NewGuid.ToString.Replace("-", ""), GoodPoster & " <" & sMod & "@" & Utils.MsgDomain & ">", sNZBGroup, "", NzbMsgId, "", zErr) Then
                zErr = ("Kan NZB niet posten?" & vbCrLf & vbCrLf & "Details: " & zErr) : Return False
            End If

            Dim xOut(0) As String

            zErr = ""
            Dim ImgMsgId As String = vbNullString
            Dim ImageList As List(Of String) = Utils.SplitLinesGZIP(Utils.GetLatin(bImage))

            If ImageList.Count > 1 Then
                zErr = ("Plaatje is te groot!") : Return False
            End If

            If Not Utils.PostData(tPhuse, ImageList, Guid.NewGuid.ToString.Replace("-", ""), GoodPoster & " <" & sMod & "@" & Utils.MsgDomain & ">", sNZBGroup, "", ImgMsgId, "", zErr) Then
                zErr = ("Kan plaatje niet posten?" & vbCrLf & vbCrLf & "Details: " & zErr) : Return False
            End If

            zErr = ""

            If Not Spots.CreateSpot2(GoodPoster, GoodTitle, NzbMsgId, NzbSize, GoodURL, ImgMsgId, GoodDesc, bCat, sSubCats, GoodTag, xOut, "spot.sign.bot.nu", SizeX, SizeY, zErr) Then
                zErr = ("Kan spot niet toevoegen?" & vbCrLf & vbCrLf & "Details: " & zErr)
                Return False
            End If

            Dim zXML As String = xOut(1)
            Dim zXMLSign As String = Utils.SpecialString(xOut(2))
            Dim zHeader As String = Utils.SpecialString(xOut(0))

            zErr = vbNullString

            Dim zExtra As String = Utils.SplitLinesXML(zXML, "X-XML:", 911) & "X-XML-Signature: " & zXMLSign & vbCrLf
            Dim zSign As String = Utils.CreateUserSignature(Utils.MakeMsg(sHashMsgID), cRSA)

            zExtra += "X-User-Key: " & cRSA.ToXmlString(False).Replace(vbTab, "").Replace(vbCrLf, "") & vbCrLf
            zExtra += "X-User-Signature: " & zSign & vbCrLf

            If Not bAvatar Is Nothing Then
                zExtra += Utils.SplitLinesXML(Convert.ToBase64String(bAvatar).Replace("=", ""), "X-User-Avatar:", 911)
            End If

            Dim zFrom As String = GoodPoster & " <" & sMod & "." & zSign & "@" & zHeader & ">"
            Dim zTitle As String = GoodTitle & Utils.sIIF(GoodTag.Length > 0, " | " & GoodTag, vbNullString)

            Dim iPos As Integer

            If Not Utils.IsAscii(zExtra, iPos) Then
                zErr = ("Spot XML is niet ASCII: " & Mid(zExtra, iPos, 10) & "..")
                Return False
            End If

            Dim LM As New Worker
            Dim xRetSpot As SpotEx

            xRetSpot = LM.ParseSpot(zTitle, zFrom, sHashMsgID, Settings, zErr)

            If xRetSpot Is Nothing Then
                zErr = "Kan spot niet parsen: " & zErr
                Return False
            End If

            If Not Utils.PostData(tPhuse, Utils.SplitLines(sDesc, True, 911), zTitle, zFrom, Newsgroup, zExtra, "", sHashMsgID, zErr) Then
                zErr = ("Kan spot niet posten?" & vbCrLf & vbCrLf & "Details: " & zErr) : Return False
            End If

            Return True

        End Function

        Public Shared Function CreateSpot2(ByVal sPoster As String, ByVal sTitle As String, ByVal sNZB As String, ByVal lFileSize As Long, ByVal sLink As String, ByVal sImageID As String, ByVal sDesc As String, ByVal HCat As Byte, ByVal sCat As String, ByVal sTag As String, ByRef xOut() As String, ByRef SignServer As String, ByVal lImgX As Long, ByVal lImgY As Long, ByRef zErr As String) As Boolean

            Dim zPost As String = ""

            zPost = "cat=" & Str(HCat).Trim & "&"
            zPost += "sub=" & Utils.StripNonAlphaNumericCharacters(sCat).ToLower & "&"
            zPost += "title=" & Utils.MakeP(sTitle) & "&"
            zPost += "from=" & Utils.StripNonAlphaNumericCharacters(sPoster) & "&"
            zPost += "desc=" & Utils.MakeP(sDesc) & "&"
            zPost += "tag=" & Utils.StripNonAlphaNumericCharacters(sTag) & "&"
            zPost += "msgid=" & Utils.MakeP(sNZB) & "&"
            If Len(sLink) > 0 Then zPost += "link=" & Utils.MakeP(sLink) & "&"
            zPost += "img=" & Utils.MakeP(sImageID) & "&"
            zPost += "imgx=" & Str(lImgX).Trim & "&"
            zPost += "imgy=" & Str(lImgY).Trim & "&"
            zPost += "size=" & Str(lFileSize).Trim & "&"

            Dim sResult As String = Utils.GetHash(zPost, SignServer, "signxml", zErr)

            If sResult Is Nothing Then Return False
            If sResult.Length < 20 Then zErr = sResult : Return False
            If sResult.Substring(sResult.Length - 10).ToLower <> "</spotnet>" Then zErr = sResult : Return False

            Try

                Dim zHeader As String = sResult.Substring(0, sResult.IndexOf(" "))
                Dim zXMLSign As String = sResult.Substring(zHeader.Length + 1)

                zXMLSign = zXMLSign.Substring(0, zXMLSign.IndexOf(" "))

                Dim zXML As String = sResult.Substring(zHeader.Length + zXMLSign.Length + 2)

                ReDim xOut(2)

                xOut(0) = zHeader
                xOut(1) = zXML
                xOut(2) = zXMLSign

                Return True

            Catch ex As Exception

                zErr = "CreateSpotSigned:: " & ex.Message
                Return False

            End Try

        End Function

        Public Shared Function CreateComment(ByVal tPhuse As Phuse.Engine, ByVal cFrom As String, ByVal cDesc As String, ByVal cGroup As String, ByVal cOrgMessageID As String, ByVal cOrgTitle As String, ByVal bAvatar() As Byte, ByVal cRSA As RSACryptoServiceProvider, ByVal HashMessageID As String, ByRef zErr As String) As Boolean

            cDesc = cDesc.Trim
            cFrom = cFrom.Trim

            If Len(cDesc) = 0 Then zErr = "Vul een reactie in!" : Return False
            If Len(cDesc) < 3 Then zErr = "Reactie is te kort!" : Return False
            If Len(cDesc) > 900 Then zErr = "Reactie is te lang!" : Return False
            If Len(cFrom) < 3 Then zErr = ("Afzender niet ingevuld!") : Return False
            If Len(cFrom) > 22 Then zErr = ("Afzender is te lang!") : Return False
            If Not Utils.CheckFrom(cFrom) Then zErr = ("Afzender ongeldig!") : Return False
            If cFrom <> Utils.StripNonAlphaNumericCharacters(cFrom) Then zErr = "Afzender bevat ongeldige tekens!" : Return False

            If Not bAvatar Is Nothing Then
                If UBound(bAvatar) > 4000 Then zErr = ("Avatar is te groot!") : Return False
            End If

            If Len(HashMessageID) < 7 Then zErr = ("Message-ID ongeldig!") : Return False
            If Not Utils.CheckHash(HashMessageID) Then zErr = ("Message-ID ongeldig!") : Return False
            If Not HashMessageID.StartsWith("<") Then zErr = ("Message-ID ongeldig!") : Return False

            If cRSA Is Nothing Then zErr = ("RSA sleutel ongeldig!") : Return False
            If cRSA.PublicOnly Then zErr = ("RSA sleutel ongeldig!") : Return False

            cOrgTitle = cOrgTitle.Replace(vbCrLf, "")
            cOrgMessageID = Utils.MakeMsg(cOrgMessageID, False)

            Dim CoodDesc As String = Microsoft.VisualBasic.Strings.Left(cDesc, 999).Trim
            Dim CoodPoster As String = Microsoft.VisualBasic.Strings.Left(Utils.StripNonAlphaNumericCharacters(cFrom), 22).Trim.Replace(" ", "")

            Do While CoodDesc.Contains(vbCrLf & vbCrLf & vbCrLf)
                CoodDesc = CoodDesc.Replace(vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
            Loop

            Do While CoodDesc.StartsWith(vbCrLf)
                CoodDesc = CoodDesc.Substring(2)
            Loop

            Do While CoodDesc.EndsWith(vbCrLf)
                CoodDesc = CoodDesc.Substring(0, CoodDesc.Length - 2)
            Loop

            CoodDesc += vbCrLf

            Dim zSub As String = "Re: " & cOrgTitle
            Dim kDesc As List(Of String) = Utils.SplitLines(CoodDesc, True, 911)

            Dim Sig As String = Utils.CreateUserSignature(Utils.MakeMsg(HashMessageID), cRSA)

            Dim sExtra As String = "References: <" & cOrgMessageID & ">" & vbCrLf & "X-User-Signature: " & Sig & vbCrLf
            sExtra += "X-User-Key: " & cRSA.ToXmlString(False).Replace(vbTab, "").Replace(vbCrLf, "") & vbCrLf

            If Not bAvatar Is Nothing Then
                sExtra += Utils.SplitLinesXML(Convert.ToBase64String(bAvatar).Replace("=", ""), "X-User-Avatar:", 911)
            End If

            Dim xOut As String = ""
            Dim sMod As String = Utils.SpecialString(Convert.ToBase64String(cRSA.ExportParameters(False).Modulus))

            If Not Utils.PostData(tPhuse, kDesc, zSub, CoodPoster & " <" & sMod & "." & Sig & "@" & Utils.MsgDomain & ">", cGroup, sExtra, xOut, HashMessageID, zErr) Then
                zErr = ("Kan reactie niet plaatsen?" & vbCrLf & vbCrLf & "Details: " & zErr) : Return False
            End If

            zErr = "Ongeldige MsgID?: " & xOut
            Return (xOut.Length > 0)

        End Function

        Public Shared Function CreatReport(ByVal tPhuse As Phuse.Engine, ByVal cFrom As String, ByVal cDesc As String, ByVal cGroup As String, ByVal cOrgMessageID As String, ByVal cOrgTitle As String, ByRef zErr As String) As Boolean

            cDesc = cDesc.Trim
            cFrom = cFrom.Trim

            If Len(cDesc) = 0 Then zErr = "Vul een beschrijving in!" : Return False
            If Len(cDesc) < 3 Then zErr = "beschrijving is te kort!" : Return False
            If Len(cDesc) > 900 Then zErr = "beschrijving is te lang!" : Return False

            If Len(cFrom) < 3 Then cFrom = "Afzender"
            If Len(cFrom) > 22 Then zErr = ("Afzender te lang!") : Return False

            cOrgMessageID = Utils.MakeMsg(cOrgMessageID, True)
            cOrgTitle = cOrgTitle.Replace(vbCrLf, "")

            Dim CoodDesc As String = Microsoft.VisualBasic.Strings.Left(cDesc, 999).Trim
            Dim CoodPoster As String = Microsoft.VisualBasic.Strings.Left(Utils.StripNonAlphaNumericCharacters(cFrom), 22).Trim.Replace(" ", "")

            Do While CoodDesc.Contains(vbCrLf & vbCrLf & vbCrLf)
                CoodDesc = CoodDesc.Replace(vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
            Loop

            Do While CoodDesc.StartsWith(vbCrLf)
                CoodDesc = CoodDesc.Substring(2)
            Loop

            Do While CoodDesc.EndsWith(vbCrLf)
                CoodDesc = CoodDesc.Substring(0, CoodDesc.Length - 2)
            Loop

            CoodDesc += vbCrLf

            Dim zErr2 As String = vbNullString

            Dim zSub As String = "REPORT " & cOrgMessageID & " - " & cOrgTitle
            Dim kDesc As List(Of String) = Utils.SplitLines(CoodDesc, True, 911)

            Dim zExtra As String = "References: " & cOrgMessageID & vbCrLf

            Return Utils.PostData(tPhuse, kDesc, zSub, CoodPoster & " <" & CoodPoster.ToLower & "@" & Utils.MsgDomain & ">", cGroup, zExtra, "", "", zErr)

        End Function

        Public Shared Function GetSpot(ByVal tPhuse As Engine, ByVal Newsgroup As String, ByVal ArticleID As Long, ByVal MessageID As String, ByRef xOut As String, ByRef SpotOut As SpotEx, ByVal xParam As NNTPSettings, ByRef sError As String) As Boolean

            Dim cNNTP As cNNTP
            Dim lRet As Integer = -1
            Dim sOut As String = ""
            Dim sRet As Boolean = False

            cNNTP = New cNNTP(tPhuse)

            If ArticleID > 0 Then

                sRet = cNNTP.GetHeader(Newsgroup, CStr(ArticleID), sOut, lRet, sError)

                If Not sRet Then
                    If (lRet <> 423) Then Return False
                End If

            End If

            If Not sRet Then

                If Len(MessageID) > 0 Then
                    sRet = cNNTP.GetHeader(Newsgroup, MessageID, sOut, lRet, sError)
                End If

            End If

            If Not sRet Then

                If (lRet = 430 Or lRet = 423) Then
                    sError = "Spot niet gevonden, waarschijnlijk te oud."
                End If

                Return False

            End If

            Dim XL() As String
            Dim OrgXML As String = ""
            Dim tMsgID As String = ""
            Dim zFrom As String = ""
            Dim zSubject As String = ""

            Dim zOrgXmlSig As String = ""
            Dim zUs As New UserInfo

            Dim RR As Integer

            XL = Split(sOut, vbCrLf)

            For iXL = 1 To UBound(XL) - 1

                If UCase(XL(iXL)).StartsWith("SUBJECT: ") Then
                    zSubject = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                    For RR = iXL + 1 To UBound(XL) - 1
                        If XL(RR).IndexOf(":") = -1 Or XL(RR).StartsWith(" ") Or XL(RR).StartsWith(Chr(9)) Then
                            zSubject += XL(RR) ' Multine subjects
                        Else
                            Exit For
                        End If
                    Next
                End If

                If UCase(XL(iXL)).StartsWith("FROM: ") Then
                    zFrom = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                    For RR = iXL + 1 To UBound(XL) - 1
                        If XL(RR).IndexOf(":") = -1 Or XL(RR).StartsWith(" ") Or XL(RR).StartsWith(Chr(9)) Then
                            zFrom += XL(RR) ' Multine from
                        Else
                            Exit For
                        End If
                    Next
                End If

                If UCase(XL(iXL)).StartsWith("MESSAGE-ID: ") Then
                    tMsgID = Utils.MakeMsg(Mid(XL(iXL), XL(iXL).IndexOf(":") + 3))
                End If

                If UCase(XL(iXL)).StartsWith("X-USER-AVATAR: ") Then
                    zUs.Avatar += Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                End If

                If UCase(XL(iXL)).StartsWith("X-XML: ") Then

                    OrgXML += Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)

                End If

                If UCase(XL(iXL)).StartsWith("X-USER-KEY: ") Then

                    With zUs

                        .Modulus = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)

                        If .Modulus.ToLower.Contains("<modulus>") Then

                            .Modulus = .Modulus.Substring(.Modulus.ToLower.IndexOf("<modulus>") + 9)
                            If .Modulus.Contains("<") Then .Modulus = .Modulus.Substring(0, .Modulus.IndexOf("<"))

                        Else

                            .Modulus = Utils.FixPadding(Utils.UnSpecialString(.Modulus))

                        End If

                    End With

                End If

                If UCase(XL(iXL)).StartsWith("ORGANIZATION: ") Then
                    zUs.Organisation = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                    zUs.Organisation = zUs.Organisation.Substring(0, 1).ToUpper & zUs.Organisation.Substring(1) ' Capitalize first char
                End If

                If UCase(XL(iXL)).StartsWith("X-TRACE: ") Then
                    zUs.Trace += vbCrLf & Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                End If

                If UCase(XL(iXL)).StartsWith("NNTP-POSTING-HOST: ") Then
                    Dim sHost As String = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                    If sHost.IndexOf(" (") > 0 Then
                        sHost = sHost.Replace(")", "")
                        sHost = sHost.Substring(0, sHost.IndexOf(" ("))
                    End If
                    zUs.Trace += vbCrLf & sHost
                End If

                If UCase(XL(iXL)).StartsWith("X-ORIGINATING-IP: ") Then
                    zUs.Trace += vbCrLf & Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                End If

                If UCase(XL(iXL)).StartsWith("X-USER-SIGNATURE: ") Then
                    zUs.Signature = Utils.UnSpecialString(Mid(XL(iXL), XL(iXL).IndexOf(":") + 3))
                End If

                If UCase(XL(iXL)).StartsWith("X-XML-SIGNATURE: ") Then
                    zOrgXmlSig = Mid(XL(iXL), XL(iXL).IndexOf(":") + 3)
                End If

            Next

            zUs.Trace = zUs.Trace.Replace(vbCrLf & zUs.Organisation, "")
            If zUs.Trace = vbCrLf Then zUs.Trace = ""

            If Len(zUs.Avatar) > 0 Then
                zUs.Avatar = Utils.FixPadding(Utils.UnSpecialString(zUs.Avatar))
            End If

            Dim LM As New Worker
            Dim xRetSpot As SpotEx

            If xParam.BlackList.Contains(zUs.Modulus) Then
                sError = "Afzender staat op de zwarte lijst."
                Return False
            End If

            xRetSpot = LM.ParseSpot(zSubject, zFrom, tMsgID, xParam, sError)

            If xRetSpot Is Nothing Then
                sError = "Kan spot niet parsen: " & sError
                Return False
            End If

            xRetSpot.User = zUs

            If Len(xRetSpot.Modulus) > 0 Then
                If xRetSpot.User.Modulus <> xRetSpot.Modulus Then
                    sError = "Handtekening is niet correct."
                    Return False
                End If
            End If

            If xRetSpot.KeyID = 1 Then
                OrgXML = OrgXML.Replace("<Signature><", "<Signature xmlns=" & Chr(34) & "http://www.w3.org/2000/09/xmldsig#" & Chr(34) & "><")
                OrgXML = OrgXML.Replace("<SignedInfo><", "<SignedInfo><CanonicalizationMethod Algorithm=" & Chr(34) & "http://www.w3.org/TR/2001/REC-xml-c14n-20010315" & Chr(34) & " /><SignatureMethod Algorithm=" & Chr(34) & "http://www.w3.org/2000/09/xmldsig#dsa-sha1" & Chr(34) & " /><")
                OrgXML = OrgXML.Replace("<Reference URI=" & Chr(34) & Chr(34) & ">", "<Reference URI=" & Chr(34) & Chr(34) & "><Transforms><Transform Algorithm=" & Chr(34) & "http://www.w3.org/2000/09/xmldsig#enveloped-signature" & Chr(34) & " /></Transforms><DigestMethod Algorithm=" & Chr(34) & "http://www.w3.org/2000/09/xmldsig#sha1" & Chr(34) & " />")
            End If

            Dim _doc As New XmlDocument
            Dim KL As New ASCIIEncoding

            Try

                _doc.XmlResolver = Nothing
                _doc.LoadXml(Utils.MakeAscii((OrgXML)))

            Catch ex As Exception
                sError = "Kan spot niet ophalen: XML niet geldig" & vbCrLf & vbCrLf & ex.Message
                Return False
            End Try

            xOut = _doc.OuterXml.Replace("<SpotNet>", "<Spotnet>").Replace("</SpotNet>", "</Spotnet>")

            xRetSpot = LM.ParseSpotXML(xRetSpot, _doc, zOrgXmlSig, xParam.CheckSignatures, sError)

            If xRetSpot Is Nothing Then
                sError = "Kan spot niet openen: " & sError
                Return False
            End If

            SpotOut = xRetSpot
            Return True

        End Function

        Public Shared Function GetNZB(ByVal tPhuse As Engine, ByVal Newsgroup As String, ByVal xMsgID As List(Of String), ByRef sxOut As String, ByRef sError As String) As Boolean

            Dim bArr() As Byte = Nothing

            If Not Utils.GetBinary(tPhuse, Newsgroup, xMsgID, bArr, sError) Then Return False

            sxOut = Utils.UnzipStr(bArr)

            If sxOut Is Nothing Then
                sError = "Kan NZB niet uitpakken!"
                Return False
            End If

            Return True

        End Function

        Public Shared Function GetImage(ByVal tPhuse As Engine, ByVal Newsgroup As String, ByVal xMsgID As List(Of String), ByRef sxOut() As Byte, ByRef sError As String) As Boolean

            Return Utils.GetBinary(tPhuse, Newsgroup, xMsgID, sxOut, sError)

        End Function

        Public Shared Function TestConnection(ByVal tPhuse As Engine, ByVal Newsgroup As String, ByRef sError As String) As Boolean

            Dim cNNTP As cNNTP
            Dim sRet As Boolean

            cNNTP = New cNNTP(tPhuse)

            sRet = cNNTP.SelectGroup(Newsgroup, 0, 0, 0, 0, sError)

            cNNTP = Nothing

            Return sRet

        End Function

        Public Shared Function FindSpots(ByVal tPhuse As Engine, ByVal xParam As NNTPSettings) As Headers

            Dim hw As New Headers

            hw.FindHeaders(tPhuse, xParam)

            Return hw

        End Function

        Public Shared Function GetComments(ByVal tPhuse As Engine, ByVal ArticleIDs As List(Of Long), ByVal xParam As NNTPSettings) As Comments

            Dim hw As New Comments

            hw.GetComments(tPhuse, ArticleIDs, xParam)

            Return hw

        End Function

        Public Shared Function FindComments(ByVal tPhuse As Engine, ByVal xParam As NNTPSettings) As Comments

            Dim hw As New Comments

            hw.FindComments(tPhuse, xParam)

            Return hw

        End Function

        Public Shared Function GetComments(Db As Database, ByVal xMsg As String, ByRef Param As Parameters, ByRef Fetch As FetchCache, ByRef xErr As String) As List(Of Long)

            Dim lFnd As Long
            Dim sErr As String = ""
            Dim sMessageID As String = ""
            Dim DR As DbDataReader = Nothing

            Try

                sMessageID = Utils.MakeMsg(xMsg, False)
                Dim AtPos As Integer = sMessageID.IndexOf("@")
                sMessageID = sMessageID.Substring(0, AtPos)

                DR = Db.ExecuteReader("SELECT docid FROM comments WHERE spot MATCH '" & sMessageID.Replace("'"c, "") & "' ORDER BY docid ASC", sErr)

                If DR Is Nothing Then Throw New Exception("Datareader not available: " & sErr)

                With DR
                    While .Read

                        If IsDBNull(.Item(0)) Then Continue While

                        lFnd = CLng(.Item(0))

                        If lFnd > 0 Then

                            If Not Fetch.CacheHash.Contains(lFnd) Then
                                Fetch.Cache.Add(lFnd)
                                Fetch.CacheHash.Add(lFnd)
                            End If

                        End If

                    End While
                End With

                DR.Close()
                Db.Close()

                Return Fetch.Cache

            Catch ex As Exception

                Try
                    DR.Close()
                Catch
                End Try

                Try
                    Db.Close()
                Catch
                End Try

                xErr = "GetComments: " & ex.Message
                Return Nothing

            End Try

        End Function

        Public Shared Function DoComments(xMsg As String, DatabaseFile As String, Param As Parameters, Fetch As FetchCache) As List(Of Long)

            Dim zErr As String = ""
            Dim Db As Database = New Database

            If Not Db.Connect(DatabaseFile, True) Then
                Throw New Exception("Db.Connect")
            End If

            If Not Db.ExecuteNonQuery("PRAGMA temp_store = MEMORY;", "") = 0 Then Throw New Exception("PRAGMA temp_store")
            If Not Db.ExecuteNonQuery("PRAGMA cache_size = " & CStr(Param.DatabaseCache), "") = 0 Then Throw New Exception("PRAGMA cache_size")

            Return Spots.GetComments(Db, xMsg, Param, Fetch, zErr)

        End Function

    End Class

End Namespace
