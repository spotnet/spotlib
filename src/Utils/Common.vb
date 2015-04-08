Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Threading
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Data.Common

Namespace Spotlib

    Public Class Utils

        Friend Const SearchM As String = "Zoeken: "
        Friend Const CancelMSG As String = "Geannuleerd"
        Friend Const Spotname As String = "Spotnet"
        Friend Const DefaultFilter As String = "cat < 9"

        Friend Shared ReadOnly EPOCH As Date = New Date(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
        Private Declare Function SHGetKnownFolderPath Lib "shell32" (ByRef knownFolder As Guid, ByVal flags As UInteger, ByVal htoken As IntPtr, ByRef path As IntPtr) As Integer

        Private Enum convTo
            B = 0
            KB = 1
            MB = 2
            GB = 3  'Enumerations for file size conversions
            TB = 4
            PB = 5
            EB = 6
            ZI = 7
            YI = 8
        End Enum

        Friend Const MsgDomain As String = "spot.net"

        Friend Shared Function LatinEnc() As Encoding

            Return Encoding.GetEncoding(&H6FAF)

        End Function

        Friend Shared Function GetLatin(ByRef zText() As Byte) As String

            Return LatinEnc.GetString(zText)

        End Function

        Friend Shared Function MakeLatin(ByVal zText As String) As Byte()

            Return LatinEnc.GetBytes(zText)

        End Function
        Friend Shared Function ConvertDate(ByVal sDate As String) As Date

            Try

                Dim dt As Date
                Dim RFC822 As String = ""
                Dim sLeft As String = sDate.Substring(0, 1)

                If Not IsNumeric(sLeft) Then
                    If sDate.IndexOf(",") = 3 Then
                        sDate = sDate.Substring(4)
                    End If
                End If

                sDate = sDate.Replace(" GMT", "").Trim

                If (Len(sDate) > 20) Then
                    RFC822 = "dd MMM yyyy HH:mm:ss zzz"
                    dt = DateTime.ParseExact(sDate, RFC822, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None)
                Else
                    RFC822 = "dd MMM yyyy HH:mm:ss"
                    dt = DateTime.ParseExact(sDate, RFC822, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.AssumeUniversal)
                End If

                Return dt

            Catch ex As Exception

                Return (New Date)

            End Try

        End Function
        Friend Shared Function FixPadding(ByRef sIn As String) As String

            Select Case (sIn.Length Mod 4)
                Case 0
                    Return sIn
                Case 1
                    Return sIn & "==="
                Case 2
                    Return sIn & "=="
                Case 3
                    Return sIn & "="
            End Select

            Return vbNullString

        End Function

        Friend Shared Function IsAscii(ByVal text As String, ByRef lPos As Integer) As Boolean

            Dim iPos As Integer = 0

            If Len(Trim$(text)) = 0 Then Return True
            Dim chars() As Char = text.ToCharArray

            For Each c As Char In chars
                Dim value As Integer = Convert.ToInt32(c)
                If (value > 126) Then
                    lPos = iPos
                    Return False
                End If
                iPos += 1
            Next

            Return True

        End Function

        Friend Shared Function MakeAscii(ByVal text As String) As String

            If Len(Trim$(text)) = 0 Then Return ""

            Dim chars() As Char = text.ToCharArray
            Dim tOut As New StringBuilder

            For Each c As Char In chars
                Dim value As Integer = Convert.ToInt32(c)
                If (value < 127) And (value > 31) Then
                    tOut.Append(c)
                End If
            Next

            Return tOut.ToString

        End Function

        Friend Shared Function SplitBySizEx(ByVal strInput As String, ByVal iSize As Integer) As String()

            Dim iLength As Integer = strInput.Length()
            Dim iWords As Integer = CInt(iLength / iSize)
            If (iLength Mod iSize) <> 0 Then iWords += 1

            Dim j As Integer = 0
            Dim strA(iWords) As String

            For i As Integer = 0 To iLength Step iSize
                strA(j) = Mid(strInput, i + 1, iSize)
                j += 1
            Next

            Return strA

        End Function

        Friend Shared Function UTFEncode(ByVal sInput As String) As String

            Dim i As Integer
            Dim Ascii As Integer
            Dim xEncodedChar(0) As Byte
            Dim xEncodedCnt As Integer
            Dim ReturnString As New StringBuilder
            Dim Chars() As Char = sInput.ToCharArray()

            For i = 0 To Chars.Length - 1
                Ascii = Asc(Chars(i))
                If (Ascii < 32) Or Ascii = 63 Or (Ascii > 126) Then
                    Select Case Ascii
                        Case 9, 10, 13
                            If xEncodedCnt > 0 Then ReturnString.Append("=?UTF-8?B?" & System.Convert.ToBase64String(xEncodedChar) & "?=")
                            xEncodedCnt = 0
                            ReturnString.Append(Chars(i))
                        Case Else
                            ReDim Preserve xEncodedChar(xEncodedCnt)
                            xEncodedChar(xEncodedCnt) = CByte(Ascii)
                            xEncodedCnt += 1
                    End Select
                Else
                    If xEncodedCnt > 0 Then ReturnString.Append("=?UTF-8?B?" & System.Convert.ToBase64String(xEncodedChar) & "?=")
                    xEncodedCnt = 0
                    ReturnString.Append(Chars(i))
                End If
            Next

            If xEncodedCnt > 0 Then ReturnString.Append("=?UTF-8?B?" & System.Convert.ToBase64String(xEncodedChar) & "?=")

            Return ReturnString.ToString

        End Function

        Private Shared Function ParseEncodedWord(ByVal input As String) As String

            Dim enc As Encoding
            Dim sb As New StringBuilder()

            Try

                If Not input.StartsWith("=") Then Return input
                If Not input.EndsWith("?=") Then Return input

                Dim encodingName As String = input.Substring(2, input.IndexOf("?", 2) - 2)

                If encodingName.ToUpper = "UTF8" Then encodingName = "UTF-8"
                enc = Encoding.GetEncoding(encodingName)

                Dim type As Char = input.Chars(encodingName.Length + 3)
                Dim i As Int32 = encodingName.Length + 5

                Select Case Char.ToLowerInvariant(type)
                    Case "q"c
                        Do While i < input.Length
                            Dim currentChar As Char = input.Chars(i)
                            Dim peekAhead(1) As Char
                            Select Case currentChar
                                Case "="c
                                    peekAhead = If(i >= input.Length - 2, Nothing, New Char() {input.Chars(i + 1), input.Chars(i + 2)})
                                    If peekAhead Is Nothing Then Exit Select
                                    Dim decodedChar As String = enc.GetString(New Byte() {Convert.ToByte(New String(peekAhead, 0, 2), 16)})
                                    sb.Append(decodedChar)
                                    i += 3
                                Case "?"c
                                    If input.Chars(i + 1) = "=" Then
                                        i += 2
                                    End If
                                Case Else
                                    sb.Append(currentChar)
                                    i += 1
                            End Select
                        Loop
                    Case "b"c
                        Dim baseString As String = input.Substring(i, input.Length - i - 2)
                        Dim baseDecoded() As Byte = Convert.FromBase64String(baseString)
                        sb.Append(enc.GetString(baseDecoded))
                End Select

                Return sb.ToString()

            Catch
                Return vbNullString
            End Try

        End Function

        Friend Shared Function Parse(ByVal input As String) As String

            Dim sb As New StringBuilder()
            Dim currentWord As New StringBuilder()
            Dim readingWord As Boolean = False

            Dim i As Int32 = 0

            Do While i < input.Length
                Dim currentChar As Char = input.Chars(i)
                Dim peekAhead As Char
                Select Case currentChar
                    Case "="c
                        peekAhead = If(i = input.Length - 1, " "c, input.Chars(i + 1))

                        If peekAhead = "?" Then
                            readingWord = True
                        End If

                    Case "?"c
                        peekAhead = If(i = input.Length - 1, " "c, input.Chars(i + 1))

                        If peekAhead = "=" Then
                            readingWord = False

                            currentWord.Append(currentChar)
                            currentWord.Append(peekAhead)

                            sb.Append(ParseEncodedWord(currentWord.ToString()))
                            currentWord = New StringBuilder()

                            i += 2
                            Continue Do
                        End If
                End Select

                If readingWord Then
                    currentWord.Append(currentChar)
                    i += 1
                Else
                    sb.Append(currentChar)
                    i += 1
                End If
            Loop

            Return sb.ToString()

        End Function

        Friend Shared Function ZipStr(ByRef Inz() As Byte) As String

            Try

                Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream
                Dim sw As System.IO.Compression.DeflateStream = New System.IO.Compression.DeflateStream(ms, System.IO.Compression.CompressionMode.Compress, True)

                sw.Write(Inz, 0, Inz.Length)
                sw.Close()
                ms.Seek(0, System.IO.SeekOrigin.Begin)

                Dim reader As New StreamReader(ms, LatinEnc)
                Return reader.ReadToEnd

            Catch
                Return Nothing
            End Try

        End Function

        Friend Shared Function UnzipStr(ByRef Inz() As Byte) As String

            Try

                Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(Inz)
                Dim sw As System.IO.Compression.DeflateStream = New System.IO.Compression.DeflateStream(ms, System.IO.Compression.CompressionMode.Decompress, True)
                Dim reader As New StreamReader(sw, LatinEnc)

                Return reader.ReadToEnd

            Catch ex As Exception

                Return Nothing

            End Try

        End Function

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

        Private Shared Function UnGZIP(ByRef Inz() As Byte) As String

            Try

                Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(Inz)
                Dim sw As System.IO.Compression.GZipStream = New System.IO.Compression.GZipStream(ms, System.IO.Compression.CompressionMode.Decompress, True)
                Dim reader As New StreamReader(sw, LatinEnc)
                Return reader.ReadToEnd

            Catch ex As Exception

                Return Nothing

            End Try

        End Function

        Friend Shared Function CheckHash(ByVal sMsg As String) As Boolean

            Dim ShA As New SHA1Managed
            Dim ShABytes() As Byte = ShA.ComputeHash(MakeLatin(sMsg))

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

            bytArguments = MakeLatin(zPostData)

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

                Dim reader As New IO.StreamReader(myRequest.GetResponse().GetResponseStream, LatinEnc)

                sReturn = UnGZIP(MakeLatin(reader.ReadToEnd))

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

                Return RSAalg.VerifyHash((New SHA1Managed).ComputeHash(MakeLatin(sOrg)), Nothing, Convert.FromBase64String(FixPadding(sSignature)))

            Catch
            End Try

            Return False

        End Function

        Friend Shared Function SpecialString(ByVal sDataIn As String) As String

            Return sDataIn.Replace("/", "-s").Replace("+", "-p").Replace("=", "")

        End Function

        Friend Shared Function UnSpecialString(ByVal sDataIn As String) As String

            Return sDataIn.Replace("-s", "/").Replace("-p", "+")

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

        Friend Shared Function SplitLines(ByVal sIn As String, ByVal AllowBlankLines As Boolean, ByVal lMax As Integer) As List(Of String)

            Dim Lin As Integer
            Dim Lin2 As Integer
            Dim Linez3 As String()
            Dim Linez2 As String()
            Dim LineCount As Long

            Dim xTheBody As StringBuilder
            Dim xOut As New List(Of String)

            Linez3 = Split(sIn, vbCrLf)
            xTheBody = New StringBuilder

            For Lin2 = 0 To UBound(Linez3)

                Linez2 = SplitBySizEx(Linez3(Lin2), lMax)

                For Lin = 0 To UBound(Linez2)

                    If Not Linez2(Lin) Is Nothing Then

                        If (Len(Linez2(Lin)) > 0) Or AllowBlankLines Then

                            If Linez2(Lin).StartsWith(".") Then
                                xTheBody.AppendLine("." & Linez2(Lin))
                            Else
                                xTheBody.AppendLine(Linez2(Lin))
                            End If

                            LineCount += 1

                        End If

                    End If
                Next
            Next

            If LineCount > 0 Then xOut.Add(xTheBody.ToString)

            Return xOut

        End Function

        Friend Shared Function SplitLinesGZIP(ByVal sIn As String) As List(Of String)

            Dim Lin As Integer
            Dim Linez2 As String()
            Dim LineCount As Long

            Dim xTheBody As StringBuilder
            Dim xOut As New List(Of String)

            xTheBody = New StringBuilder

            If sIn Is Nothing Then Return Nothing

            Linez2 = SplitBySizEx(sIn, 900)

            For Lin = 0 To UBound(Linez2)

                If Not Linez2(Lin) Is Nothing Then

                    If Len(Linez2(Lin)) > 0 Then

                        LineCount += 1
                        If Linez2(Lin).StartsWith(".") Then Linez2(Lin) = "." & Linez2(Lin)
                        xTheBody.AppendLine(Linez2(Lin).Replace("=", "=D").Replace(vbLf, "=C").Replace(vbCr, "=B").Replace(Chr(0), "=A"))

                        If (LineCount = 900) Then

                            xOut.Add(xTheBody.ToString)

                            LineCount = 0
                            xTheBody = New StringBuilder

                        End If

                    End If

                End If
            Next

            If LineCount > 0 Then xOut.Add(xTheBody.ToString)

            Return xOut

        End Function

        Friend Shared Function SplitLinesXML(ByVal sIn As String, ByVal sPrefix As String, ByVal lMax As Integer) As String

            Dim Lin As Integer
            Dim Lin2 As Integer
            Dim Linez3 As String()
            Dim Linez2 As String()

            Dim xTheBody As StringBuilder

            Linez3 = Split(sIn, vbCrLf)
            xTheBody = New StringBuilder

            For Lin2 = 0 To UBound(Linez3)

                Linez2 = SplitBySizEx(Linez3(Lin2), lMax)

                For Lin = 0 To UBound(Linez2)

                    If Not Linez2(Lin) Is Nothing Then

                        If Len(Linez2(Lin)) > 0 Then

                            xTheBody.AppendLine(sPrefix & " " & Linez2(Lin))

                        End If

                    End If
                Next
            Next

            Return xTheBody.ToString

        End Function

        Friend Shared Function GetLocation(ByVal URL As String) As String

            Try

                Dim myWebRequest As Net.WebRequest = Net.WebRequest.Create(URL)

                myWebRequest.Proxy = Nothing

                Dim myWebResponse As Net.WebResponse = myWebRequest.GetResponse()
                Return myWebResponse.ResponseUri.Host

            Catch ex As Exception
                Return vbNullString
            End Try

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

                If Not IsAscii(zPostData, iPos) Then
                    zError = "Postdata niet ASCII: " & Mid(zPostData, iPos, 10) & ".."
                    Return vbNullString
                End If

                bytArguments = MakeLatin(zPostData)

                Dim sDot As String = "."

                If Len(SignServer) = 0 Then
                    zError = "Geen server opgegeven."
                    Return vbNullString
                End If

                sLoc = GetLocation("http://" & SignServer.Replace("http:", "").Replace("/", "") & "/")

                If Len(sLoc) = 0 Then

                    zError = "Kan de server niet bereiken. Probeer het later nog eens."
                    Return vbNullString

                End If

                Try

                    sReturn = GetLatin(oWeb.UploadData("http://" & sLoc & "/" & SignCmd, "POST", bytArguments))

                Catch ex As Exception

                    zError = "Kan de server niet bereiken (" & ex.Message & "). Probeer het later nog eens."
                    Return vbNullString

                End Try

                oWeb.Dispose()
                oWeb = Nothing

                If Not IsAscii(sReturn, iPos) Then
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

            Return SpecialString(Convert.ToBase64String(cRSA.SignHash((New SHA1Managed).ComputeHash(MakeLatin(sDataIn)), Nothing)))

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

                Dim sHeader As String = "From: " & zFrom & vbCrLf & "Subject: " & zSub & CStr(sIIF((zInput.Count > 1), " [" & XC & "/" & zInput.Count & "]", "")) & vbCrLf & "Newsgroups: " & zGroup & vbCrLf & "Message-ID: " & TempId & vbCrLf & zExtra & "Content-Type: text/plain; charset=ISO-8859-1" & vbCrLf & "Content-Transfer-Encoding: 8bit"

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

            Return Convert.ToBase64String(MakeLatin(sIn)).Replace("=", "%3d").Replace("+", "%2b").Replace("&", "%26").Replace$("/", "%2f").Trim

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

            sxOut = MakeLatin(Retr.ToString.Replace("=C", vbLf).Replace("=B", vbCr).Replace("=A", Chr(0)).Replace("=D", "="))

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

        Friend Shared Function GetCode(ByVal sResponse As String) As Integer

            Try
                Return CInt(Val(Left(sResponse, 3)))
            Catch ex As Exception
                Return 0
            End Try

        End Function

        Public Shared Function ConvertToTimestamp(ByVal value As DateTime) As Integer

            Dim span As TimeSpan = (value - New DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime)
            Return CInt(span.TotalSeconds)

        End Function

        Public Shared Function GetFileSize(ByVal MyFilePath As String) As Long

            If Not FileExists(MyFilePath) Then Return 0

            Dim MyFile As New FileInfo(MyFilePath)
            Dim FileSize As Long = MyFile.Length
            Return FileSize

        End Function

        Public Shared Sub Wait(ByVal ms As Integer)

            Using wh As New ManualResetEvent(False)
                wh.WaitOne(ms)
            End Using

        End Sub

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

        Private Shared Function AddDirSep(ByVal strPathName As String) As String

            If Microsoft.VisualBasic.Right(Trim(strPathName), 1) <> "\" Then Return Trim(strPathName) & "\"
            Return Trim(strPathName)

        End Function

        Public Shared Function DirectoryExists(ByVal sDirName As String) As Boolean

            Try
                Dim dDir As New DirectoryInfo(AddDirSep(sDirName))
                Return dDir.Exists
            Catch
                Return False
            End Try

        End Function

        Public Shared Function FileExists(ByVal sFilename As String) As Boolean

            Try
                If Len(sFilename) = 0 Then Return False
                Dim fFile As New FileInfo(sFilename)
                Return fFile.Exists
            Catch
                Return False
            End Try

        End Function

        Public Shared Function HtmlEncode(ByVal text As String) As String

            If Len(text) = 0 Then Return ""

            Dim sOut As String = System.Net.WebUtility.HtmlEncode(HtmlDecode(text))

            Dim chars() As Char = sOut.ToCharArray
            Dim tOut As New StringBuilder(chars.Length * 2)

            For Each c As Char In chars

                Dim value As Integer = Convert.ToInt32(c)

                If (value > 31) And (value < 127) And (value <> 96) Then
                    tOut.Append(c)
                Else
                    tOut.Append("&#")
                    tOut.Append(value)
                    tOut.Append(";")
                End If

            Next

            Return tOut.ToString

        End Function

        Public Shared Function HtmlDecode(ByVal text As String) As String

            If Len(text) = 0 Then Return ""

            Return System.Net.WebUtility.HtmlDecode(text.Replace("&amp;", "&")).Replace(vbLf, "").Replace(vbCr, "").Replace(vbTab, "")

        End Function

        Public Shared Function SafeHref(ByVal Text As String) As String

            Return AddHttp(Text).Replace(Chr(34), "%22").Replace("`", "%60").Replace("'", "%27")

        End Function

        Public Shared Function AddHttp(ByVal Text As String) As String

            If Not HasHttp(Text) And Text.Length > 0 Then Return "http://" & Text
            Return Text

        End Function

        Public Shared Function HasHttp(ByVal Text As String) As Boolean

            If Text.IndexOf(":") > 1 Then

                Select Case Split(Text, ":")(0).ToLower
                    Case "http", "https"
                        Return True
                End Select

            End If

            Return False

        End Function

        Public Shared Function URLEncode(ByVal Text As String) As String

            Dim i As Integer
            Dim aCode As Integer
            Dim xURLEncode As String = Text

            For i = Len(xURLEncode) To 1 Step -1

                aCode = Asc(Mid$(xURLEncode, i, 1))

                Select Case aCode

                    Case 48 To 57, 65 To 90, 97 To 122
                    ' don't touch alphanumeric chars

                    Case 32
                        ' replace space with "+"
                        Mid$(xURLEncode, i, 1) = "+"

                    Case 42, 46, 47, 58, 95 ' / : . - _
                    ' sla over

                    Case Else

                        ' replace punctuation chars with "%hex"
                        xURLEncode = Microsoft.VisualBasic.Strings.Left(xURLEncode, i - 1) & "%" & Hex$(aCode) & Mid$(xURLEncode, i + 1)

                End Select
            Next

            Return xURLEncode

        End Function

        Public Shared Function SafeName(ByVal Text As String) As String

            Dim i As Integer
            Dim aCode As Integer
            Dim sOut As String = ""
            Dim AllowNumeric As Boolean = False

            Text = Text.Trim
            If IsNumeric(Text.Replace(".", "").Replace(":", "")) Then AllowNumeric = True

            For i = Len(Text) To 1 Step -1
                aCode = Asc(Mid$(Text, i, 1))
                Select Case aCode

                    Case 48 To 57
                        ' cijfers niet toestaan (ivm reader13.eweka.nl, etc.)

                        If AllowNumeric Then
                            sOut = Chr(aCode) + sOut
                        End If

                    Case 65 To 90, 97 To 122    ' lettters
                        sOut = Chr(aCode) + sOut

                    Case 97 To 122 ' lettters
                        sOut = Chr(aCode - 32) + sOut

                    Case 46 ' puntje
                        sOut = Chr(aCode) + sOut

                    Case Else

                        ' Doe niets

                End Select
            Next

            Return sOut

        End Function

        Public Shared Function URLDecode(ByVal StringToDecode As String) As String

            Dim TempAns As String = ""
            Dim CurChr As Integer = 1

            Do Until CurChr - 1 = Len(StringToDecode)
                Select Case Mid(StringToDecode, CurChr, 1)
                    Case "+"
                        TempAns = TempAns & " "
                    Case "%"
                        TempAns = TempAns & Chr(CInt(Val("&h" & Mid(StringToDecode, CurChr + 1, 2))))
                        CurChr = CurChr + 2
                    Case Else
                        TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
                End Select
                CurChr = CurChr + 1
            Loop

            URLDecode = TempAns

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
                                Case 2 : Return sIIF(bStrict, "", "Compilatie")
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
                                Case 10 : Return sIIF(bStrict, "", "Kleinkunst")
                                Case 11 : Return "Hollands"
                                Case 12 : Return sIIF(bStrict, "", "New Age")
                                Case 13 : Return "Pop"
                                Case 14 : Return "RnB"
                                Case 15 : Return "Hiphop"
                                Case 16 : Return "Reggae"
                                Case 17 : Return "Religieus"
                                Case 18 : Return "Rock"
                                Case 19 : Return "Soundtrack"
                                Case 20 : Return "" ''"Anders"
                                Case 21 : Return sIIF(bStrict, "", "Hardstyle")
                                Case 22 : Return sIIF(bStrict, "", "Aziatisch")
                                Case 23 : Return "Disco"
                                Case 24 : Return "Classics"
                                Case 25 : Return "Metal"
                                Case 26 : Return "Country"
                                Case 27 : Return "Dubstep"
                                Case 28 : Return sIIF(bStrict, "", "Nederhop")
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
                                Case 16 : Return sIIF(bStrict, "", "Nintendo 3DS")
                                Case Else : Return ""
                            End Select

                        Case "b"

                            Select Case CInt(sCat.Substring(1))
                                Case 1 : Return "Rip"
                                Case 0 : Return sIIF(bStrict, "", "ISO")
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
                                Case 3 : Return sIIF(bStrict, "", "CD/DVD Tools")
                                Case 4 : Return sIIF(bStrict, "", "Media spelers")
                                Case 5 : Return sIIF(bStrict, "", "Rippers & Encoders")
                                Case 6 : Return sIIF(bStrict, "", "Plugins")
                                Case 7 : Return sIIF(bStrict, "", "Database tools")
                                Case 8 : Return sIIF(bStrict, "", "Email software")
                                Case 9 : Return "Foto"
                                Case 10 : Return sIIF(bStrict, "", "Screensavers")
                                Case 11 : Return sIIF(bStrict, "", "Skin software")
                                Case 12 : Return sIIF(bStrict, "", "Drivers")
                                Case 13 : Return sIIF(bStrict, "", "Browsers")
                                Case 14 : Return sIIF(bStrict, "", "Download managers")
                                Case 15 : Return "Download"
                                Case 16 : Return sIIF(bStrict, "", "Usenet software")
                                Case 17 : Return sIIF(bStrict, "", "RSS Readers")
                                Case 18 : Return sIIF(bStrict, "", "FTP software")
                                Case 19 : Return sIIF(bStrict, "", "Firewalls")
                                Case 20 : Return sIIF(bStrict, "", "Antivirus software")
                                Case 21 : Return sIIF(bStrict, "", "Antispyware software")
                                Case 22 : Return sIIF(bStrict, "", "Optimalisatiesoftware")
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
                                Case 4 : Return sIIF(bStrict, "", "HD Overig")
                                Case 5 : Return "ePub"
                                Case 6 : Return "Bluray"
                                Case 7 : Return sIIF(bStrict, "", "HD-DVD")
                                Case 8 : Return sIIF(bStrict, "", "WMV HD")
                                Case 9 : Return "x264"
                                Case 10 : Return "DVD9"
                                Case Else : Return ""
                            End Select

                        Case "b"

                            Select Case CInt(sCat.Substring(1))
                                Case 4 : Return sIIF(bStrict, "", "TV")
                                Case 1 : Return sIIF(bStrict, "", "(S)VCD")
                                Case 6 : Return sIIF(bStrict, "", "Satelliet")
                                Case 2 : Return sIIF(bStrict, "", "Promo")
                                Case 3 : Return "Retail"
                                Case 7 : Return "R5"
                                Case 0 : Return "Cam"
                                Case 8 : Return sIIF(bStrict, "", "Telecine")
                                Case 9 : Return "Telesync"
                                Case 5 : Return "" '' "Anders"
                                Case 10 : Return "Scan"
                                Case Else : Return ""
                            End Select

                        Case "c"

                            Select Case CInt(sCat.Substring(1))
                                Case 0 : Return "Geen ondertitels"
                                Case 3 : Return "Engels ondertiteld (extern)"
                                Case 4 : Return sIIF(hCat <> 5, "Engels ondertiteld (ingebakken)", "Engels geschreven")
                                Case 7 : Return "Engels ondertiteld (instelbaar)"
                                Case 1 : Return "Nederlands ondertiteld (extern)"
                                Case 2 : Return sIIF(hCat <> 5, "Nederlands ondertiteld (ingebakken)", "Nederlands geschreven")
                                Case 6 : Return "Nederlands ondertiteld (instelbaar)"
                                Case 10 : Return "Engels gesproken"
                                Case 11 : Return "Nederlands gesproken"
                                Case 12 : Return sIIF(hCat <> 5, "Duits gesproken", "Duits geschreven")
                                Case 13 : Return sIIF(hCat <> 5, "Frans gesproken", "Frans geschreven")
                                Case 14 : Return sIIF(hCat <> 5, "Spaans gesproken", "Spaans geschreven")
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
                                Case 11 : Return sIIF(bStrict, "", "Televisie")
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

                                Case 75 : Return sIIF(bStrict, "", "Hetero")
                                Case 74 : Return sIIF(bStrict, "", "Homo")
                                Case 73 : Return sIIF(bStrict, "", "Lesbo")
                                Case 72 : Return sIIF(bStrict, "", "Bi")

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

        Public Shared Function GetFileContents(ByVal FullPath As String, Optional ByRef ErrInfo As String = "") As String

            Dim strContents As String
            Dim objReader As StreamReader

            Try

                objReader = New StreamReader(FullPath, LatinEnc)
                strContents = objReader.ReadToEnd()
                objReader.Close()

                Return strContents

            Catch Ex As Exception

                ErrInfo = Ex.Message
                Return vbNullString

            End Try

        End Function

        Public Shared Function StripNonAlphaNumericCharacters(ByVal sText As String) As String

            If Len(sText) = 0 Then Return ""
            Return System.Text.RegularExpressions.Regex.Replace(sText, "[^A-Za-z0-9]", "").Trim

        End Function

        Private Shared Function ConvertBytes(ByVal bytes As Long, ByVal convertTo As convTo) As Double

            If convTo.IsDefined(GetType(convTo), convertTo) Then

                Return bytes / (1024 ^ convertTo)

            Else

                Return -1 'An invalid value was passed to this function so exit

            End If

        End Function

        Public Shared Function ConvertSize(ByVal fileSize As Long) As String

            Dim sizeOfKB As Long = 1024              ' Actual size in bytes of 1KB
            Dim sizeOfMB As Long = 1048576           ' 1MB
            Dim sizeOfGB As Long = 1073741824        ' 1GB
            Dim sizeOfTB As Long = 1099511627776     ' 1TB
            Dim sizeofPB As Long = 1125899906842624  ' 1PB

            Dim tempFileSize As Double
            Dim tempFileSizeString As String

            Dim myArr() As Char = {CChar("0"), CChar(".")}  'Characters to strip off the end of our string after formating

            If fileSize < sizeOfKB Then 'Filesize is in Bytes
                tempFileSize = ConvertBytes(fileSize, convTo.B)
                If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
                tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr) ' Strip the 0's and 1's off the end of the string
                Return Math.Round(tempFileSize) & " bytes"

            ElseIf fileSize >= sizeOfKB And fileSize < sizeOfMB Then 'Filesize is in Kilobytes
                tempFileSize = ConvertBytes(fileSize, convTo.KB)
                If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
                tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
                Return Math.Round(tempFileSize) & " KB"

            ElseIf fileSize >= sizeOfMB And fileSize < sizeOfGB Then ' Filesize is in Megabytes
                tempFileSize = ConvertBytes(fileSize, convTo.MB)
                If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
                tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
                Return Math.Round(tempFileSize, 1) & " MB"

            ElseIf fileSize >= sizeOfGB And fileSize < sizeOfTB Then 'Filesize is in Gigabytes
                tempFileSize = ConvertBytes(fileSize, convTo.GB)
                If tempFileSize = -1 Then Return Nothing
                tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
                Return Math.Round(tempFileSize, 1) & " GB"

            ElseIf fileSize >= sizeOfTB And fileSize < sizeofPB Then 'Filesize is in Terabytes
                tempFileSize = ConvertBytes(fileSize, convTo.TB)
                If tempFileSize = -1 Then Return Nothing
                tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
                Return Math.Round(tempFileSize, 1) & " TB"

                'Anything bigger than that is silly ;)

            Else

                Return "" 'Invalid filesize so return Nothing

            End If

        End Function

        Public Shared Sub Foutje(ByVal sMsg As String, Optional ByVal sCaption As String = "Error")

            Throw New Exception(sCaption + ": " + sMsg)

        End Sub

        Public Shared Function MakeFilename(ByVal sIn As String) As String

            Return sIn.Replace(Chr(34), "").Replace("*", "").Replace(":", "").Replace("<", "").Replace(">", "").Replace("?", "").Replace("\", "").Replace("/", "").Replace("|", "").Replace(".", "").Replace("%", "").Replace("[", "").Replace("]", "").Replace(";", "").Replace("=", "").Replace(",", "")

        End Function

        Private Shared Function CreateHash(ByVal sLeft As String, ByVal sRight As String) As String

            Dim RetBytes() As Byte
            Dim CharCnt As Integer = 62
            Dim lByteCount As Integer = 5

            Dim SHA As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider()

            Dim OrgBytes() As Byte = MakeLatin(sLeft)
            Dim OrgBytes2() As Byte = MakeLatin(sRight)

            Dim OrgLen As Integer = UBound(OrgBytes) + 1
            Dim OrgLen2 As Integer = UBound(OrgBytes2) + 1

            Dim TryBytes(OrgLen + OrgLen2 - 2 + lByteCount) As Byte

            For Zl As Integer = 0 To UBound(OrgBytes)
                TryBytes(Zl) = OrgBytes(Zl)
            Next

            For Zl As Integer = UBound(TryBytes) - (OrgLen2 - 1) To UBound(TryBytes)
                TryBytes(Zl) = OrgBytes2(Zl - (UBound(TryBytes) - (OrgLen2 - 1)))
            Next

            Dim xPos1 As Integer = UBound(TryBytes) - 3 - OrgLen2
            Dim xPos2 As Integer = UBound(TryBytes) - 2 - OrgLen2
            Dim xPos3 As Integer = UBound(TryBytes) - 1 - OrgLen2
            Dim xPos4 As Integer = UBound(TryBytes) - OrgLen2

            Dim zR As New Random

            Dim Xp1B(CharCnt - 1) As Byte
            Dim Xp2B(CharCnt - 1) As Byte
            Dim Xp3B(CharCnt - 1) As Byte
            Dim Xp4B(CharCnt - 1) As Byte

            For i0 As Integer = 0 To CharCnt - 1
                Xp1B(i0) = CByte(i0)
                Xp2B(i0) = CByte(i0)
                Xp3B(i0) = CByte(i0)
                Xp4B(i0) = CByte(i0)
            Next

            Dim r As Integer = 0
            Dim t As Byte = 0

            For i0 As Integer = 0 To CharCnt - 1

                r = zR.Next(0, CharCnt)

                t = Xp1B(i0)
                Xp1B(i0) = Xp1B(r)
                Xp1B(r) = t

                r = zR.Next(0, CharCnt)

                t = Xp2B(i0)
                Xp2B(i0) = Xp2B(r)
                Xp2B(r) = t

                r = zR.Next(0, CharCnt)

                t = Xp3B(i0)
                Xp3B(i0) = Xp3B(r)
                Xp3B(r) = t

                r = zR.Next(0, CharCnt)

                t = Xp4B(i0)
                Xp4B(i0) = Xp4B(r)
                Xp4B(r) = t
            Next

            For i0 As Integer = 0 To CharCnt - 1
                Xp1B(i0) = GetBaseChar(Xp1B(i0))
                Xp2B(i0) = GetBaseChar(Xp2B(i0))
                Xp3B(i0) = GetBaseChar(Xp3B(i0))
                Xp4B(i0) = GetBaseChar(Xp4B(i0))
            Next

            For i1 As Integer = 0 To CharCnt - 1
                For i2 As Integer = 0 To CharCnt - 1
                    For i3 As Integer = 0 To CharCnt - 1
                        For i4 As Integer = 0 To CharCnt - 1

                            TryBytes(xPos1) = Xp1B(i1)
                            TryBytes(xPos2) = Xp2B(i2)
                            TryBytes(xPos3) = Xp3B(i3)
                            TryBytes(xPos4) = Xp4B(i4)

                            RetBytes = SHA.ComputeHash(TryBytes)

                            If (RetBytes(0) = 0) Then

                                If (RetBytes(1) = 0) Then

                                    Return GetLatin(TryBytes)

                                End If

                            End If

                        Next
                    Next
                Next
            Next

            Throw New Exception("Error 422")
            Return Nothing

        End Function

        Private Shared Function GetBaseChar(ByVal lIndex As Byte) As Byte

            Select Case lIndex
                Case 0 To 25
                    GetBaseChar = CByte(65 + lIndex)
                Case 26 To 51
                    GetBaseChar = CByte(97 + (lIndex - 26))
                Case 52 To 62
                    GetBaseChar = CByte(48 + (lIndex - 52))
                Case Else
                    GetBaseChar = CByte(65)
            End Select

        End Function

        Friend Shared Function sIIF(ByVal bExpression As Boolean, ByVal sTrue As String, ByVal sFalse As String) As String

            If bExpression Then
                Return sTrue
            Else
                Return sFalse
            End If

        End Function

        Friend Shared Function GetKey() As RSACryptoServiceProvider

            Static cKey As RSACryptoServiceProvider

            If cKey Is Nothing Then

                Dim CP As New CspParameters

                CP.KeyContainerName = Spotname & " User Key"
                CP.Flags = CType(CspProviderFlags.NoPrompt + CspProviderFlags.UseArchivableKey, CspProviderFlags)

                cKey = New RSACryptoServiceProvider(384, CP)

            End If

            Return cKey

        End Function

        Friend Shared Function GetModulus() As String

            Return Convert.ToBase64String(GetKey.ExportParameters(False).Modulus)

        End Function

        Private Shared Function DownloadString(ByVal zUrl As String, ByRef zError As String) As String

            Dim oWeb As System.Net.WebClient

            Try

                oWeb = New System.Net.WebClient
                Return (oWeb.DownloadString(zUrl))

            Catch ex As Exception
                zError = ex.Message
                Return vbNullString
            End Try

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

        'Public Shared Function FindNZB(ByVal sName As String, ByVal sNewsGroup As String, ByVal sTitle As String) As String

        '    Dim ssNzb As String = ""
        '    ''ssNzb = Spots.FindNZB(sName, 600, sNewsGroup, 999, True, False)

        '    If (Len(ssNzb) = 0) Then

        '        MsgBox("Kan NZB niet vinden, klik de bestandsnaam om handmatig te zoeken.", CType(MsgBoxStyle.Information + MsgBoxStyle.OkOnly, MsgBoxStyle), "NZB niet gevonden!")
        '        Return vbNullString

        '    Else

        '        Dim TmpFile As String
        '        Dim objReader As StreamWriter

        '        Dim sFile As String = Trim(MakeFilename(sTitle))

        '        If sFile.Length > 0 Then
        '            TmpFile = (System.IO.Path.GetTempPath & sFile & ".nzb")
        '        Else
        '            TmpFile = (System.IO.Path.GetTempFileName & ".nzb")
        '        End If

        '        Try
        '            objReader = New StreamWriter(TmpFile, False, LatinEnc)
        '            objReader.Write(ssNzb)
        '            objReader.Close()
        '        Catch Ex As Exception
        '            Foutje("Fout tijdens het schrijven.")
        '            Return vbNullString
        '        End Try

        '        Return TmpFile

        '    End If

        '    Return vbNullString

        'End Function

        Public Shared Function CreateMsgID(Optional ByVal sPrefix As String = "") As String

            Dim ZL(7) As Byte
            Dim ZK As New Random
            Dim sDomain As String = "spot.net"

            ZK.NextBytes(ZL)

            Dim Span As TimeSpan = (DateTime.UtcNow - EPOCH)
            Dim CurDate As Integer = CInt(Span.TotalSeconds)

            Dim sRandom As String = Convert.ToBase64String(ZL) & Convert.ToBase64String(BitConverter.GetBytes(CurDate))
            sRandom = sRandom.Replace("/", "s").Replace("+", "p").Replace("=", "")

            If Len(sPrefix) = 0 Then
                Return CreateHash("<" & sRandom, "@" & sDomain & ">")
            Else
                Return CreateHash("<" & sPrefix.Replace(".", "") & ".0." & sRandom & ".", "@" & sDomain & ">")
            End If

        End Function

        Public Shared ReadOnly Property LastPosition(ByVal db As Database, ByVal sTable As String) As Long

            Get

                Dim sErr As String = ""

                Try

                    Dim dbCmd As DbCommand = db.CreateCommand
                    dbCmd.CommandText = "SELECT MAX(rowid) FROM " & sTable

                    Dim tObj As Object = dbCmd.ExecuteScalar()

                    If IsDBNull(tObj) Then Return -1

                    Return CType(tObj, Long)

                Catch ex As Exception

                    Throw New Exception("LastPosition: " & ex.Message)

                End Try

            End Get

        End Property

        Friend Shared Function MakeMD5(ByVal sModulus As String) As String

            Static xMD5 As New MD5CryptoServiceProvider()

            If Len(sModulus) = 0 Then Return "Onbekend"

            Try

                Dim bBytes() As Byte = xMD5.ComputeHash(Convert.FromBase64String(sModulus))
                Dim sBuilder As New StringBuilder((UBound(bBytes) + 1) * 2, (UBound(bBytes) + 1) * 2)

                For i = 0 To bBytes.Length - 1
                    sBuilder.Append(bBytes(i).ToString("x2"))
                Next

                Return sBuilder.ToString()

            Catch
                Return "Onbekend"
            End Try

        End Function

        Friend Shared Sub ShowOnce(ByVal sMsg As String, ByVal sTitle As String)

            Static DidList As New HashSet(Of String)

            If DidList.Add(sMsg) Then
                MsgBox(sMsg, MsgBoxStyle.Information, sTitle)
            End If

        End Sub

        Friend Shared Function IsSearchQuery(ByVal sQuery As String) As Boolean

            Dim sTest As String = sQuery.ToLower.Trim
            Return sTest.Contains(" match ")

        End Function

    End Class

End Namespace
