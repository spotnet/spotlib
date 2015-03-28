Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Threading
Imports System.Security.Cryptography
Imports System.Globalization
Imports System.Security

Module sModule

    Friend Const MsgDomain As String = "spot.net"
    Friend Const CancelMSG As String = "Geannuleerd"
    Friend ReadOnly EPOCH As Date = New Date(1970, 1, 1, 0, 0, 0)

    Friend Function LatinEnc() As Encoding

        Return Encoding.GetEncoding(&H6FAF)

    End Function

    Friend Function GetLatin(ByRef zText() As Byte) As String

        Return LatinEnc.GetString(zText)

    End Function

    Friend Function MakeLatin(ByVal zText As String) As Byte()

        Return LatinEnc.GetBytes(zText)

    End Function

    Friend Function ConvertDate(ByVal sDate As String) As Date

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

    Friend Function FixPadding(ByRef sIn As String) As String

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

    Friend Function IsAscii(ByVal text As String, ByRef lPos As Integer) As Boolean

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

    Friend Function MakeAscii(ByVal text As String) As String

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

    Friend Function SplitBySizEx(ByVal strInput As String, ByVal iSize As Integer) As String()

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

    Friend Function UTFEncode(ByVal sInput As String) As String

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

    Private Function ParseEncodedWord(ByVal input As String) As String

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

    Friend Function Parse(ByVal input As String) As String

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

    Friend Function ZipStr(ByRef Inz() As Byte) As String

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

    Friend Function UnzipStr(ByRef Inz() As Byte) As String

        Try

            Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(Inz)
            Dim sw As System.IO.Compression.DeflateStream = New System.IO.Compression.DeflateStream(ms, System.IO.Compression.CompressionMode.Decompress, True)
            Dim reader As New StreamReader(sw, LatinEnc)

            Return reader.ReadToEnd

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    Public Function FileExists(ByVal sFilename As String) As Boolean

        Try
            If Len(sFilename) = 0 Then Return False
            Dim fFile As New FileInfo(sFilename)
            Return fFile.Exists
        Catch
            Return False
        End Try

    End Function

    Public Function GetFileContents(ByVal FullPath As String, Optional ByRef ErrInfo As String = "") As String

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

    Public Function IsNZB(ByVal _FileLnk As String, ByRef sSize As Long) As Boolean

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

    Public Function StripNonAlphaNumericCharacters(ByVal sText As String) As String

        Return System.Text.RegularExpressions.Regex.Replace(sText, "[^A-Za-z0-9]", "").Trim

    End Function

    Public Function AddHttp(ByVal Text As String) As String

        If Not HasHttp(Text) And Text.Length > 0 Then Return "http://" & Text
        Return Text

    End Function

    Public Function HasHttp(ByVal Text As String) As Boolean

        If Text.IndexOf(":") > 1 Then

            Select Case Split(Text, ":")(0).ToLower
                Case "http", "https"
                    Return True
            End Select

        End If

        Return False

    End Function

    Private Function URLEncode(ByVal Text As String) As String

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

    Private Function UnGZIP(ByRef Inz() As Byte) As String

        Try

            Dim ms As System.IO.MemoryStream = New System.IO.MemoryStream(Inz)
            Dim sw As System.IO.Compression.GZipStream = New System.IO.Compression.GZipStream(ms, System.IO.Compression.CompressionMode.Decompress, True)
            Dim reader As New StreamReader(sw, LatinEnc)
            Return reader.ReadToEnd

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    Public Sub Wait(ByVal ms As Integer)

        Using wh As New ManualResetEvent(False)
            wh.WaitOne(ms)
        End Using

    End Sub

    Friend Function CheckHash(ByVal sMsg As String) As Boolean

        Dim ShA As New SHA1Managed
        Dim ShABytes() As Byte = ShA.ComputeHash(MakeLatin(sMsg))

        If ShABytes(0) = 0 Then
            If ShABytes(1) = 0 Then

                Return True

            End If
        End If

        Return False

    End Function

    Private Function CreateHash(ByVal sLeft As String, ByVal sRight As String) As String

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

    Private Function GetBaseChar(ByVal lIndex As Byte) As Byte

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

    Private Function GetNZB(ByVal zPostData As String, ByVal OrgUrl As String, ByRef zError As String) As String

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

    Public Function CheckUserSignature(ByVal sOrg As String, ByVal sSignature As String, ByVal sUserKey As String) As Boolean

        Try

            Dim RSAalg As RSACryptoServiceProvider = MakeRSA(sUserKey)

            If RSAalg Is Nothing Then Return False

            Return RSAalg.VerifyHash((New SHA1Managed).ComputeHash(MakeLatin(sOrg)), Nothing, Convert.FromBase64String(FixPadding(sSignature)))

        Catch
        End Try

        Return False

    End Function

    Private Sub LaunchBrowser(ByVal sUrl As String)

        Try
            System.Diagnostics.Process.Start(sUrl)
        Catch
        End Try

    End Sub

    Friend Function SpecialString(ByVal sDataIn As String) As String

        Return sDataIn.Replace("/", "-s").Replace("+", "-p").Replace("=", "")

    End Function

    Friend Function UnSpecialString(ByVal sDataIn As String) As String

        Return sDataIn.Replace("-s", "/").Replace("-p", "+")

    End Function

    Public Function MakeMsg(ByVal sMes As String, Optional ByVal Tag As Boolean = True) As String

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

    Public Function HtmlEncode(ByVal text As String) As String

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

    Public Function HtmlDecode(ByVal text As String) As String

        If Len(text) = 0 Then Return ""

        Return System.Net.WebUtility.HtmlDecode(text.Replace("&amp;", "&")).Replace(vbLf, "").Replace(vbCr, "").Replace(vbTab, "")

    End Function

    Friend Function CreateWork(ByVal sFirst As Long, ByVal sLast As Long) As List(Of NNTPWork)

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

    Friend Function FormatLong(ByVal zLong As Long) As String

        Dim sT As String
        sT = FormatLong2(zLong)
        If sT = "0" Then Return "Geen"
        Return sT

    End Function

    Friend Function FormatLong2(ByVal zLong As Long) As String

        If zLong = 0 Then Return "0"
        Return zLong.ToString("#,#", CultureInfo.InvariantCulture).Replace(",", ".")

    End Function

    Friend Function SplitLines(ByVal sIn As String, ByVal AllowBlankLines As Boolean, ByVal lMax As Integer) As List(Of String)

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

    Friend Function SplitLinesGZIP(ByVal sIn As String) As List(Of String)

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

    Friend Function SplitLinesXML(ByVal sIn As String, ByVal sPrefix As String, ByVal lMax As Integer) As String

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

    Friend Function GetLocation(ByVal URL As String) As String

        Try

            Dim myWebRequest As Net.WebRequest = Net.WebRequest.Create(URL)

            myWebRequest.Proxy = Nothing

            Dim myWebResponse As Net.WebResponse = myWebRequest.GetResponse()
            Return myWebResponse.ResponseUri.Host

        Catch ex As Exception
            Return vbNullString
        End Try

    End Function

    Friend Function GetHash(ByVal zPostData As String, ByRef SignServer As String, ByVal SignCmd As String, ByRef zError As String) As String

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

    Friend Function CreateUserSignature(ByVal sDataIn As String, ByVal cRSA As RSACryptoServiceProvider) As String

        Return SpecialString(Convert.ToBase64String(cRSA.SignHash((New SHA1Managed).ComputeHash(MakeLatin(sDataIn)), Nothing)))

    End Function

    Friend Function CheckFrom(ByVal sFrom As String) As Boolean

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

    Friend Function PostData(ByVal tPhuse As Phuse.Engine, ByRef zInput As List(Of String), ByVal zSub As String, ByVal zFrom As String, ByVal zGroup As String, ByVal zExtra As String, ByRef xOutID As String, ByVal MsgID As String, ByRef sError As String) As Boolean

        Dim cNNTP As NNTP
        Dim XC As Long = 0

        Dim sRet As Boolean = False
        Dim zNZB As String = vbNullString
        Dim sTmp As String = vbNullString
        Dim xOutList As String = vbNullString
        Dim TheData As String = vbNullString

        Dim TempId As String = ""

        cNNTP = New NNTP(tPhuse)

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

    Friend Function CreateMsgID(Optional ByVal sPrefix As String = "") As String

        Dim ZL(7) As Byte
        Dim ZK As New Random

        ZK.NextBytes(ZL)

        Dim Span As TimeSpan = (DateTime.UtcNow - EPOCH)
        Dim CurDate As Integer = CInt(Span.TotalSeconds)

        Dim sRandom As String = Convert.ToBase64String(ZL) & Convert.ToBase64String(BitConverter.GetBytes(CurDate))
        sRandom = sRandom.Replace("/", "s").Replace("+", "p").Replace("=", "")

        If Len(sPrefix) = 0 Then
            Return CreateHash("<" & sRandom, "@" & MsgDomain & ">")
        Else
            Return CreateHash("<" & sPrefix.Replace(".", "") & ".0." & sRandom & ".", "@" & MsgDomain & ">")
        End If

    End Function

    Friend Function CreateSpot(ByVal sPoster As String, ByVal sTitle As String, ByVal sNZB As String, ByVal lFileSize As Long, ByVal sLink As String, ByVal sImageID As String, ByVal sDesc As String, ByVal HCat As Byte, ByVal sCat As String, ByVal sTag As String, ByRef xOut() As String, ByRef SignServer As String, ByVal lImgX As Long, ByVal lImgY As Long, ByRef zErr As String) As Boolean

        Dim zPost As String = ""

        zPost = "cat=" & Str(HCat).Trim & "&"
        zPost += "sub=" & StripNonAlphaNumericCharacters(sCat).ToLower & "&"
        zPost += "title=" & MakeP(sTitle) & "&"
        zPost += "from=" & StripNonAlphaNumericCharacters(sPoster) & "&"
        zPost += "desc=" & MakeP(sDesc) & "&"
        zPost += "tag=" & StripNonAlphaNumericCharacters(sTag) & "&"
        zPost += "msgid=" & MakeP(sNZB) & "&"
        If Len(sLink) > 0 Then zPost += "link=" & MakeP(sLink) & "&"
        zPost += "img=" & MakeP(sImageID) & "&"
        zPost += "imgx=" & Str(lImgX).Trim & "&"
        zPost += "imgy=" & Str(lImgY).Trim & "&"
        zPost += "size=" & Str(lFileSize).Trim & "&"

        Dim sResult As String = GetHash(zPost, SignServer, "signxml", zErr)

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

    Friend Function MakeP(ByVal sIn As String) As String

        Return Convert.ToBase64String(MakeLatin(sIn)).Replace("=", "%3d").Replace("+", "%2b").Replace("&", "%26").Replace$("/", "%2f").Trim

    End Function

    Friend Function sIIF(ByVal bExpression As Boolean, ByVal sTrue As String, ByVal sFalse As String) As String

        If bExpression Then
            Return sTrue
        Else
            Return sFalse
        End If

    End Function

    Friend Function GetBinary(ByVal tPhuse As Phuse.Engine, ByVal Newsgroup As String, ByVal xMsgID As List(Of String), ByRef sxOut() As Byte, ByRef sError As String) As Boolean

        Dim cNNTP As NNTP
        Dim Retr As New StringBuilder

        If xMsgID Is Nothing Then Return False
        If xMsgID.Count < 1 Then Return False

        cNNTP = New NNTP(tPhuse)

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

    Friend Function GetRSA(ByVal TrustedKeys() As String) As RSACryptoServiceProvider()

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

    Friend Function MakeRSA(ByVal sModulus As String) As RSACryptoServiceProvider

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

    Friend Function IsEro(ByVal sVal As String) As Boolean

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

    Friend Function IsEbook(ByVal sVal As String) As Boolean

        If sVal.Contains("a5|") Then Return True
        If sVal.Contains("z2|") Then Return True

        Return False

    End Function

    Friend Function IsTV(ByVal sVal As String) As Boolean

        If sVal.Contains("b4|") Then Return True
        If sVal.Contains("d11|") Then Return True
        If sVal.Contains("z1|") Then Return True

        Return False

    End Function

    Friend Function GetCode(ByVal sResponse As String) As Integer

        Try
            Return CInt(Val(Left(sResponse, 3)))
        Catch ex As Exception
            Return 0
        End Try

    End Function

End Module
