Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Globalization
Imports System.Security.Cryptography

Namespace Spotlib

    Public Class Utils

        Public Shared ReadOnly EPOCH As Date = New Date(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
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

        Public Shared Function LatinEnc() As Encoding

            Return Encoding.GetEncoding(&H6FAF)

        End Function

        Public Shared Function GetLatin(ByRef zText() As Byte) As String

            Return LatinEnc.GetString(zText)

        End Function

        Public Shared Function MakeLatin(ByVal zText As String) As Byte()

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

        Public Shared Function UnGZIP(ByRef Inz() As Byte) As String

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

        Friend Shared Function SpecialString(ByVal sDataIn As String) As String

            Return sDataIn.Replace("/", "-s").Replace("+", "-p").Replace("=", "")

        End Function

        Friend Shared Function UnSpecialString(ByVal sDataIn As String) As String

            Return sDataIn.Replace("-s", "/").Replace("-p", "+")

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

        Friend Shared Function MakeP(ByVal sIn As String) As String

            Return Convert.ToBase64String(MakeLatin(sIn)).Replace("=", "%3d").Replace("+", "%2b").Replace("&", "%26").Replace$("/", "%2f").Trim

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

        Public Shared Function GetFolder(ByVal pGuid As Guid) As String

            Dim txtOldTmp As String = ""
            Dim ptPath As IntPtr

            If SHGetKnownFolderPath(pGuid, 0, IntPtr.Zero, ptPath) = 0 Then
                txtOldTmp = System.Runtime.InteropServices.Marshal.PtrToStringUni(ptPath)
                System.Runtime.InteropServices.Marshal.FreeCoTaskMem(ptPath)
            End If

            GetFolder = txtOldTmp

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

        Public Shared Function StripChars(ByVal sText As String) As String

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

        Public Shared Function CreateHash(ByVal sLeft As String, ByVal sRight As String) As String

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

        Public Shared Function sIIF(ByVal bExpression As Boolean, ByVal sTrue As String, ByVal sFalse As String) As String

            If bExpression Then
                Return sTrue
            Else
                Return sFalse
            End If

        End Function

        Public Shared Function DownloadString(ByVal zUrl As String, ByRef zError As String) As String

            Dim oWeb As System.Net.WebClient

            Try

                oWeb = New System.Net.WebClient
                Return (oWeb.DownloadString(zUrl))

            Catch ex As Exception
                zError = ex.Message
                Return vbNullString
            End Try

        End Function

        Public Shared Function MakeMD5(ByVal sModulus As String) As String

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

        Public Shared Function GetCrc() As CRC32

            Static CC As CRC32 = Nothing

            If CC Is Nothing Then CC = New CRC32

            Return CC

        End Function

    End Class

    Public Class CRC32

        Private crcLookup() As Integer = {
        &H0, &H77073096, &HEE0E612C, &H990951BA,
        &H76DC419, &H706AF48F, &HE963A535, &H9E6495A3,
        &HEDB8832, &H79DCB8A4, &HE0D5E91E, &H97D2D988,
        &H9B64C2B, &H7EB17CBD, &HE7B82D07, &H90BF1D91,
        &H1DB71064, &H6AB020F2, &HF3B97148, &H84BE41DE,
        &H1ADAD47D, &H6DDDE4EB, &HF4D4B551, &H83D385C7,
        &H136C9856, &H646BA8C0, &HFD62F97A, &H8A65C9EC,
        &H14015C4F, &H63066CD9, &HFA0F3D63, &H8D080DF5,
        &H3B6E20C8, &H4C69105E, &HD56041E4, &HA2677172,
        &H3C03E4D1, &H4B04D447, &HD20D85FD, &HA50AB56B,
        &H35B5A8FA, &H42B2986C, &HDBBBC9D6, &HACBCF940,
        &H32D86CE3, &H45DF5C75, &HDCD60DCF, &HABD13D59,
        &H26D930AC, &H51DE003A, &HC8D75180, &HBFD06116,
        &H21B4F4B5, &H56B3C423, &HCFBA9599, &HB8BDA50F,
        &H2802B89E, &H5F058808, &HC60CD9B2, &HB10BE924,
        &H2F6F7C87, &H58684C11, &HC1611DAB, &HB6662D3D,
        &H76DC4190, &H1DB7106, &H98D220BC, &HEFD5102A,
        &H71B18589, &H6B6B51F, &H9FBFE4A5, &HE8B8D433,
        &H7807C9A2, &HF00F934, &H9609A88E, &HE10E9818,
        &H7F6A0DBB, &H86D3D2D, &H91646C97, &HE6635C01,
        &H6B6B51F4, &H1C6C6162, &H856530D8, &HF262004E,
        &H6C0695ED, &H1B01A57B, &H8208F4C1, &HF50FC457,
        &H65B0D9C6, &H12B7E950, &H8BBEB8EA, &HFCB9887C,
        &H62DD1DDF, &H15DA2D49, &H8CD37CF3, &HFBD44C65,
        &H4DB26158, &H3AB551CE, &HA3BC0074, &HD4BB30E2,
        &H4ADFA541, &H3DD895D7, &HA4D1C46D, &HD3D6F4FB,
        &H4369E96A, &H346ED9FC, &HAD678846, &HDA60B8D0,
        &H44042D73, &H33031DE5, &HAA0A4C5F, &HDD0D7CC9,
        &H5005713C, &H270241AA, &HBE0B1010, &HC90C2086,
        &H5768B525, &H206F85B3, &HB966D409, &HCE61E49F,
        &H5EDEF90E, &H29D9C998, &HB0D09822, &HC7D7A8B4,
        &H59B33D17, &H2EB40D81, &HB7BD5C3B, &HC0BA6CAD,
        &HEDB88320, &H9ABFB3B6, &H3B6E20C, &H74B1D29A,
        &HEAD54739, &H9DD277AF, &H4DB2615, &H73DC1683,
        &HE3630B12, &H94643B84, &HD6D6A3E, &H7A6A5AA8,
        &HE40ECF0B, &H9309FF9D, &HA00AE27, &H7D079EB1,
        &HF00F9344, &H8708A3D2, &H1E01F268, &H6906C2FE,
        &HF762575D, &H806567CB, &H196C3671, &H6E6B06E7,
        &HFED41B76, &H89D32BE0, &H10DA7A5A, &H67DD4ACC,
        &HF9B9DF6F, &H8EBEEFF9, &H17B7BE43, &H60B08ED5,
        &HD6D6A3E8, &HA1D1937E, &H38D8C2C4, &H4FDFF252,
        &HD1BB67F1, &HA6BC5767, &H3FB506DD, &H48B2364B,
        &HD80D2BDA, &HAF0A1B4C, &H36034AF6, &H41047A60,
        &HDF60EFC3, &HA867DF55, &H316E8EEF, &H4669BE79,
        &HCB61B38C, &HBC66831A, &H256FD2A0, &H5268E236,
        &HCC0C7795, &HBB0B4703, &H220216B9, &H5505262F,
        &HC5BA3BBE, &HB2BD0B28, &H2BB45A92, &H5CB36A04,
        &HC2D7FFA7, &HB5D0CF31, &H2CD99E8B, &H5BDEAE1D,
        &H9B64C2B0, &HEC63F226, &H756AA39C, &H26D930A,
        &H9C0906A9, &HEB0E363F, &H72076785, &H5005713,
        &H95BF4A82, &HE2B87A14, &H7BB12BAE, &HCB61B38,
        &H92D28E9B, &HE5D5BE0D, &H7CDCEFB7, &HBDBDF21,
        &H86D3D2D4, &HF1D4E242, &H68DDB3F8, &H1FDA836E,
        &H81BE16CD, &HF6B9265B, &H6FB077E1, &H18B74777,
        &H88085AE6, &HFF0F6A70, &H66063BCA, &H11010B5C,
        &H8F659EFF, &HF862AE69, &H616BFFD3, &H166CCF45,
        &HA00AE278, &HD70DD2EE, &H4E048354, &H3903B3C2,
        &HA7672661, &HD06016F7, &H4969474D, &H3E6E77DB,
        &HAED16A4A, &HD9D65ADC, &H40DF0B66, &H37D83BF0,
        &HA9BCAE53, &HDEBB9EC5, &H47B2CF7F, &H30B5FFE9,
        &HBDBDF21C, &HCABAC28A, &H53B39330, &H24B4A3A6,
        &HBAD03605, &HCDD70693, &H54DE5729, &H23D967BF,
        &HB3667A2E, &HC4614AB8, &H5D681B02, &H2A6F2B94,
        &HB40BBE37, &HC30C8EA1, &H5A05DF1B, &H2D02EF8D}

        Public Function Calculate(ByVal s As String, ByVal e As System.Text.Encoding) As Integer

            Dim buffer() As Byte = e.GetBytes(s)
            Return Calculate(buffer)

        End Function

        Public Function Calculate(ByVal b() As Byte) As Integer

            Dim result As Integer = &HFFFFFFFF

            Dim len As Integer = b.Length
            Dim lookup As Integer

            For i As Integer = 0 To len - 1
                lookup = (result And &HFF) Xor b(i)
                result = ((result And &HFFFFFF00) \ &H100) And &HFFFFFF
                result = result Xor crcLookup(lookup)
            Next i

            Return Not (result)

        End Function

    End Class

End Namespace