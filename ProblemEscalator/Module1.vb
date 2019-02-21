Imports System.Diagnostics
Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Configuration
Imports System.Text
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Security.Principal
Imports System.Security.Cryptography
Imports System.Management

Module Module1

    Public Sub LogEvent(ByVal sMessage As String, ByVal strEventType As EventLogEntryType)
        ' Write error into the event viewer
        Try
            Dim oEventLog As EventLog = New EventLog("Application")
            If Not Diagnostics.EventLog.SourceExists("NetkaQuartz Escalator") Then
                Diagnostics.EventLog.CreateEventSource("NetkaQuartz Escalator", "Application")
            End If
            Diagnostics.EventLog.WriteEntry("NetkaQuartz Escalator", sMessage, strEventType)
        Catch e As Exception
        End Try
    End Sub

    Public Sub Debug(ByVal log As String)
        Dim ObjFile As New FileInfo("c:\netkaquartz\NetkaQuartz Escalator.log")
        Dim ObjStreamWriter As StreamWriter = ObjFile.AppendText
        ObjStreamWriter.WriteLine(Now & " " & log)
        ObjStreamWriter.Close()
    End Sub

    Public Function DateStamp() As String
        DateStamp = Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
    End Function

    Public Function TimeStamp() As String
        TimeStamp = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    End Function

    Public Function FormatDate(ByVal strDate As String, ByVal strOldFormat As String, ByVal strNewFormat As String) As String
        ' Sample of strOldFormat or strNewFormat = yyMMdd, dd-MMM-yy
        Dim dt As DateTime = DateTime.ParseExact(strDate, strOldFormat, Nothing)
        FormatDate = dt.ToString(strNewFormat)
    End Function

    Public Function FetchField(ByVal sql As String, ByVal strConnection As String, ByVal field As String) As Object
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            FetchField = ds.Tables("*").Rows(0).Item(field)
        Catch
            FetchField = Nothing
        End Try
    End Function
    Function ReplaceString(ByVal str As String) As String
        Dim temp As String = str
        Try
            temp = Replace(temp, "'", "\'")
            temp = Replace(temp, ",", "\,")
        Catch ex As Exception
            ReplaceString = ""
        End Try
        Return temp
    End Function

    Public Function RowCountSql(ByVal sql As String, ByVal strConnection As String, ByVal conn As OdbcConnection) As Long
        Try
            ' Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            RowCountSql = ds.Tables("*").Rows.Count
        Catch
            RowCountSql = 0
        End Try
    End Function

    Public Function RowCountTable(ByVal tb As String, ByVal strConnection As String) As Long
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim sql As String = "SELECT * FROM " & tb
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, tb)
            RowCountTable = ds.Tables(tb).Rows.Count
        Catch
            RowCountTable = 0
        End Try
    End Function

    Public Function FindLastID(ByVal tb As String, ByVal strConnection As String, ByVal field As String) As Long
        Dim id As Integer
        Try
            Dim sql As String = "SELECT " & field & " FROM " & tb & " ORDER BY " & field & " DESC"
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            If ds.Tables("*").Rows.Count > 0 Then
                id = ds.Tables("*").Rows(0).Item(field)
            End If
        Catch
            id = 0
        End Try
        FindLastID = id
    End Function

    Public Function AESEncrypt(ByVal plainText As String, ByVal key As String, ByVal iv As String) As String
        Dim keyBytes() As Byte = Encoding.ASCII.GetBytes(key)
        Dim initVectorBytes() As Byte = Encoding.ASCII.GetBytes(iv)

        Dim symmetricKey As New RijndaelManaged()
        symmetricKey.Mode = CipherMode.CBC
        symmetricKey.Padding = PaddingMode.Zeros

        'Get an encryptor.
        Dim encryptor As ICryptoTransform = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)

        'Encrypt the data.
        Dim memoryStream As New MemoryStream()
        Dim cryptoStream As New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)

        'Convert the data to a byte array.
        Dim plainTextBytes() As Byte
        plainTextBytes = Encoding.UTF8.GetBytes(plainText)

        'Write all data to the crypto stream and flush it.
        cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)
        cryptoStream.FlushFinalBlock()

        'Get encrypted array of bytes.
        Dim cipherTextBytes() As Byte = memoryStream.ToArray()
        Dim cipherText As String = System.Convert.ToBase64String(cipherTextBytes)
        memoryStream.Close()
        cryptoStream.Close()
        Return cipherText
    End Function

    Public Function AESDecrypt(ByVal cipherText As String, ByVal key As String, ByVal iv As String) As String
        Dim keyBytes() As Byte = Encoding.ASCII.GetBytes(key)
        Dim initVectorBytes() As Byte = Encoding.ASCII.GetBytes(iv)
        Dim cipherTextBytes() As Byte = System.Convert.FromBase64String(cipherText)

        Dim symmetricKey As New RijndaelManaged()
        symmetricKey.Mode = CipherMode.CBC
        symmetricKey.Padding = PaddingMode.Zeros

        'Get an decryptor.
        Dim decryptor As ICryptoTransform = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)

        'Decrypt the data.
        Dim memoryStream As New MemoryStream(cipherTextBytes)
        Dim cryptoStream As New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)
        Dim plainTextBytes() As Byte
        plainTextBytes = New Byte(cipherTextBytes.Length) {}

        'Read the data out of the crypto stream.
        Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)
        Dim plainText As String = Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount)
        memoryStream.Close()
        cryptoStream.Close()

        'Replace null character
        plainText = Replace(plainText, Chr(0), "")

        Return plainText
    End Function

    Function SendEmail(ByVal strFrom As String, ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, ByVal strAttachment As String, ByVal smtp_server As String, ByVal smtp_sender As String, ByVal smtp_password As String, ByVal bSsl As Boolean, ByVal sslPort As Integer) As String

        Dim apppath As String = ConfigurationManager.AppSettings("AppPath")
        Dim mail As New MailMessage()
        Dim avHTMLBody As AlternateView
        mail.From = New System.Net.Mail.MailAddress(smtp_sender, strFrom)
        Dim arrTo() As String = Split(strTo, ";")
        Dim intUBound As Integer = UBound(arrTo)
        Dim i As Integer
        If intUBound >= 0 Then
            For i = 0 To intUBound
                If arrTo(i) <> "" Then
                    mail.To.Add(arrTo(i))
                    'mail.Bcc.Add(arrTo(i))
                End If
            Next
        End If

        mail.Subject = strSubj
        '        mail.Body = strBody
        mail.IsBodyHtml = True
        If strAttachment <> "" Then
            mail.Attachments.Add(New Attachment(strAttachment))
        End If

        Dim img1 As LinkedResource = New LinkedResource(apppath & "\image\email_footer.jpg", MediaTypeNames.Image.Jpeg)
        img1.ContentId = "companylogo"
        avHTMLBody = AlternateView.CreateAlternateViewFromString(strBody, Nothing, MediaTypeNames.Text.Html)
        avHTMLBody.LinkedResources.Add(img1)
        mail.AlternateViews.Add(avHTMLBody)

        Dim smtp As New SmtpClient()
        smtp.Host = smtp_server
        smtp.Credentials = New System.Net.NetworkCredential(smtp_sender, smtp_password)
        If bSsl Then
            smtp.EnableSsl = True
            smtp.Port = sslPort     'ssl use 465, tls use 587
        End If
        Try
            smtp.Send(mail)
            SendEmail = "OK"
        Catch ex As System.Net.Mail.SmtpException
            SendEmail = "Error: " + ex.Message
        End Try
    End Function
    Function SendEmailSat(ByVal strFrom As String, ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, ByVal strAttachment As String, ByVal smtp_server As String, ByVal smtp_sender As String, ByVal smtp_password As String, ByVal bSsl As Boolean, ByVal sslPort As Integer) As String

        Dim apppath As String = ConfigurationManager.AppSettings("AppPath")
        Dim mail As New MailMessage()
        Dim avHTMLBody As AlternateView
        mail.From = New System.Net.Mail.MailAddress(smtp_sender, strFrom)
        Dim arrTo() As String = Split(strTo, ";")
        Dim intUBound As Integer = UBound(arrTo)
        Dim i As Integer
        If intUBound >= 0 Then
            For i = 0 To intUBound
                If arrTo(i) <> "" Then
                    mail.To.Add(arrTo(i))
                    'mail.Bcc.Add(arrTo(i))
                End If
            Next
        End If

        mail.Subject = strSubj
        '        mail.Body = strBody
        mail.IsBodyHtml = True
        If strAttachment <> "" Then
            mail.Attachments.Add(New Attachment(strAttachment))
        End If

        Dim img1 As LinkedResource = New LinkedResource(apppath & "\image\email_footer.jpg", MediaTypeNames.Image.Jpeg)
        img1.ContentId = "companylogo"
        avHTMLBody = AlternateView.CreateAlternateViewFromString(strBody, Nothing, MediaTypeNames.Text.Html)
        avHTMLBody.LinkedResources.Add(img1)
        mail.AlternateViews.Add(avHTMLBody)

        Dim smtp As New SmtpClient()
        smtp.Host = smtp_server
        smtp.Credentials = New System.Net.NetworkCredential(smtp_sender, smtp_password)
        If bSsl Then
            smtp.EnableSsl = True
            smtp.Port = sslPort     'ssl use 465, tls use 587
        End If
        Try
            smtp.Send(mail)
            SendEmailSat = "OK"
        Catch ex As System.Net.Mail.SmtpException
            SendEmailSat = "Error: " + ex.Message
        End Try
    End Function

    Public Function HTTPPost(ByVal TheURL As String, ByVal postdata As String) As String
        Dim page As String = ""
        Try
            Dim Uri As Uri = New Uri(TheURL)
            Dim request As HttpWebRequest = DirectCast(WebRequest.Create(Uri), HttpWebRequest)

            'request.KeepAlive = False
            'request.ProtocolVersion = HttpVersion.Version10
            'request.Timeout = System.Threading.Timeout.Infinite
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"
            request.ContentLength = postdata.Length
            'request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.1.4322)"
            'request.Proxy = GlobalProxySelection.GetEmptyWebProxy()
            'request.ReadWriteTimeout = -1

            Dim writeStream As Stream = request.GetRequestStream()
            Dim bytes As Byte() = System.Text.Encoding.ASCII.GetBytes(postdata)
            writeStream.Write(bytes, 0, bytes.Length)
            writeStream.Close()

            Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
            Dim responseStream As Stream = response.GetResponseStream()
            Dim readStream As StreamReader = New StreamReader(responseStream, Encoding.UTF8)

            page = readStream.ReadToEnd()
        Catch ee As Exception
            page = ee.Message & vbCrLf & ee.StackTrace
        End Try

        Return page
    End Function

    Public Function MacAddress(ByVal ip As String) As String
        Dim strMc As String = ""
        Try
            Dim mc As System.Management.ManagementClass
            Dim mo As ManagementObject
            mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            For Each mo In moc
                If mo.Item("IPEnabled") Then
                    If mo.Item("IPAddress")(0).ToString() = ip Then
                        strMc = Replace(mo.Item("MacAddress").ToString(), ":", "")
                        Exit For
                    End If
                End If
            Next
        Catch
        End Try
        Return strMc
    End Function
    Function GetDurationTime(ByVal strtime As String) As String
        'Dim ts As TimeSpan 
        Dim tSpan As TimeSpan
        Dim totalSeconds As Double
        Dim arr() As String
        Dim strValue As String
        arr = Split(strtime, ":")
        totalSeconds = (CInt(arr(0)) * 3600) + (CInt(arr(1)) * 60) + CInt(arr(2))
        tSpan = TimeSpan.FromSeconds(totalSeconds)
        strValue = String.Format("{0} Days {1:0} Hrs {2:0} Mins", tSpan.Days, tSpan.Hours, tSpan.Minutes)
        If tSpan.Days = 0 Then
            strValue = String.Format("{0:0} Hrs {1:0} Mins", tSpan.Hours, tSpan.Minutes)
        End If
        If tSpan.Days = 0 And tSpan.Hours = 0 Then
            strValue = String.Format("{0:0} Mins", tSpan.Minutes)
        End If
        If tSpan.Days = 0 And tSpan.Hours = 0 And tSpan.Minutes = 0 Then
            strValue = "-"
        End If

        Return strValue
    End Function
End Module
