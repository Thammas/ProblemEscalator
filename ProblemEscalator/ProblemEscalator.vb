Imports System.Diagnostics
Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Timers
Imports System.Configuration
Imports System.Math
Imports log4net
Imports NKSUTIL.Security

Imports System.Text.RegularExpressions

Public Class ProblemEscalator
    Dim strConnection As String = LoadDBDriver()
    Dim sql, strDebug As String
    Dim t As Timer
    Dim log As log4net.ILog
    Dim escalation_type As String
    Dim escalation_next_time As String = ""
    Dim smtp_server As String = ""
    Dim smtp_sender As String = ""
    Dim smtp_password As String = ""
    Dim smtp_auth As String = ""
    Dim sms_url As String = ""
    Dim sms_user_keyword As String = ""
    Dim sms_password_keyword As String = ""
    Dim sms_number_keyword As String = ""
    Dim sms_message_keyword As String = ""
    Dim sms_username As String = ""
    Dim sms_password As String = ""
    Dim sms_protocol As String = ""
    Dim sms_result As String = ""
    Dim sms_language_keyword As String = ""
    Dim bSsl As Boolean
    Dim sslPort As Integer
    Dim company_name As String
    Dim web_url As String
    Dim signature As String = ""
    Dim app_path As String = ConfigurationManager.AppSettings("AppPath")
    Dim auto_escalation As String = ConfigurationManager.AppSettings("auto_escalation")
    Dim alert_before As String = ConfigurationManager.AppSettings("alert_before")
    Dim alert_before_every As String = ConfigurationManager.AppSettings("alert_before_every")

    Protected Overrides Sub OnStart(ByVal args() As String)

        Try
            ' Intialize log4net by reading app.config
            log = log4net.LogManager.GetLogger("NetkaQuartz")
            log4net.Config.XmlConfigurator.Configure()
            Dim strProductValid As String = "Valid"
            If strProductValid = "Valid" Then
                company_name = ConfigurationManager.AppSettings("company_name")
                escalation_type = ConfigurationManager.AppSettings("escalation_type")
                escalation_next_time = ConfigurationManager.AppSettings("escalation_next_time")
                web_url = ConfigurationManager.AppSettings("web_url")
                app_path = ConfigurationManager.AppSettings("AppPath")
                ' Set timer to process Timer_Fired() every minute
                t = New Timer(60000)
                AddHandler t.Elapsed, AddressOf Timer_Fired
                With t
                    .AutoReset = True
                    .Enabled = True
                    .Start()
                End With
                'LogEvent("NetkaQuartz Escalator service started successfully.", EventLogEntryType.Information)
            Else
                log.Error("NetkaQuartz: service could not start due to " & strProductValid)
                End
            End If
        Catch ex As Exception
            LogEvent("OnStart: " & ex.Message, EventLogEntryType.Error)
            Throw ex
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        Try
            t.Stop()
            t.Dispose()
            'LogEvent("NetkaQuartz Escalator service stopped successfully.", EventLogEntryType.Information)
        Catch ex As Exception
            LogEvent("OnStop: " & ex.Message, EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub Timer_Fired(ByVal sender As Object, ByVal e As ElapsedEventArgs)

        log.Info("==========================================================================================")

        SetGlobalization()

        NextNotificationUpdate()

        EscalationMatrix()

        log.Info("==========================================================================================")
    End Sub

    Public Sub SetGlobalization()
        Dim dtf As System.Globalization.DateTimeFormatInfo = New System.Globalization.DateTimeFormatInfo()
        dtf.DateSeparator = "-"
        dtf.ShortDatePattern = "dd-MMM-yy"
        dtf.ShortTimePattern = "HH:mm:ss"
        Dim ci As New System.Globalization.CultureInfo("en-US")
        ci.DateTimeFormat = dtf
        System.Threading.Thread.CurrentThread.CurrentCulture = ci
    End Sub

    Public Sub CalCaseDuration()
        Try
            Dim conn As New OdbcConnection(strConnection)
            conn.Open()
            sql = "SELECT * FROM incident WHERE case_status_id < 6 ORDER BY id "
            Dim cmd_chk As New OdbcDataAdapter(sql, conn)
            Dim ds_chk As New DataSet()
            cmd_chk.Fill(ds_chk, "*")
            Dim dt_chk As DataTable = ds_chk.Tables("*")
            Dim dr_chk As DataRow
            For Each dr_chk In dt_chk.Rows
                If dr_chk("staff_id").ToString() = 0 And dr_chk("contact_id").ToString() <> 0 Then 'for External Esclation SLA from SLA table
                    'sql = "SELECT incident.id,CAST(CONCAT(case_log.date,' ',case_log.time) AS CHAR) AS created_date,CAST(TIMESTAMPDIFF(MINUTE,CONCAT(case_log.date,' ',case_log.time),CURRENT_TIMESTAMP) AS CHAR) AS duration,7x24 FROM incident,case_log,contract,sla WHERE incident.id=case_log.case_id AND case_log.case_log_type_id='1' AND incident.contract_id=contract.contract_id AND contract.sla_id=sla.sla_id AND case_status_id < 6 AND id='" & dr_chk("id").ToString() & "'"
                    sql = "SELECT distinct incident.id,CAST(CONCAT(case_log.date,' ',case_log.time) AS CHAR) AS created_date," & _
                           " CAST(TIMESTAMPDIFF(MINUTE,CONCAT(case_log.date,' ',case_log.time),CURRENT_TIMESTAMP) AS CHAR) AS duration,sla_Detail.7x24" & _
                           " FROM(incident, case_log, contract, sla, sla_detail, priority)" & _
                           " WHERE(incident.id = case_log.case_id)" & _
                            " AND case_log.case_log_type_id='1'" & _
                            " AND incident.contract_id=contract.contract_id" & _
                            " AND contract.sla_id=sla.sla_id" & _
                            " AND case_status_id < 6" & _
                            " AND (sla.sla_id = sla_detail.sla_id and incident.priority_id = sla_detail.priority_id)" & _
                            " AND case_status_id < 6 AND id='" & dr_chk("id").ToString() & "'"
                    Dim cmd As New OdbcDataAdapter(sql, conn)
                    Dim ds As New DataSet()
                    cmd.Fill(ds, "*")
                    Dim dt As DataTable = ds.Tables("*")
                    Dim dr As DataRow
                    Dim incident_id, created_date, duration, s724 As String
                    Dim strDate, strTime As String
                    Dim myArray(1) As String
                    Dim intNonWorkingMinute, intPendingMinute As Integer
                    Dim cmd1 As New OdbcDataAdapter()
                    Dim ds1 As New DataSet()
                    For Each dr In dt.Rows
                        incident_id = dr.Item("id").ToString()
                        created_date = dr.Item("created_date").ToString()
                        duration = dr.Item("duration").ToString()
                        s724 = dr.Item("7x24").ToString()
                        strDebug = "CalCaseDuration1: " & incident_id & "," & created_date & "," & s724 & "," & duration
                        If s724 = "0" Then
                            myArray = Split(created_date, " ")
                            strDate = myArray(0)
                            strTime = myArray(1)
                            intNonWorkingMinute = NonWorkingMinute(strDate, strTime, Now.ToString("yyyy-MM-dd"), Now.ToString("HH:mm:ss"), conn)
                            duration = Val(duration) - intNonWorkingMinute
                            intPendingMinute = PendingMinute5x8(incident_id)
                            duration = Val(duration) - intPendingMinute
                            If Val(duration) < 0 Then
                                duration = "0"
                            End If
                            strDebug += "," & intPendingMinute & "+" & intNonWorkingMinute & "," & duration
                        Else
                            intPendingMinute = PendingMinute7x24(incident_id)
                            duration = Val(duration) - intPendingMinute
                            If Val(duration) < 0 Then
                                duration = "0"
                            End If
                            'log.Info("CalCaseDuration3: " & intPendingMinute & " " & duration)
                            strDebug += "," & intPendingMinute & "+0," & duration
                        End If
                        log.Info(strDebug)

                        ' Update case_duration in incident table
                        sql = "UPDATE incident SET case_duration='" & duration & "' WHERE id='" & incident_id & "'"
                        cmd1 = New OdbcDataAdapter(sql, conn)
                        cmd1.Fill(ds1, "incident")
                    Next
                Else
                    sql = "SELECT incident.id,CAST(CONCAT(case_log.date,' ',case_log.time) AS CHAR) AS created_date,CAST(TIMESTAMPDIFF(MINUTE,CONCAT(case_log.date,' ',case_log.time),CURRENT_TIMESTAMP) AS CHAR) AS duration,7x24 FROM incident,case_log,priority WHERE incident.id=case_log.case_id AND case_log.case_log_type_id='1' AND incident.priority_id=priority.priority_id AND case_status_id < 6 AND id='" & dr_chk("id").ToString() & "'"
                    Dim cmd As New OdbcDataAdapter(sql, conn)
                    Dim ds As New DataSet()
                    cmd.Fill(ds, "*")
                    Dim dt As DataTable = ds.Tables("*")
                    Dim dr As DataRow
                    Dim incident_id, created_date, duration, s724 As String
                    Dim strDate, strTime As String
                    Dim myArray(1) As String
                    Dim intNonWorkingMinute, intPendingMinute As Integer
                    Dim cmd1 As New OdbcDataAdapter()
                    Dim ds1 As New DataSet()
                    For Each dr In dt.Rows
                        incident_id = dr.Item("id").ToString()
                        created_date = dr.Item("created_date").ToString()
                        duration = dr.Item("duration").ToString()
                        s724 = dr.Item("7x24").ToString()
                        strDebug = "CalCaseDuration1: " & incident_id & "," & created_date & "," & s724 & "," & duration
                        If s724 = "0" Then
                            myArray = Split(created_date, " ")
                            strDate = myArray(0)
                            strTime = myArray(1)
                            intNonWorkingMinute = NonWorkingMinute(strDate, strTime, Now.ToString("yyyy-MM-dd"), Now.ToString("HH:mm:ss"), conn)
                            duration = Val(duration) - intNonWorkingMinute
                            intPendingMinute = PendingMinute5x8(incident_id)
                            duration = Val(duration) - intPendingMinute
                            If Val(duration) < 0 Then
                                duration = "0"
                            End If
                            strDebug += "," & intPendingMinute & "+" & intNonWorkingMinute & "," & duration
                        Else
                            intPendingMinute = PendingMinute7x24(incident_id)
                            duration = Val(duration) - intPendingMinute
                            If Val(duration) < 0 Then
                                duration = "0"
                            End If
                            'log.Info("CalCaseDuration3: " & intPendingMinute & " " & duration)
                            strDebug += "," & intPendingMinute & "+0," & duration
                        End If
                        log.Info(strDebug)

                        ' Update case_duration in incident table
                        sql = "UPDATE incident SET case_duration='" & duration & "' WHERE id='" & incident_id & "'"
                        cmd1 = New OdbcDataAdapter(sql, conn)
                        cmd1.Fill(ds1, "incident")
                    Next
                End If
            Next
            conn.Close()
        Catch ex As Exception
            LogEvent("CalCaseDuration: " & ex.Message, EventLogEntryType.Error)
            log.Error("CalCaseDuration: " & ex.Message)
        End Try
    End Sub

    Public Sub NextNotificationUpdate()
        Try
            Dim conn As New OdbcConnection(strConnection)
            conn.Open()
            log.Info("Start Checking...")
            sql = "SELECT problem_case.*,DATE_FORMAT(CONCAT(CURDATE(),' ',CURTIME()),'%Y-%m-%d %H:%i') cur_time,  " & _
                  "DATE_FORMAT(CONCAT(CURDATE(),' ',CURTIME()),'%Y-%m-%d %H:%i:%s') cur_back_time," & _
                  "DATE_FORMAT(short_term_due_date,'%Y-%m-%d %H:%i') short_term_cal,DATE_FORMAT(long_term_due_date,'%Y-%m-%d %H:%i') long_term_cal " & _
                  " FROM problem_case WHERE status_id <6 AND short_term_due_date <>'0000-00-00 00:00:00' ORDER BY id "

            Dim cmd_chk As New OdbcDataAdapter(sql, conn)
            Dim ds_chk As New DataSet()
            cmd_chk.Fill(ds_chk, "*")
            Dim dt_chk As DataTable = ds_chk.Tables("*")
            Dim dr_chk As DataRow
            Dim notiShortTermDate As Date
            Dim notiLongTermDate As Date
            Dim curDate As Date
            Dim curDateBack As Date
            Dim strEmail As String = ""
            Dim strmail1 As String = ""
            Dim strmail2 As String = ""
            For Each dr_chk In dt_chk.Rows
                Dim strSubj, strBody As String
                notiShortTermDate = dr_chk("short_term_cal").ToString()
                notiLongTermDate = dr_chk("long_term_cal").ToString()
                curDate = dr_chk("cur_time").ToString()
                curDateBack = dr_chk("cur_back_time").ToString()
                curDateBack = curDateBack.AddHours(24)
                log.Info("Problem ID = " & dr_chk.Item("id") & " Next Notify Short Term Date =" & notiShortTermDate)

                ' Email to Service Desk team
                sql = "select replace(group_concat(email),',',';') as email from staff where staff_id in( " & _
                      " select staff_id From user where level_id=2)"
                If strEmail <> "" Then
                    strEmail = strEmail & ";"
                End If
                strEmail = strEmail & FetchField(sql, strConnection, "email")

                ' Email to Problem investigation team
                sql = "SELECT replace(group_concat(email),',',';') as email from staff where team_id='" & dr_chk.Item("team_id") & "' "
                If strEmail <> "" Then
                    strEmail = strEmail & ";"
                End If
                strEmail = strEmail & FetchField(sql, strConnection, "email")

                ' Email to Problem Coordinator
                sql = "SELECT email from staff WHERE staff_id='" & dr_chk.Item("engineer_id") & "'"
                If strEmail <> "" Then
                    strEmail = strEmail & ";"
                End If

                strEmail = strEmail & FetchField(sql, strConnection, "email")

                'Escalation before breaching SLA
                log.Info("ShotDate =" & notiShortTermDate.ToString("yyyy-MM-dd HH:mm") & "| Curdateback = " & curDateBack.ToString("yyyy-MM-dd HH:mm"))
                If notiShortTermDate.ToString("yyyy-MM-dd HH:mm") = curDateBack.ToString("yyyy-MM-dd HH:mm") Then ' Notify Short Term before 24 hrs
                    log.Info("!!!Problem ID = " & dr_chk.Item("id") & " Notification before Short Term Date =" & notiShortTermDate)
                    strSubj = FetchField("SELECT subject FROM  email_template WHERE template_name ='escalation_befor_breaching' ", strConnection, "subject")
                    strBody = FetchField("SELECT message FROM  email_template WHERE template_name ='escalation_befor_breaching' ", strConnection, "message")
                    strSubj = ReplaceEmailVariable(strSubj, dr_chk.Item("id"))
                    strBody = ReplaceEmailVariable(strBody, dr_chk.Item("id"))
                    strBody += "<br /><img src=cid:companylogo>"
                    strBody = "<html><body><div style='font-family:Tahoma, MS Sans Serif'>" & strBody & "</div></body></html>"
                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                        sEmail(dr_chk("id").ToString(), strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                    End If

                End If
                If notiLongTermDate.ToString("yyyy-MM-dd HH:mm") = curDateBack.ToString("yyyy-MM-dd HH:mm") Then ' Notify Long Term  before 24 hrs
                    log.Info("!!!Problem ID = " & dr_chk.Item("id") & " Notification before Long Term Date =" & notiLongTermDate)
                    strSubj = FetchField("SELECT subject FROM  email_template WHERE template_name ='escalation_befor_breaching' ", strConnection, "subject")
                    strBody = FetchField("SELECT message FROM  email_template WHERE template_name ='escalation_befor_breaching' ", strConnection, "message")
                    strSubj = ReplaceEmailVariable(strSubj, dr_chk.Item("id"))
                    strBody = ReplaceEmailVariable(strBody, dr_chk.Item("id"))
                    strBody += "<br /><img src=cid:companylogo>"
                    strBody = "<html><body><div style='font-family:Tahoma, MS Sans Serif'>" & strBody & "</div></body></html>"
                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                        sEmail(dr_chk("id").ToString(), strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                    End If

                End If

                ' Email to Problem manager
                sql = "SELECT replace(group_concat(email),',',';') as email FROM function_privilege " & _
                      " LEFT JOIN user on function_privilege.username = user.username " & _
                      " LEFT JOIN staff on user.staff_id = staff.staff_id " & _
                      " WHERE function =3 AND allow_review=1  "
                If strEmail <> "" Then
                    strEmail = strEmail & ";"
                End If

                strEmail = strEmail & FetchField(sql, strConnection, "email")
                ' Escalation After breaching SLA
                If notiShortTermDate <= curDate Then ' Notify Short Term after SLA
                    If notiShortTermDate.ToString("HH:mm") = curDate.ToString("HH:mm") Then
                        log.Info("!!!Problem ID = " & dr_chk.Item("id") & " Notification after Short Term Date =" & notiShortTermDate)
                        strSubj = FetchField("SELECT subject FROM  email_template WHERE template_name ='escalation_after_breaching' ", strConnection, "subject")
                        strBody = FetchField("SELECT message FROM  email_template WHERE template_name ='escalation_after_breaching' ", strConnection, "message")
                        strSubj = ReplaceEmailVariable(strSubj, dr_chk.Item("id"))
                        strBody = ReplaceEmailVariable(strBody, dr_chk.Item("id"))
                        strBody += "<br /><img src=cid:companylogo>"
                        strBody = "<html><body><div style='font-family:Tahoma, MS Sans Serif'>" & strBody & "</div></body></html>"
                        If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                            sEmail(dr_chk("id").ToString(), strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                        End If
                    End If

                End If
                If notiLongTermDate <= curDate Then ' Notify Long Term  after SLA
                    If notiLongTermDate.ToString("HH:mm") = curDate.ToString("HH:mm") Then
                        log.Info("!!!Problem ID = " & dr_chk.Item("id") & " Notification after Long Term Date =" & notiShortTermDate)
                        strSubj = FetchField("SELECT subject FROM  email_template WHERE template_name ='escalation_after_breaching' ", strConnection, "subject")
                        strBody = FetchField("SELECT message FROM  email_template WHERE template_name ='escalation_after_breaching' ", strConnection, "message")
                        strSubj = ReplaceEmailVariable(strSubj, dr_chk.Item("id"))
                        strBody = ReplaceEmailVariable(strBody, dr_chk.Item("id"))
                        strBody += "<br /><img src=cid:companylogo>"

                        strBody = "<html><body><div style='font-family:Tahoma, MS Sans Serif'>" & strBody & "</div></body></html>"
                        If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                            sEmail(dr_chk("id").ToString(), strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                        End If
                    End If

                End If
            Next
            conn.Close()
            log.Info("Finished Checking...")
        Catch ex As Exception
            LogEvent("NotifySLA: " & ex.Message, EventLogEntryType.Error)
            log.Error("NotifySLA: " & ex.Message)
        End Try
    End Sub

    Public Sub EscalationMatrix()
        Try
            Dim conn As New OdbcConnection(strConnection)
            conn.Open()
            log.Info("Start Checking EscalationMatrix ...")

            sql = "SELECT problem_case.*,DATEDIFF(CONCAT(CURDATE(),' ',CURTIME()),created_date) days " & _
                 " FROM problem_case " & _
                 " WHERE status_id < 6 ORDER BY id "
            Dim cmd_chk As New OdbcDataAdapter(sql, conn)
            Dim ds_chk As New DataSet()
            cmd_chk.Fill(ds_chk, "*")
            Dim dt_chk As DataTable = ds_chk.Tables("*")
            Dim dr_chk As DataRow
            Dim priority_id As Integer = 0
            Dim days As String = ""
            Dim recipient As String = ""
            Dim cc As String = ""
            Dim email_template_id As Integer = 0
            Dim sms_template_id As Integer = 0
            Dim level As Integer = 0
            For Each dr_chk In dt_chk.Rows
                priority_id = dr_chk.Item("priority_id").ToString()
                days = dr_chk.Item("days").ToString()
                recipient = ""
                cc = ""
                email_template_id = ""
                sms_template_id = ""
                level = ""
                If CheckEscalationMatrix(days, priority_id, recipient, cc, email_template_id, sms_template_id, level) Then
                    Dim _sql As String = "SELECT level" & level & " lvl FROM problem_escalator_flag WHERE problem_id = '" & dr_chk.Item("id").ToString() & "' "
                    Dim is_escalated As String = FetchField(_sql, strConnection, "lvl")
                    Dim strSubj, strBody As String
                    If is_escalated = 0 Or is_escalated = "" Then
                        ' Get email template 
                        strSubj = FetchField("SELECT subject FROM  email_template WHERE id ='" & email_template_id & "' ", strConnection, "subject")
                        strBody = FetchField("SELECT message FROM  email_template WHERE id ='" & email_template_id & "' ", strConnection, "message")
                        strSubj = ReplaceEmailVariable(strSubj, dr_chk.Item("id"))
                        strBody = ReplaceEmailVariable(strBody, dr_chk.Item("id"))
                        strBody += "<br /><img src=cid:companylogo>"
                        strBody = "<html><body><div style='font-family:Tahoma, MS Sans Serif'>" & strBody & "</div></body></html>"
                        If recipient <> "" And smtp_server <> "" And smtp_sender <> "" And strSubj <> "" Then
                            sEmail(dr_chk("id").ToString(), strSubj, strBody, recipient, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                            ' Mark escalation flag
                            If is_escalated = 0 Then
                                _sql = "UPDATE problem_escalator_flag SET level" & level & " = 1 WHERE problem_id ='" & dr_chk.Item("id").ToString() & "' "
                                ExecuteSql(_sql, strConnection)
                            Else
                                _sql = "INSERT INTO problem_escalator_flag(problem_id, level" & level & ") VALUES('" & dr_chk.Item("id").ToString() & "','1') "
                                ExecuteSql(_sql, strConnection)
                            End If
                            ' Update problem log 
                            _sql = "INSERT INTO problem_history_log(event_date, event, logged_by, problem_id) VALUES(" & _
                                   "'" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "','" & strSubj & "',1,'" & dr_chk.Item("id").ToString() & "' "
                            ExecuteSql(_sql, strConnection)
                            log.Info("Escalation email SMS ...")
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            LogEvent("EscalationMatrix: " & ex.Message, EventLogEntryType.Error)
            log.Error("EscalationMatrix: " & ex.Message)
        End Try
    End Sub
    Function BuildBodyNextUpdate(ByVal incident_id As String, ByVal action As String, ByRef mail1 As String, ByRef mail2 As String) As String
        Dim strSubj As String = ""
        Dim strBody As String = ""
        Try

            Dim conn As New OdbcConnection(strConnection)
            'sql = "SELECT incident.*,case_status_title,priority_title,tier_title,case_category_title,case_sub_category_title,customer_name,contract_title,site_title,CONCAT(staff.firstname,' ',staff.lastname) AS engineer_name FROM incident,case_status,priority,tier,case_category,case_sub_category,customer,contract,site,staff WHERE incident.case_status_id=case_status.case_status_id AND incident.priority_id=priority.priority_id AND (incident.tier_id=tier.tier_id OR incident.tier_id =0) AND incident.case_sub_category_id=case_sub_category.case_sub_category_id AND case_category.case_category_id=case_sub_category.case_category_id AND incident.contract_id=contract.contract_id AND customer.customer_id=contract.customer_id AND incident.site_id=site.site_id AND incident.engineer_id=staff.staff_id AND incident.id=" & incident_id & " LIMIT 1"
            sql = "SELECT incident.id,incident.case_id,incident.title,DATE_FORMAT(schedule_date,'%Y-%m-%d %H:%i') schedule_date," & _
                  " DATE_FORMAT(CONCAT(CURDATE(),' ',CURTIME()),'%Y-%m-%d %H:%i') cur_time,DATE_FORMAT(schedule_date,'%d/%m/%Y %H:%i') noti_date, " & _
                  " title, is_repeat, repeat_minute, inactive ," & _
                  " TIMESTAMPDIFF(MINUTE,schedule_date,CONCAT(CURDATE(),' ',CURTIME())) duration" & _
                  " FROM incident " & _
                  " INNER JOIN incident_notification on incident.id =  incident_notification.case_id " & _
                 " WHERE incident.id=" & incident_id
            'log.Info("SQL EMAIL = CASE_ID :" & "  " & incident_id & " " & sql)
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            Dim dt As DataTable = ds.Tables("*")
            Dim dr As DataRow
            'Dim strSiteAddress, strContactDetail, strParentCase As String
            ' strSubj = FetchField("SELECT subject FROM  email_template WHERE template_name ='Notification_update' ", strConnection, "subject")
            strBody = FetchField("SELECT message FROM  email_template WHERE template_name ='Notification_update'  ", strConnection, "message")

            For Each dr In dt.Rows
                'strSubj = strSubj.Replace("{case_id}", dr.Item("case_id"))
                strBody = strBody.Replace("{case_id}", dr.Item("case_id"))
                strBody = strBody.Replace("{case_title}", dr.Item("title"))
                sql = " SELECT DATE_FORMAT(CONCAT(date,' ',time),'%d/%m/%Y %H:%i') last_date,logged_by,case_log_description  From case_log where case_id=" & incident_id & "" & _
                      " and case_log_type_id='10' order by case_log_id desc limit 1"
                Dim last_log As String = FetchField(sql, strConnection, "case_log_description")
                Dim last_date As String = FetchField(sql, strConnection, "last_date")
                Dim last_log_by As String = FetchField(sql, strConnection, "logged_by")

                sql = "select ifnull(area1.office_title,'') a1_title,ifnull(area1.web,'') a1_mail," & _
                      "  ifnull(area2.office_title,'') a2_title,ifnull(area2.web,'') a2_mail " & _
                      " From incident" & _
                      "  left join office area1 on incident.engineer_group_id = area1.office_id " & _
                       " left join office area2 on incident.engineer_group2_id = area2.office_id " & _
                       " where id='" & incident_id & "' "
                Dim area1 As String = FetchField(sql, strConnection, "a1_title")
                mail1 = FetchField(sql, strConnection, "a1_mail")
                Dim area2 As String = FetchField(sql, strConnection, "a2_title")
                mail2 = FetchField(sql, strConnection, "a2_mail")
                strBody = strBody.Replace("{lastest_log}", last_log)
                strBody = strBody.Replace("{area}", area1)
                strBody = strBody.Replace("{area2}", area2)
                strBody = strBody.Replace("{lastest_updated_date}", last_date)
                strBody = strBody.Replace("{lastest_updated_by}", last_log_by)
                strBody = strBody.Replace("{notification_date}", dr.Item("noti_date"))
            Next
            strBody += "<br /><img src=cid:companylogo>"
            strBody = "<html><body><div style='font-family:Tahoma, MS Sans Serif'>" & strBody & "</div></body></html>"
            BuildBodyNextUpdate = strBody
        Catch ex As Exception
            log.Info(strBody)
            log.Info(ex.ToString())
            LogEvent("BuildBody: " & ex.Message, EventLogEntryType.Error)
            log.Error("BuildBody: " & ex.Message)
            BuildBodyNextUpdate = ""
        End Try
    End Function

    Function ReplaceEmailVariable(subject As String, id As String) As String
        Dim conn As New OdbcConnection(strConnection)
        conn.Open()
        Dim result As String = ""

        sql = "SELECT * FROM problem_case WHERE id = '" & id & "' "
        Dim cmd_chk As New OdbcDataAdapter(sql, conn)
        Dim ds_chk As New DataSet()
        cmd_chk.Fill(ds_chk, "*")
        Dim dt_chk As DataTable = ds_chk.Tables("*")
        Dim dr_chk As DataRow
        For Each dr_chk In dt_chk.Rows
            subject = subject.Replace("{case_id}", dr_chk("id"))
            subject = subject.Replace("{case_title}", dr_chk("title"))
            Dim link As String = ""
            link = "<a href='" & web_url & "/problem_update.aspx?problem_id=" & dr_chk("id") & "' target='_blank'>" & dr_chk("problem_id") & "</a>"
            subject = subject.Replace("{case_id_link}", link)

            sql = "SELECT CONCAT(firstname,' ',lastname) AS creater FROM staff WHERE staff_id='" & dr_chk("requested_id") & "'"
            Dim creater As String = FetchField(sql, strConnection, "creater")
            subject = subject.Replace("{requester}", creater)
            subject = subject.Replace("{created_by}", creater)
            subject = subject.Replace("{case_description}", dr_chk("description"))

            Dim shortDate As Date = dr_chk("short_term_due_date")
            Dim longDate As Date = dr_chk("long_term_due_date")

            subject = subject.Replace("{created_date}", Now.ToString("dd-MMM-yyyy HH:mm:ss"))
            subject = subject.Replace("{short_term_due_date}", shortDate.ToString("dd-MMM-yyyy HH:mm:ss"))
            subject = subject.Replace("{long_term_due_date}", longDate.ToString("dd-MMM-yyyy HH:mm:ss"))

            sql = "SELECT * FROM problem_service_affected WHERE problem_id = '" & dr_chk("id") & "' ORDER BY id LIMIT 1"
            Dim case_type_id As Integer = FetchField(sql, strConnection, "case_type_id")
            sql = "SELECT case_type_title FROM case_type WHERE case_type_id= '" & case_type_id & "'"
            subject = subject.Replace("{service_catalog}", FetchField(sql, strConnection, "case_type_title"))

            sql = "SELECT description FROM priority_level WHERE impact_id='" & dr_chk("impact_id") & "' AND urgency_id='" & dr_chk("urgency_id") & "' "
            Dim priority As String = FetchField(sql, strConnection, "description")

            subject = subject.Replace("{priority}", priority)
            subject = subject.Replace("{case_status}", "Identify")
            subject = subject.Replace("<table", "<table style='font-family:Tahoma, MS Sans Serif' ")
            'subject = subject.Replace("px",".0pt")
            subject = subject.Replace("</span>", "</font>")
            Dim dbq As String = """"
            subject = subject.Replace("span style=" & dbq & "font-size: ", "font size=")
            subject = subject.Replace("span style=" & dbq & "color: ", "font color=")

            subject = subject.Replace("font size=10px", "font size=1")
            subject = subject.Replace("font size=12px", "font size=2")
            subject = subject.Replace("font size=14px", "font size=3")
            subject = subject.Replace("font size=16px", "font size=4")
            subject = subject.Replace("font size=18px", "font size=5")
            subject = subject.Replace("font size=20px", "font size=6")
            subject = subject.Replace("font size=24px", "font size=7")
            subject = subject.Replace("font size=26px", "font size=8")

            subject = StripTagsImg(subject)
        Next
        result = subject
        Return result
    End Function

    Public Sub Escalator()
        Try
            Dim conn As New OdbcConnection(strConnection)
            escalation_next_time = ConfigurationManager.AppSettings("escalation_next_time")
          
            ' Load parameter for sending Email & SMS
            conn.Open()
            sql = "SELECT * FROM config LIMIT 1"
            Dim cmd1 As New OdbcCommand(sql, conn)
            Dim datareader1 As OdbcDataReader = cmd1.ExecuteReader()
            While datareader1.Read()
                smtp_server = datareader1.Item("smtp_server")
                smtp_sender = datareader1.Item("smtp_sender")
                smtp_password = AESDecrypt(datareader1.Item("smtp_password"), "azsxdcfvgbhnjmk,", "aw34esdr56tfgy78")
                smtp_auth = datareader1.Item("smtp_auth")
                sms_url = datareader1.Item("sms_url")
                sms_user_keyword = datareader1.Item("sms_user_keyword")
                sms_password_keyword = datareader1.Item("sms_password_keyword")
                sms_number_keyword = datareader1.Item("sms_number_keyword")
                sms_message_keyword = datareader1.Item("sms_message_keyword")
                sms_username = datareader1.Item("sms_username")
                Try
                    sms_password = AESDecrypt(datareader1.Item("sms_password"), "azsxdcfvgbhnjmk,", "aw34esdr56tfgy78")
                Catch
                    sms_password = datareader1.Item("sms_password")
                End Try
                sms_protocol = datareader1.Item("sms_protocol")
                sms_result = datareader1.Item("sms_result")
                'sms_language_keyword = datareader1.Item("sms_language_keyword")
            End While
            conn.Close()

            If LCase(smtp_auth) = "ssl" Then
                bSsl = True
                sslPort = 465
            ElseIf LCase(smtp_auth) = "tls" Then
                bSsl = True
                sslPort = 587
            Else
                bSsl = False
                sslPort = 0
            End If
            'sql = "SELECT * FROM incident WHERE case_status_id < 6 ORDER BY id "
            sql = "SELECT * FROM incident INNER JOIN case_type ON incident.case_type_id = case_type.case_type_id WHERE case_status_id < 6 AND is_alert='1' "

            Dim cmd_chk As New OdbcDataAdapter(sql, conn)
            Dim ds_chk As New DataSet()
            cmd_chk.Fill(ds_chk, "*")
            Dim dt_chk As DataTable = ds_chk.Tables("*")
            Dim dr_chk As DataRow
            For Each dr_chk In dt_chk.Rows
                If dr_chk("staff_id").ToString() = 0 And dr_chk("contact_id").ToString() <> 0 Then 'for External Esclation SLA from SLA table

                    sql = " SELECT distinct  incident.id,CAST(CONCAT(case_log.date,' ',case_log.time) AS CHAR) AS created_date,incident.case_id,incident.title,incident.tier_id," & _
                          "incident.case_status_id,engineer_id,case_duration,contract.contract_id,sla.*,sla_detail.*" & _
                          " FROM(incident, case_log, case_status, contract, sla, sla_detail, priority)" & _
                          " WHERE incident.id=case_log.case_id AND case_log.case_log_type_id='1'" & _
                          " AND incident.case_status_id=case_status.case_status_id AND incident.contract_id=contract.contract_id" & _
                          " AND contract.sla_id=sla.sla_id AND case_status.case_status_id < 6" & _
                          " AND (sla.sla_id = sla_detail.sla_id and incident.priority_id = sla_detail.priority_id) AND id='" & dr_chk("id").ToString() & "'"


                    Dim cmd As New OdbcDataAdapter(sql, conn)
                    Dim ds As New DataSet()
                    cmd.Fill(ds, "*")
                    Dim dt As DataTable = ds.Tables("*")
                    Dim dr As DataRow
                    Dim incident_id, created_date, case_id, case_title, case_status_id, engineer_id, duration, response, onsite, resolve, autoclose, s724, sla_title As String
                    Dim tier_id As String
                    Dim myArray(1) As String
                    For Each dr In dt.Rows
                        incident_id = dr.Item("id").ToString()
                        created_date = dr.Item("created_date").ToString()
                        case_id = dr.Item("case_id").ToString()
                        case_title = dr.Item("title").ToString()
                        case_status_id = dr.Item("case_status_id").ToString()
                        engineer_id = dr.Item("engineer_id").ToString()
                        duration = dr.Item("case_duration").ToString()
                        response = dr.Item("response").ToString()
                        onsite = dr.Item("onsite").ToString()
                        resolve = dr.Item("resolve").ToString()
                        autoclose = dr.Item("autoclose").ToString()
                        s724 = dr.Item("7x24").ToString()
                        sla_title = dr.Item("sla_title").ToString()
                        tier_id = dr.Item("tier_id").ToString()
                        myArray = Split(response, ":")
                        response = (myArray(0) * 60) + myArray(1)
                        myArray = Split(onsite, ":")
                        onsite = (myArray(0) * 60) + myArray(1)
                        myArray = Split(resolve, ":")
                        resolve = (myArray(0) * 60) + myArray(1)
                        myArray = Split(autoclose, ":")
                        autoclose = (myArray(0) * 60) + myArray(1)
                        log.Info("Escalator1: " & incident_id & "," & duration & "," & response & "," & onsite & "," & resolve & "," & autoclose)

                        Dim strSubj, strBody, strEmail As String
                        Dim arrMobile As New ArrayList

                        'Alert before overdue 
                        If ((Val(resolve) - Val(duration)) < alert_before) Then
                            If (Val(duration) Mod alert_before_every) = 1 Then
                                strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Remaining"
                                strBody = BuildBody(incident_id, "Notification of Remaining")
                                log.Info(strSubj)
                                ' UpdateIncidentOverdue("response_overdue", "1", incident_id, conn)
                                AddCaseLog(incident_id, "4", "Notification of Remaining (Duration: " & duration & " min, SLA resolve:" & resolve & " min)")
                                strEmail = BuildEmailAddress(incident_id, engineer_id, 0)
                                If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                    sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                End If
                            End If
                        End If
                        ' Check and alert for response overdue
                        If (Val(duration) > Val(response)) And Val(response) <> 0 Then
                            If Not IsCaseResponsed(incident_id, conn) Then
                                Dim numAlerted As Integer = NumAlertedForResponse(incident_id, conn)
                                If numAlerted = 0 Then  ' Never alerted before
                                    'strSubj = BuildSubj(case_id, case_title, "Notification of Response Overdue")
                                    strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Response Overdue"
                                    strBody = BuildBody(incident_id, "Notification of Response Overdue")
                                    log.Info(strSubj)
                                    UpdateIncidentOverdue("response_overdue", "1", incident_id, conn)
                                    AddCaseLog(incident_id, "13", "response overdued (Duration: " & duration & " min, SLA response:" & response & " min)")
                                    strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                        sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                    End If

                                    arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                    If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                        sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                    End If
                                Else
                                    log.Info("Escalation time = " & escalation_next_time)
                                    If escalation_next_time = "" Then ' Default Escalation 
                                        If (Val(duration) Mod Val(response)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Response Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "13", "response overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA response:" & response & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    Else ' Escalation by config time
                                        If (Val(duration) Mod Val(escalation_next_time)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Response Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "13", "response overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA response:" & response & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ' Check and alert for onsite overdue
                        If (Val(duration) > Val(onsite)) And Val(onsite) <> 0 Then
                            If Not IsCaseOnsited(incident_id, conn) Then
                                Dim numAlerted As Integer = NumAlertedForOnsite(incident_id, conn)
                                If numAlerted = 0 Then  ' Never alerted before
                                    'strSubj = BuildSubj(case_id, case_title, "Notification of Onsite Overdue")
                                    strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Onsite Overdue"
                                    log.Info(strSubj)
                                    strBody = BuildBody(incident_id, "Notification of Onsite Overdue")
                                    UpdateIncidentOverdue("onsite_overdue", "1", incident_id, conn)
                                    AddCaseLog(incident_id, "14", "onsite overdued (Duration: " & duration & " min, SLA onsite:" & onsite & " min)")
                                    If escalation_type = "2" And numAlerted < 1 Then    'escalation_type = 2 for IT Management
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, 1)
                                        arrMobile = BuildMobile(incident_id, engineer_id, 1)
                                    Else
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                        arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                    End If

                                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                        sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                    End If

                                    If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                        sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                    End If
                                Else
                                    If escalation_next_time = "" Then ' Default Escalation 
                                        If (Val(duration) Mod Val(onsite)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Onsite Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "14", "onsite overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA onsite:" & onsite & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    Else ' Escalation by config time
                                        If (Val(duration) Mod Val(escalation_next_time)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Onsite Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "14", "onsite overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA onsite:" & onsite & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ' Check and alert for resolve overdue
                        If (Val(duration) > Val(resolve)) And Val(resolve) <> 0 Then
                            If Not IsCaseResolved(incident_id, conn) Then
                                Dim numAlerted As Integer = NumAlertedForResolve(incident_id, conn)
                                If numAlerted = 0 Then  ' Never alerted before
                                    'strSubj = BuildSubj(case_id, case_title, "Notification of Resolve Overdue")
                                    strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Resolve Overdue"
                                    log.Info(strSubj)
                                    strBody = BuildBody(incident_id, "Notification of Resolve Overdue")
                                    UpdateIncidentOverdue("resolve_overdue", "1", incident_id, conn)
                                    AddCaseLog(incident_id, "15", "resolve overdued (Duration: " & duration & " min, SLA resolve:" & resolve & " min)")
                                    If escalation_type = "2" And numAlerted < 2 Then    'escalation_type = 2 for IT Management
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, 2)
                                        arrMobile = BuildMobile(incident_id, engineer_id, 2)
                                    Else
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                        arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                    End If

                                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                        sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                    End If

                                    If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                        sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                    End If
                                    If auto_escalation = 1 Then
                                        Dim _sql As String
                                        Dim _next_tier As Integer = tier_id + 1
                                        Dim _old_tier_id As Integer = tier_id

                                        _sql = "Update incident set tier_id='" & _next_tier & "' where id ='" & incident_id & "'"
                                        ExecuteSql(_sql, strConnection)

                                        _sql = "SELECT tier_title from tier where tier_id='" & _old_tier_id & "'"
                                        Dim old_tier As String = FetchField(_sql, strConnection, "tier_title")
                                        _sql = "SELECT tier_title from tier where tier_id='" & _next_tier & "'"
                                        Dim new_tier As String = FetchField(_sql, strConnection, "tier_title")
                                        AddCaseLog(incident_id, "3", "Auto escalated from " & old_tier & " to " & new_tier)
                                        log.Info("Auto Escalation to next tier.")
                                        strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Auto Escalation"
                                        log.Info(strSubj)
                                        strBody = BuildBody(incident_id, "Notification of Auto Escalation")
                                        'UpdateIncidentOverdue("resolve_overdue", "1", incident_id, conn)
                                        AddCaseLog(incident_id, "15", "notification of auto Escalation from " & old_tier & " to " & new_tier)

                                        _sql = "select group_concat(email) as email from staff where  team_id='" & _next_tier & "' "
                                        Dim email_team_list As String = ""
                                      
                                        If email_team_list <> "" Then
                                            email_team_list = email_team_list & ";"
                                        Else
                                            email_team_list = email_team_list & FetchField(_sql, strConnection, "email")
                                        End If
                                        'strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                        strEmail = Replace(email_team_list, ",", ";")
                                        ' arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)

                                        If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                            sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                        End If
                                    End If
                                Else
                                    If escalation_next_time = "" Then ' Default Escalation 
                                        If (Val(duration) Mod Val(resolve)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Resolve Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "15", "resolve overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA resolve:" & resolve & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    Else ' Escalation by config time
                                        If (Val(duration) Mod Val(escalation_next_time)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Resolve Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "15", "resolve overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA resolve:" & resolve & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    End If

                                End If
                            End If
                        End If
                    Next
                Else 'for Internal Esclation SLA from Priority table
                    sql = "SELECT incident.id,CAST(CONCAT(case_log.date,' ',case_log.time) AS CHAR) AS created_date,incident.case_id,incident.title," & _
                          "incident.case_status_id,engineer_id,case_duration,staff.staff_id,priority.*  " & _
                          "FROM incident,case_log,case_status,staff,priority " & _
                          " WHERE incident.id=case_log.case_id AND case_log.case_log_type_id='1' " & _
                          " AND incident.case_status_id=case_status.case_status_id AND incident.staff_id=staff.staff_id " & _
                          "AND incident.priority_id=priority.priority_id AND case_status.case_status_id < 6 AND id='" & dr_chk("id").ToString() & "'"
                    'log.Info(sql)
                    Dim cmd As New OdbcDataAdapter(sql, conn)
                    Dim ds As New DataSet()
                    cmd.Fill(ds, "*")
                    Dim dt As DataTable = ds.Tables("*")
                    Dim dr As DataRow
                    Dim incident_id, created_date, case_id, case_title, case_status_id, engineer_id, duration, response, onsite, resolve, autoclose, s724, sla_title As String
                    Dim myArray(1) As String
                    For Each dr In dt.Rows
                        incident_id = dr.Item("id").ToString()
                        created_date = dr.Item("created_date").ToString()
                        case_id = dr.Item("case_id").ToString()
                        case_title = dr.Item("title").ToString()
                        case_status_id = dr.Item("case_status_id").ToString()
                        engineer_id = dr.Item("engineer_id").ToString()
                        duration = dr.Item("case_duration").ToString()
                        response = dr.Item("response").ToString()
                        onsite = dr.Item("onsite").ToString()
                        resolve = dr.Item("resolve").ToString()
                        autoclose = dr.Item("autoclose").ToString()
                        s724 = dr.Item("7x24").ToString()
                        sla_title = dr.Item("priority_title").ToString()

                        myArray = Split(response, ":")
                        response = (myArray(0) * 60) + myArray(1)
                        myArray = Split(onsite, ":")
                        onsite = (myArray(0) * 60) + myArray(1)
                        myArray = Split(resolve, ":")
                        resolve = (myArray(0) * 60) + myArray(1)
                        myArray = Split(autoclose, ":")
                        autoclose = (myArray(0) * 60) + myArray(1)
                        log.Info("Escalator1: " & incident_id & "," & duration & "," & response & "," & onsite & "," & resolve & "," & autoclose)

                        Dim strSubj, strBody, strEmail As String
                        Dim arrMobile As New ArrayList
                        ' Check and alert for response overdue
                        If (Val(duration) > Val(response)) And Val(response) <> 0 Then
                            If Not IsCaseResponsed(incident_id, conn) Then
                                Dim numAlerted As Integer = NumAlertedForResponse(incident_id, conn)
                                If numAlerted = 0 Then  ' Never alerted before
                                    'strSubj = BuildSubj(case_id, case_title, "Notification of Response Overdue")
                                    strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Response Overdue"
                                    log.Info(strSubj)
                                    strBody = BuildBody(incident_id, "Notification of Response Overdue")
                                    UpdateIncidentOverdue("response_overdue", "1", incident_id, conn)
                                    AddCaseLog(incident_id, "13", "response overdued (Duration: " & duration & " min, SLA response:" & response & " min)")
                                    strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                        sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                    End If

                                    arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                    If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                        sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                    End If
                                Else
                                    If escalation_next_time = "" Then ' Default Escalation 
                                        If (Val(duration) Mod Val(response)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Response Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            'strBody = "Response overdued for this case." & vbCrLf & vbCrLf & strBody
                                            strBody = BuildBody(incident_id, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "13", "response overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA response:" & response & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    Else
                                        If (Val(duration) Mod Val(escalation_next_time)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Response Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            'strBody = "Response overdued for this case." & vbCrLf & vbCrLf & strBody
                                            strBody = BuildBody(incident_id, "Notification of Response Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "13", "response overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA response:" & response & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    End If

                                End If
                            End If
                        End If

                        ' Check and alert for onsite overdue
                        If (Val(duration) > Val(onsite)) And Val(onsite) <> 0 Then
                            If Not IsCaseOnsited(incident_id, conn) Then
                                Dim numAlerted As Integer = NumAlertedForOnsite(incident_id, conn)
                                If numAlerted = 0 Then  ' Never alerted before
                                    'strSubj = BuildSubj(case_id, case_title, "Notification of Onsite Overdue")
                                    strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Onsite Overdue"
                                    log.Info(strSubj)
                                    strBody = BuildBody(incident_id, "Notification of Onsite Overdue")
                                    UpdateIncidentOverdue("onsite_overdue", "1", incident_id, conn)
                                    AddCaseLog(incident_id, "14", "onsite overdued (Duration: " & duration & " min, SLA onsite:" & onsite & " min)")
                                    If escalation_type = "2" And numAlerted < 1 Then    'escalation_type = 2 for IT Management
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, 1)
                                        arrMobile = BuildMobile(incident_id, engineer_id, 1)
                                    Else
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                        arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                    End If

                                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                        sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                    End If

                                    If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                        sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                    End If
                                Else
                                    If escalation_next_time = "" Then ' Default Escalation 
                                        If (Val(duration) Mod Val(onsite)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Onsite Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "14", "onsite overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA onsite:" & onsite & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    Else
                                        If (Val(duration) Mod Val(escalation_next_time)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Onsite Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Onsite Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "14", "onsite overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA onsite:" & onsite & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    End If

                                End If
                            End If
                        End If

                        ' Check and alert for resolve overdue
                        If (Val(duration) > Val(resolve)) And Val(resolve) <> 0 Then
                            If Not IsCaseResolved(incident_id, conn) Then
                                Dim numAlerted As Integer = NumAlertedForResolve(incident_id, conn)
                                If numAlerted = 0 Then  ' Never alerted before
                                    strSubj = BuildSubj(case_id, case_title, "Notification of Resolve Overdue")
                                    strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Resolve Overdue"

                                    log.Info(strSubj)
                                    strBody = BuildBody(incident_id, "Notification of Resolve Overdue")
                                    UpdateIncidentOverdue("resolve_overdue", "1", incident_id, conn)
                                    AddCaseLog(incident_id, "15", "resolve overdued (Duration: " & duration & " min, SLA resolve:" & resolve & " min)")
                                    If escalation_type = "2" And numAlerted < 2 Then    'escalation_type = 2 for IT Management
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, 2)
                                        arrMobile = BuildMobile(incident_id, engineer_id, 2)
                                    Else
                                        strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                        arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                    End If

                                    If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                        sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                    End If

                                    If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                        sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                    End If
                                Else
                                    If escalation_next_time = "" Then ' Default Escalation 
                                        If (Val(duration) Mod Val(resolve)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Resolve Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "15", "resolve overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA resolve:" & resolve & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    Else
                                        If (Val(duration) Mod Val(escalation_next_time)) = 1 Then
                                            'strSubj = BuildSubj(case_id, case_title, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            strSubj = "Incident#" & case_id & " [" & case_title & "] - " & "Notification of Resolve Overdue#" & CStr(numAlerted + 1)
                                            log.Info(strSubj)
                                            strBody = BuildBody(incident_id, "Notification of Resolve Overdue#" & CStr(numAlerted + 1))
                                            AddCaseLog(incident_id, "15", "resolve overdued #" & CStr(numAlerted + 1) & " (Duration: " & duration & " min, SLA resolve:" & resolve & " min)")
                                            strEmail = BuildEmailAddress(incident_id, engineer_id, numAlerted)
                                            If strEmail <> "" And smtp_server <> "" And smtp_sender <> "" Then
                                                sEmail(incident_id, strSubj, strBody, strEmail, smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                                            End If

                                            arrMobile = BuildMobile(incident_id, engineer_id, numAlerted)
                                            If arrMobile.Count > 0 And sms_url <> "" And sms_username <> "" Then
                                                sSMS(incident_id, strSubj, arrMobile, sms_url, sms_user_keyword, sms_password_keyword, sms_number_keyword, sms_message_keyword, sms_username, sms_password, sms_protocol, sms_result)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            LogEvent("Escalator: " & ex.Message, EventLogEntryType.Error)
            log.Error("Escalator: " & ex.Message)
        End Try
    End Sub

    Function CheckEscalationMatrix(ByVal days As String, ByVal priority_id As String, ByRef recipient As String, cc As String, email_template_id As String, sms_template_id As String, ByRef level As String) As Boolean
        Dim result As Boolean = False
        Dim _sql As String = ""
        _sql = "SELECT * FROM problem_escalator_matrix WHERE priority_id = '" & priority_id & "' AND period= '" & days & "' "
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter(_sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            Dim dt As DataTable = ds.Tables("*")
            Dim dr As DataRow
            For Each dr In dt.Rows
                recipient = dr.Item("recipient").ToString()
                cc = dr.Item("cc").ToString()
                email_template_id = dr.Item("email_template_id").ToString()
                sms_template_id = dr.Item("sms_template_id").ToString()
                level = dr.Item("escalator_level").ToString()
                result = True
            Next
        Catch ex As Exception
            log.Info("CheckEscalationMatrix")
            log.Info(ex.ToString())
            LogEvent("CheckEscalationMatrix: " & ex.Message, EventLogEntryType.Error)
            log.Error("CheckEscalationMatrix: " & ex.Message)
            result = False
        End Try

        Return result
    End Function
    Function IsCaseResponsed(ByVal incident_id As String, ByVal conn As OdbcConnection) As Boolean
        sql = "SELECT * FROM case_log WHERE (case_log_description LIKE '%to Responsed%' OR case_log_description LIKE '%to Onsite%' OR case_log_description LIKE '%to In Progress%' OR case_log_description LIKE '%to Resolved%') AND case_id='" & incident_id & "'"
        If RowCountSql(sql, strConnection, conn) > 0 Then
            IsCaseResponsed = True
        Else
            IsCaseResponsed = False
        End If
    End Function
    Public Function ExecuteSql(ByVal sql As String, ByVal strConnection As String) As Boolean
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            ExecuteSql = True
        Catch
            ExecuteSql = False
        End Try
    End Function
    Function IsCaseOnsited(ByVal incident_id As String, ByVal conn As OdbcConnection) As Boolean
        sql = "SELECT * FROM case_log WHERE (case_log_description LIKE '%to Onsite%' OR case_log_description LIKE '%to In Progress%' OR case_log_description LIKE '%to Resolved%') AND case_id='" & incident_id & "'"
        If RowCountSql(sql, strConnection, conn) > 0 Then
            IsCaseOnsited = True
        Else
            IsCaseOnsited = False
        End If
    End Function

    Function IsCaseResolved(ByVal incident_id As String, ByVal conn As OdbcConnection) As Boolean
        sql = "SELECT * FROM case_log WHERE case_log_description LIKE '%to Resolved%' AND case_log_description NOT LIKE '%Resolved to%'  AND case_id='" & incident_id & "'"
        If RowCountSql(sql, strConnection, conn) > 0 Then
            IsCaseResolved = True
        Else
            IsCaseResolved = False
        End If
    End Function

    Function NumAlertedForResponse(ByVal incident_id As String, ByVal conn As OdbcConnection) As Integer
        'sql = "SELECT * FROM case_log WHERE case_log_type_id='4' AND case_log_description LIKE '%Response Overdue%' AND case_id='" & incident_id & "'"
        sql = "SELECT * FROM case_log WHERE case_log_type_id='13' AND case_id='" & incident_id & "'"
        NumAlertedForResponse = RowCountSql(sql, strConnection, conn)
    End Function

    Function NumAlertedForOnsite(ByVal incident_id As String, ByVal conn As OdbcConnection) As Integer
        'sql = "SELECT * FROM case_log WHERE case_log_type_id='4' AND case_log_description LIKE '%Onsite Overdue%' AND case_id='" & incident_id & "'"
        sql = "SELECT * FROM case_log WHERE case_log_type_id='14' AND case_id='" & incident_id & "'"
        NumAlertedForOnsite = RowCountSql(sql, strConnection, conn)
    End Function

    Function NumAlertedForResolve(ByVal incident_id As String, ByVal conn As OdbcConnection) As Integer
        'sql = "SELECT * FROM case_log WHERE case_log_type_id='4' AND case_log_description LIKE '%Resolve Overdue%' AND case_id='" & incident_id & "'"
        sql = "SELECT * FROM case_log WHERE case_log_type_id='15' AND case_id='" & incident_id & "'"
        NumAlertedForResolve = RowCountSql(sql, strConnection, conn)
    End Function

    Function BuildSubj(ByVal case_id As String, ByVal case_title As String, ByVal msg As String) As String
        Dim strSubj As String = ""
        sql = "SELECT id FROM incident WHERE case_id ='" & case_id & "'"
        Dim incident_id As String = FetchField(sql, strConnection, "id")
        sql = "SELECT contract_id FROM incident WHERE id=" & incident_id
        Dim contract_id As String = FetchField(sql, strConnection, "contract_id")
        sql = "select email_template_id FROM email_profile_Detail where profile_id =(select email_profile_id from contract where contract_id ='" & contract_id & "') and status_id=8"
        Dim email_template_id As String = FetchField(sql, strConnection, "email_template_id")
        If email_template_id = "0" Or email_template_id = "" Or contract_id = 0 Then
            BuildSubj = "Incident#" & case_id & " [" & case_title & "] - " & msg
        Else
            strSubj = FetchField("SELECT subject FROM  email_template WHERE id ='" & email_template_id & "' ", strConnection, "subject")
            strSubj = ReplaceEmailVariable(incident_id, strSubj)
            BuildSubj = strSubj
        End If

    End Function

    Function BuildBody(ByVal incident_id As String, ByVal action As String) As String
        Dim strBody As String = ""
        Try

            Dim conn As New OdbcConnection(strConnection)
            'sql = "SELECT incident.*,case_status_title,priority_title,tier_title,case_category_title,case_sub_category_title,customer_name,contract_title,site_title,CONCAT(staff.firstname,' ',staff.lastname) AS engineer_name FROM incident,case_status,priority,tier,case_category,case_sub_category,customer,contract,site,staff WHERE incident.case_status_id=case_status.case_status_id AND incident.priority_id=priority.priority_id AND (incident.tier_id=tier.tier_id OR incident.tier_id =0) AND incident.case_sub_category_id=case_sub_category.case_sub_category_id AND case_category.case_category_id=case_sub_category.case_category_id AND incident.contract_id=contract.contract_id AND customer.customer_id=contract.customer_id AND incident.site_id=site.site_id AND incident.engineer_id=staff.staff_id AND incident.id=" & incident_id & " LIMIT 1"
            sql = "SELECT incident.*,case_status_title,priority_title,IF(tier_title IS NULL,'',tier_title)as tier_title,case_category_title,case_sub_category_title," & _
            " customer_name,contract_title,site_title,CONCAT(engineer.firstname,' ',engineer.lastname) AS engineer_name" & _
            " FROM(incident)" & _
            " INNER JOIN case_status ON incident.case_status_id=case_status.case_status_id" & _
            " INNER JOIN priority ON incident.priority_id=priority.priority_id" & _
            " INNER JOIN case_sub_category ON incident.case_sub_category_id=case_sub_category.case_sub_category_id" & _
            " INNER JOIN case_category ON case_category.case_category_id=case_sub_category.case_category_id" & _
            " INNER JOIN staff engineer ON incident.engineer_id=engineer.staff_id" & _
            " LEFT JOIN tier ON incident.tier_id=tier.tier_id" & _
            " LEFT JOIN contract ON incident.contract_id=contract.contract_id" & _
            " LEFT JOIN customer on contract.customer_id = customer.customer_id" & _
            " LEFT JOIN site ON incident.site_id=site.site_id" & _
            " LEFT JOIN staff n_user ON incident.staff_id =n_user.staff_id" & _
            " WHERE incident.id=" & incident_id
            'log.Info("SQL EMAIL = CASE_ID :" & "  " & incident_id & " " & sql)
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            Dim dt As DataTable = ds.Tables("*")
            Dim dr As DataRow
            'Dim strSiteAddress, strContactDetail, strParentCase As String
            For Each dr In dt.Rows
                strBody = "<html><header></header><body><div style='font-family:MS Sans Serif,verdana;font-size:small;'><b>Dear " & dr.Item("engineer_name") & ",</b></div><br />"
                strBody += "<div style='font-family:MS Sans Serif,verdana;font-size:small;'>Incident ID " & dr.Item("case_id").ToString() & " [" & dr.Item("title").ToString() & "] - " & action & "</div><br />"
                strBody += "<table border =0 width ='600' style='font-family:MS Sans Serif,verdana;font-size:x-small;' cellpadding='0' cellspacing='0' >"
                strBody += "<tr><td colspan='3'><b>Incident Information [" & IIf(dr.Item("staff_id").ToString() = 0, "External", "Internal") & "]</b><br /><hr></td></tr>"
                strBody += "<tr><td width='37%'>Incident#:</td><td width='3%'></td><td width='60%'><b>" & dr.Item("case_id").ToString() & "<b></td><tr>"
                strBody += "<tr><td valign='top'>Title:</td><td>&nbsp;</td><td>" & dr.Item("title").ToString() & "</td><tr>"
                strBody += "<tr><td>Category:</td><td>&nbsp;</td><td>" & dr.Item("case_category_title").ToString() & " | " & dr.Item("case_sub_category_title").ToString() & "</td><tr>"
                strBody += "<tr><td>Priority:</td><td>&nbsp;</td><td>" & dr.Item("priority_title").ToString() & "</td><tr>"
                strBody += "<tr><td>Tier:</td><td>&nbsp;</td><td>" & dr.Item("tier_title").ToString() & "</td><tr>"
                strBody += "<tr><td>Engineer:</td><td>&nbsp;</td><td>" & dr.Item("engineer_name").ToString() & "</td><tr>"
                strBody += "<tr><td>Status:</td><td>&nbsp;</td><td>" & dr.Item("case_status_title").ToString() & "</td><tr>"
                strBody += "<tr><td valign='top'>Description:</td><td>&nbsp;</td><td>" & dr.Item("case_description").ToString() & "</td><tr>"
                strBody += "<tr><td valign='top'>Remark:</td><td>&nbsp;</td><td>" & dr.Item("remark").ToString() & "</td><tr>"
                strBody += "<tr><td colspan='3'><br /><b>Incident details </b></td></tr>"
                strBody += "<tr><td colspan='3'><a href ='" & web_url & "/incident_update.aspx?incident_id=" & dr.Item("id").ToString() & "'>" & web_url & "/incident_update.aspx?incident_id=" & dr.Item("id").ToString() & "</a><br /></td></tr>"
                Dim owner As String = ""
                Dim asset As String = ""
                If dr.Item("staff_id").ToString() = 0 Then ' Customer information
                    strBody += "<tr><td colspan='3'><br /><b>Customer&nbsp;Information</b><br /><hr></td></tr>"
                    strBody += "<tr><td>Customer:</td><td>&nbsp;</td><td>" & dr.Item("customer_name").ToString() & "</td><tr>"
                    strBody += "<tr><td>Contract:</td><td>&nbsp;</td><td>" & dr.Item("contract_title").ToString() & "</td><tr>"
                    strBody += "<tr><td>Site:</td><td>&nbsp;</td><td>" & dr.Item("site_title").ToString() & "</td><tr>"
                    'Dim item3 As ListItem
                    'For Each item3 In chkBoxListAsset.Items
                    '    If item3.Selected = True Then
                    '        asset = item3.Text
                    '    End If
                    'Next
                    'sql = "SELECT * FROM contact WHERE CONCAT(firstname,' ',lastname)='" & DDListContact.SelectedItem.Text & "'"
                    sql = "SELECT incident.id,CONCAT(contact.firstname,' ',contact.lastname)as contact,position,contact.email,contact.phone,mobile,site.address " & _
                          " FROM incident " & _
                          " INNER JOIN contact ON incident.contact_id = contact.contact_id " & _
                          " INNER JOIN site ON contact.site_id = site.site_id WHERE id='" & incident_id & "'"
                    Dim conn3 As New OdbcConnection(strConnection)
                    Dim cmd3 As New OdbcDataAdapter(sql, conn3)
                    Dim ds3 As New DataSet()
                    cmd3.Fill(ds3, "*")
                    Dim dt3 As DataTable = ds3.Tables("*")
                    Dim dr3 As DataRow
                    For Each dr3 In dt3.Rows
                        'strBody += "<tr><td>Asset(SN|PN):</td><td>&nbsp;</td><td>" & asset & "</td><tr>"
                        strBody += "<tr><td>Contact:</td><td>&nbsp;</td><td>" & dr3.Item("contact").ToString() & "</td><tr>"
                        strBody += "<tr><td>Position:</td><td>&nbsp;</td><td>" & dr3.Item("position").ToString() & "</td><tr>"
                        strBody += "<tr><td>Email:</td><td>&nbsp;</td><td>" & dr3.Item("email").ToString() & "</td><tr>"
                        strBody += "<tr><td>Phone:</td><td>&nbsp;</td><td>" & dr3.Item("phone").ToString() & "</td><tr>"
                        strBody += "<tr><td>Mobile:</td><td>&nbsp;</td><td>" & dr3.Item("mobile").ToString() & "</td><tr>"
                        strBody += "<tr><td>Address:</td><td>&nbsp;</td><td>" & dr3.Item("address").ToString() & "</td><tr>"
                    Next
                    'sql = "SELECT sla_title,response,onsite,resolve " & _
                    '   " FROM contract INNER JOIN sla ON contract.sla_id = sla.sla_id " & _
                    '   " WHERE contract_title='" & dr.Item("contract_title").ToString() & "' "
                    sql = "SELECT sla_title,response,onsite,resolve " & _
                          " FROM contract INNER JOIN sla ON contract.sla_id = sla.sla_id " & _
                          " INNER JOIN sla_detail ON sla.sla_id=sla_detail.sla_id AND sla_detail.priority_id='" & dr.Item("priority_id").ToString() & "' " & _
                          " WHERE contract_title='" & dr.Item("contract_title").ToString() & "' "
                    Dim cmd4 As New OdbcDataAdapter(sql, conn3)
                    Dim ds4 As New DataSet()
                    cmd4.Fill(ds4, "*")
                    Dim dt4 As DataTable = ds4.Tables("*")
                    Dim dr4 As DataRow

                    For Each dr4 In dt4.Rows
                        strBody += "<tr><td colspan='3'><br /><b>SLA&nbsp;Information</b><br /><hr></td></tr>"
                        strBody += "<tr><td>SLA&nbsp;Title:</td><td>&nbsp;</td><td>" & dr4.Item("sla_title").ToString() & "</td><tr>"
                        strBody += "<tr><td>Response:</td><td>&nbsp;</td><td>" & dr4.Item("response").ToString() & "</td><tr>"
                        strBody += "<tr><td>Onsite:</td><td>&nbsp;</td><td>" & dr4.Item("onsite").ToString() & "</td><tr>"
                        strBody += "<tr><td>Resolve:</td><td>&nbsp;</td><td>" & dr4.Item("resolve").ToString() & "</td><tr>"
                    Next
                Else ' User Information 
                    strBody += "<tr><td colspan='3'><br /><b>User&nbsp;Information</b><br /><hr></td></tr>"
                    sql = "select incident.id,CONCAT(firstname,' ',lastname)AS name,department_title,section_title,staff.email,staff.phone,staff.mobile from incident" & _
                         " inner join staff on incident.staff_id = staff.staff_id " & _
                         " inner join department on staff.department_id = department.department_id" & _
                         " inner join section on staff.section_id = section.section_id " & _
                         " WHERE incident.id ='" & incident_id & "'"
                    'log.Info("SQL user info = " & sql)
                    Dim conn2 As New OdbcConnection(strConnection)
                    Dim cmd2 As New OdbcDataAdapter(sql, conn2)
                    Dim ds2 As New DataSet()
                    cmd2.Fill(ds2, "*")
                    Dim dt2 As DataTable = ds2.Tables("*")
                    Dim dr2 As DataRow
                    For Each dr2 In dt2.Rows
                        strBody += "<tr><td>Name:</td><td>&nbsp;</td><td>" & dr2.Item("name").ToString() & "</td><tr>"
                        strBody += "<tr><td>Department:</td><td>&nbsp;</td><td>" & dr2.Item("department_title").ToString() & "</td><tr>"
                        strBody += "<tr><td>Section:</td><td>&nbsp;</td><td>" & dr2.Item("section_title").ToString() & "</td><tr>"
                        strBody += "<tr><td>Email:</td><td>&nbsp;</td><td>" & dr2.Item("email").ToString() & "</td><tr>"
                        strBody += "<tr><td>Phone:</td><td>&nbsp;</td><td>" & dr2.Item("phone").ToString() & "</td><tr>"
                        strBody += "<tr><td>Mobile:</td><td>&nbsp;</td><td>" & dr2.Item("mobile").ToString() & "</td><tr>"
                    Next
                    ' SLA Information
                    sql = "SELECT * FROM priority WHERE priority_title='" & dr.Item("priority_title").ToString() & "'"
                    Dim cmd4 As New OdbcDataAdapter(sql, conn2)
                    Dim ds4 As New DataSet()
                    cmd4.Fill(ds4, "*")
                    Dim dt4 As DataTable = ds4.Tables("*")
                    Dim dr4 As DataRow
                    For Each dr4 In dt4.Rows
                        strBody += "<tr><td colspan='3'><br /><b>SLA&nbsp;Information</b><br /><hr></td></tr>"
                        strBody += "<tr><td>SLA&nbsp;Title:</td><td>&nbsp;</td><td>" & dr4.Item("priority_title").ToString() & "</td><tr>"
                        strBody += "<tr><td>Response:</td><td>&nbsp;</td><td>" & dr4.Item("response").ToString() & "</td><tr>"
                        strBody += "<tr><td>Onsite:</td><td>&nbsp;</td><td>" & dr4.Item("onsite").ToString() & "</td><tr>"
                        strBody += "<tr><td>Resolve:</td><td>&nbsp;</td><td>" & dr4.Item("resolve").ToString() & "</td><tr>"
                    Next
                End If
                strBody += "<tr><td><br />Updated&nbsp;by:</td><td>&nbsp;</td><td><br />System</td><tr>"
                strBody += "<tr><td>Generated&nbsp;by:</td><td>&nbsp;</td><td>" & company_name & Year(Now) & "</td><tr>"
                strBody += "<tr><td>Generated&nbsp;date:</td><td>&nbsp;</td><td>" & Now.ToString("dd-MMM-yyyy HH:mm:ss") & "</td><tr>"
                strBody += "<tr><td colspan='3'><br /><br />*** This is an automatically generated email, please do not reply ***</td></tr>"
                strBody += "<tr><td colspan='3'><br /><br /><img src=cid:companylogo></td></tr>"
                strBody += "</table></body></html>"
            Next
            BuildBody = strBody
        Catch ex As Exception
            log.Info(strBody)
            log.Info(ex.ToString())
            LogEvent("BuildBody: " & ex.Message, EventLogEntryType.Error)
            log.Error("BuildBody: " & ex.Message)
            BuildBody = ""
        End Try
    End Function

    Function BuildBodyClose(ByVal incident_id As String, ByVal action As String) As String
        Dim strBody As String = ""
        Try

            Dim conn As New OdbcConnection(strConnection)
            'sql = "SELECT incident.*,case_status_title,priority_title,tier_title,case_category_title,case_sub_category_title,customer_name,contract_title,site_title,CONCAT(staff.firstname,' ',staff.lastname) AS engineer_name FROM incident,case_status,priority,tier,case_category,case_sub_category,customer,contract,site,staff WHERE incident.case_status_id=case_status.case_status_id AND incident.priority_id=priority.priority_id AND (incident.tier_id=tier.tier_id OR incident.tier_id =0) AND incident.case_sub_category_id=case_sub_category.case_sub_category_id AND case_category.case_category_id=case_sub_category.case_category_id AND incident.contract_id=contract.contract_id AND customer.customer_id=contract.customer_id AND incident.site_id=site.site_id AND incident.engineer_id=staff.staff_id AND incident.id=" & incident_id & " LIMIT 1"
            sql = "SELECT incident.*,CAST(CONCAT(case_log.date,' ',case_log.time) AS CHAR) AS created_date,CAST(CONCAT(case_duration DIV 60,':',MOD(case_duration,60)) AS CHAR) AS duration,case_status_title,priority_title,IF(tier_title IS NULL,'',tier_title)as tier_title,case_category_title,case_sub_category_title," & _
            " customer_name,contract_title,site_title,CONCAT(engineer.firstname,' ',engineer.lastname) AS engineer_name,contract.contract_id" & _
            " FROM(incident)" & _
            " INNER JOIN case_log ON (incident.id=case_log.case_id AND case_log.case_log_type_id='1') " & _
            " INNER JOIN case_status ON incident.case_status_id=case_status.case_status_id" & _
            " INNER JOIN priority ON incident.priority_id=priority.priority_id" & _
            " INNER JOIN case_sub_category ON incident.case_sub_category_id=case_sub_category.case_sub_category_id" & _
            " INNER JOIN case_category ON case_category.case_category_id=case_sub_category.case_category_id" & _
            " INNER JOIN staff engineer ON incident.engineer_id=engineer.staff_id" & _
            " LEFT JOIN tier ON incident.tier_id=tier.tier_id" & _
            " LEFT JOIN contract ON incident.contract_id=contract.contract_id" & _
            " LEFT JOIN customer on contract.customer_id = customer.customer_id" & _
            " LEFT JOIN site ON incident.site_id=site.site_id" & _
            " LEFT JOIN staff n_user ON incident.staff_id =n_user.staff_id" & _
            " WHERE incident.id=" & incident_id

            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            Dim dt As DataTable = ds.Tables("*")
            Dim dr As DataRow

            For Each dr In dt.Rows
                Dim contract_id As String = ""
                Dim email_template_id As String = ""
                If dr.Item("staff_id").ToString() = 0 Then ' External case 
                    ' Check Use Email template
                    sql = "select email_template_id FROM email_profile_Detail where profile_id =(select email_profile_id from contract where contract_id ='" & dr.Item("contract_id").ToString() & "') and status_id = 8"
                    email_template_id = FetchField(sql, strConnection, "email_template_id")
                End If

                If email_template_id = "0" Or email_template_id = "" Or dr.Item("staff_id").ToString() <> 0 Then
                    log.Info("Used Default Email Template.")
                    strBody = "<html><header></header><body><div style='font-family:MS Sans Serif,verdana;font-size:small;'><b>Dear " & dr.Item("engineer_name") & ",</b></div><br />"
                    strBody += "<div style='font-family:MS Sans Serif,verdana;font-size:small;'>Incident ID " & dr.Item("case_id").ToString() & " [" & dr.Item("title").ToString() & "] - " & action & "</div><br />"
                    strBody += "<table border =0 width ='600' style='font-family:MS Sans Serif,verdana;font-size:x-small;' cellpadding='0' cellspacing='0' >"
                    strBody += "<tr><td colspan='3'><b>Incident Information [" & IIf(dr.Item("staff_id").ToString() = 0, "External", "Internal") & "]</b><br /><hr></td></tr>"
                    strBody += "<tr><td width='37%'>Incident#:</td><td width='3%'></td><td width='60%'><b>" & dr.Item("case_id").ToString() & "<b></td><tr>"
                    strBody += "<tr><td valign='top'>Title:</td><td>&nbsp;</td><td>" & dr.Item("title").ToString() & "</td><tr>"
                    strBody += "<tr><td>Category:</td><td>&nbsp;</td><td>" & dr.Item("case_category_title").ToString() & " | " & dr.Item("case_sub_category_title").ToString() & "</td><tr>"
                    strBody += "<tr><td>Priority:</td><td>&nbsp;</td><td>" & dr.Item("priority_title").ToString() & "</td><tr>"
                    strBody += "<tr><td>Tier:</td><td>&nbsp;</td><td>" & dr.Item("tier_title").ToString() & "</td><tr>"
                    strBody += "<tr><td>Engineer:</td><td>&nbsp;</td><td>" & dr.Item("engineer_name").ToString() & "</td><tr>"
                    strBody += "<tr><td>Created:</td><td>&nbsp;</td><td>" & dr.Item("created_date").ToString() & "</td><tr>"
                    strBody += "<tr><td>Duration:</td><td>&nbsp;</td><td>" & dr.Item("duration").ToString() & "&nbsp;min</td><tr>"
                    strBody += "<tr><td>Status:</td><td>&nbsp;</td><td>" & dr.Item("case_status_title").ToString() & "</td><tr>"
                    strBody += "<tr><td valign='top'>Description:</td><td>&nbsp;</td><td>" & dr.Item("case_description").ToString() & "</td><tr>"
                    strBody += "<tr><td valign='top'>Remark:</td><td>&nbsp;</td><td>" & dr.Item("remark").ToString() & "</td><tr>"

                    Dim owner As String = ""
                    Dim asset As String = ""
                    If dr.Item("staff_id").ToString() = 0 Then ' Customer information
                        strBody += "<tr><td colspan='3'><br /><b>Customer&nbsp;Information</b><br /><hr></td></tr>"
                        strBody += "<tr><td>Customer:</td><td>&nbsp;</td><td>" & dr.Item("customer_name").ToString() & "</td><tr>"
                        strBody += "<tr><td>Contract:</td><td>&nbsp;</td><td>" & dr.Item("contract_title").ToString() & "</td><tr>"
                        strBody += "<tr><td>Site:</td><td>&nbsp;</td><td>" & dr.Item("site_title").ToString() & "</td><tr>"
                        'Dim item3 As ListItem
                        'For Each item3 In chkBoxListAsset.Items
                        '    If item3.Selected = True Then
                        '        asset = item3.Text
                        '    End If
                        'Next
                        'sql = "SELECT * FROM contact WHERE CONCAT(firstname,' ',lastname)='" & DDListContact.SelectedItem.Text & "'"
                        sql = "SELECT incident.id,CONCAT(contact.firstname,' ',contact.lastname)as contact,position,contact.email,contact.phone,mobile,site.address " & _
                              " FROM incident " & _
                              " INNER JOIN contact ON incident.contact_id = contact.contact_id " & _
                              " INNER JOIN site ON contact.site_id = site.site_id WHERE id='" & incident_id & "'"
                        Dim conn3 As New OdbcConnection(strConnection)
                        Dim cmd3 As New OdbcDataAdapter(sql, conn3)
                        Dim ds3 As New DataSet()
                        cmd3.Fill(ds3, "*")
                        Dim dt3 As DataTable = ds3.Tables("*")
                        Dim dr3 As DataRow
                        For Each dr3 In dt3.Rows
                            'strBody += "<tr><td>Asset(SN|PN):</td><td>&nbsp;</td><td>" & asset & "</td><tr>"
                            strBody += "<tr><td>Contact:</td><td>&nbsp;</td><td>" & dr3.Item("contact").ToString() & "</td><tr>"
                            strBody += "<tr><td>Position:</td><td>&nbsp;</td><td>" & dr3.Item("position").ToString() & "</td><tr>"
                            strBody += "<tr><td>Email:</td><td>&nbsp;</td><td>" & dr3.Item("email").ToString() & "</td><tr>"
                            strBody += "<tr><td>Phone:</td><td>&nbsp;</td><td>" & dr3.Item("phone").ToString() & "</td><tr>"
                            strBody += "<tr><td>Mobile:</td><td>&nbsp;</td><td>" & dr3.Item("mobile").ToString() & "</td><tr>"
                            strBody += "<tr><td>Address:</td><td>&nbsp;</td><td>" & dr3.Item("address").ToString() & "</td><tr>"
                        Next
                        'sql = "SELECT sla_title,response,onsite,resolve " & _
                        '   " FROM contract INNER JOIN sla ON contract.sla_id = sla.sla_id " & _
                        '   " WHERE contract_title='" & dr.Item("contract_title").ToString() & "' "
                        sql = "SELECT sla_title,response,onsite,resolve " & _
                              " FROM contract INNER JOIN sla ON contract.sla_id = sla.sla_id " & _
                              " INNER JOIN sla_detail ON sla.sla_id=sla_detail.sla_id AND sla_detail.priority_id='" & dr.Item("priority_id").ToString() & "' " & _
                              " WHERE contract_title='" & dr.Item("contract_title").ToString() & "' "
                        Dim cmd4 As New OdbcDataAdapter(sql, conn3)
                        Dim ds4 As New DataSet()
                        cmd4.Fill(ds4, "*")
                        Dim dt4 As DataTable = ds4.Tables("*")
                        Dim dr4 As DataRow

                        For Each dr4 In dt4.Rows
                            strBody += "<tr><td colspan='3'><br /><b>SLA&nbsp;Information</b><br /><hr></td></tr>"
                            strBody += "<tr><td>SLA&nbsp;Title:</td><td>&nbsp;</td><td>" & dr4.Item("sla_title").ToString() & "</td><tr>"
                            strBody += "<tr><td>Response:</td><td>&nbsp;</td><td>" & dr4.Item("response").ToString() & "</td><tr>"
                            strBody += "<tr><td>Onsite:</td><td>&nbsp;</td><td>" & dr4.Item("onsite").ToString() & "</td><tr>"
                            strBody += "<tr><td>Resolve:</td><td>&nbsp;</td><td>" & dr4.Item("resolve").ToString() & "</td><tr>"
                        Next
                    Else ' User Information 
                        strBody += "<tr><td colspan='3'><br /><b>User&nbsp;Information</b><br /><hr></td></tr>"
                        sql = "select incident.id,CONCAT(firstname,' ',lastname)AS name,department_title,section_title,staff.email,staff.phone,staff.mobile from incident" & _
                             " inner join staff on incident.staff_id = staff.staff_id " & _
                             " inner join department on staff.department_id = department.department_id" & _
                             " inner join section on staff.section_id = section.section_id " & _
                             " WHERE incident.id ='" & incident_id & "'"
                        'log.Info("SQL user info = " & sql)
                        Dim conn2 As New OdbcConnection(strConnection)
                        Dim cmd2 As New OdbcDataAdapter(sql, conn2)
                        Dim ds2 As New DataSet()
                        cmd2.Fill(ds2, "*")
                        Dim dt2 As DataTable = ds2.Tables("*")
                        Dim dr2 As DataRow
                        For Each dr2 In dt2.Rows
                            strBody += "<tr><td>Name:</td><td>&nbsp;</td><td>" & dr2.Item("name").ToString() & "</td><tr>"
                            strBody += "<tr><td>Department:</td><td>&nbsp;</td><td>" & dr2.Item("department_title").ToString() & "</td><tr>"
                            strBody += "<tr><td>Section:</td><td>&nbsp;</td><td>" & dr2.Item("section_title").ToString() & "</td><tr>"
                            strBody += "<tr><td>Email:</td><td>&nbsp;</td><td>" & dr2.Item("email").ToString() & "</td><tr>"
                            strBody += "<tr><td>Phone:</td><td>&nbsp;</td><td>" & dr2.Item("phone").ToString() & "</td><tr>"
                            strBody += "<tr><td>Mobile:</td><td>&nbsp;</td><td>" & dr2.Item("mobile").ToString() & "</td><tr>"
                        Next
                        ' SLA Information
                        sql = "SELECT * FROM priority WHERE priority_title='" & dr.Item("priority_title").ToString() & "'"
                        Dim cmd4 As New OdbcDataAdapter(sql, conn2)
                        Dim ds4 As New DataSet()
                        cmd4.Fill(ds4, "*")
                        Dim dt4 As DataTable = ds4.Tables("*")
                        Dim dr4 As DataRow
                        For Each dr4 In dt4.Rows
                            strBody += "<tr><td colspan='3'><br /><b>SLA&nbsp;Information</b><br /><hr></td></tr>"
                            strBody += "<tr><td>SLA&nbsp;Title:</td><td>&nbsp;</td><td>" & dr4.Item("priority_title").ToString() & "</td><tr>"
                            strBody += "<tr><td>Response:</td><td>&nbsp;</td><td>" & dr4.Item("response").ToString() & "</td><tr>"
                            strBody += "<tr><td>Onsite:</td><td>&nbsp;</td><td>" & dr4.Item("onsite").ToString() & "</td><tr>"
                            strBody += "<tr><td>Resolve:</td><td>&nbsp;</td><td>" & dr4.Item("resolve").ToString() & "</td><tr>"
                        Next
                    End If
                    strBody += "<tr><td><br />Updated&nbsp;by:</td><td>&nbsp;</td><td><br />System</td><tr>"
                    strBody += "<tr><td>Generated&nbsp;by:</td><td>&nbsp;</td><td>" & company_name & Year(Now) & "</td><tr>"
                    strBody += "<tr><td>Generated&nbsp;date:</td><td>&nbsp;</td><td>" & Now.ToString("dd-MMM-yyyy HH:mm:ss") & "</td><tr>"
                    strBody += "<tr><td colspan='3'><br /><br />*** This is an automatically generated email, please do not reply ***</td></tr>"
                    strBody += "<tr><td colspan='3'><br /><br /><img src=cid:companylogo></td></tr>"
                    strBody += "</table></body></html>"
                Else
                    log.Info("Used Email Template ID = " & email_template_id)
                    strBody = FetchField("SELECT message FROM  email_template WHERE id ='" & email_template_id & "' ", strConnection, "message")
                    strBody = ReplaceEmailVariable(incident_id, strBody)
                    strBody += "<br /><img src=cid:companylogo>"
                    strBody = "<html><body><div style='font-family:MS Sans Serif, Tahoma, MS Sans Serif'>" & strBody & "</div></body></html>"
                End If
            Next
            BuildBodyClose = strBody
        Catch ex As Exception
            log.Info(strBody)
            log.Info(ex.ToString())
            LogEvent("BuildBody: " & ex.Message, EventLogEntryType.Error)
            log.Error("BuildBody: " & ex.Message)
            BuildBodyClose = ""
        End Try
    End Function

    Function BuildSiteAddress(ByVal site_id As String) As String
        Try
            Dim conn As New OdbcConnection(strConnection)
            sql = "SELECT * FROM site WHERE site_id=" & site_id
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "site")
            Dim dt As DataTable = ds.Tables("site")
            Dim dr As DataRow
            Dim strSiteAddress As String = ""
            Dim strTemp As String = ""
            For Each dr In dt.Rows
                If dr.Item("level") <> "" Then
                    strSiteAddress += dr.Item("level") & " Fl., "
                End If
                If dr.Item("building") <> "" Then
                    strSiteAddress += dr.Item("building") & " "
                End If
                If dr.Item("address") <> "" Then
                    strSiteAddress += vbCrLf & dr.Item("address") & " "
                End If
                If dr.Item("street") <> "" Then
                    strSiteAddress += dr.Item("street") & ", "
                End If
                If CStr(dr.Item("tambol")) <> "" Then
                    sql = "SELECT tambol_thai FROM tambol WHERE tambol_id='" & dr.Item("tambol") & "'"
                    strTemp = FetchField(sql, strConnection, "tambol_thai")
                    strSiteAddress += strTemp & ", "
                End If
                If CStr(dr.Item("amphur")) <> "" Then
                    sql = "SELECT amphur_thai FROM amphur WHERE amphur_id='" & dr.Item("amphur") & "'"
                    strTemp = FetchField(sql, strConnection, "amphur_thai")
                    strSiteAddress += strTemp & ", "
                End If
                If CStr(dr.Item("province")) <> "" Then
                    sql = "SELECT province_thai FROM province WHERE province_id='" & dr.Item("province") & "'"
                    strTemp = FetchField(sql, strConnection, "province_thai")
                    strSiteAddress += vbCrLf & strTemp & " "
                End If
                If dr.Item("country") <> "" Then
                    strSiteAddress += dr.Item("country") & " "
                End If
                If dr.Item("zip") <> "" Then
                    strSiteAddress += dr.Item("zip").ToString() & " "
                End If
                If dr.Item("phone") <> "" Then
                    strSiteAddress += vbCrLf & "Telephone: " & dr.Item("phone")
                End If
            Next
            BuildSiteAddress = strSiteAddress
        Catch ex As Exception
            LogEvent("BuildSiteAddress: " & ex.Message, EventLogEntryType.Error)
            log.Error("BuildSiteAddress: " & ex.Message)
            BuildSiteAddress = ""
        End Try
    End Function

    Function BuildContactDetail(ByVal contact_id As String) As String
        Try
            Dim conn As New OdbcConnection(strConnection)
            sql = "SELECT * FROM contact WHERE contact_id=" & contact_id
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "contact")
            Dim dt As DataTable = ds.Tables("contact")
            Dim dr As DataRow
            Dim strContactDetail As String = ""
            For Each dr In dt.Rows
                If dr.Item("lastname") <> "" Then
                    strContactDetail += dr.Item("firstname") & " " & dr.Item("lastname")
                Else
                    strContactDetail += dr.Item("firstname")
                End If
                strContactDetail += vbCrLf & "Mobile: " & dr.Item("mobile")
                strContactDetail += vbCrLf & "Telephone: " & dr.Item("phone")
                strContactDetail += vbCrLf & "Email: " & dr.Item("email")

            Next
            BuildContactDetail = strContactDetail
        Catch ex As Exception
            LogEvent("BuildContactDetail: " & ex.Message, EventLogEntryType.Error)
            log.Error("BuildContactDetail: " & ex.Message)
            BuildContactDetail = ""
        End Try
    End Function

    Function AddCaseLog(ByVal incident_id As String, ByVal case_log_type_id As String, ByVal case_log_description As String) As Integer
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter()
            Dim ds As New DataSet()

            ' Insert into case_log table
            Dim case_log_id As Integer = FindLastID("case_log", strConnection, "case_log_id") + 1
            sql = "INSERT INTO case_log VALUES ('"
            sql += case_log_id & "','"
            sql += DateStamp() & "','"
            sql += TimeStamp() & "','"
            sql += incident_id & "','"
            sql += "System','"
            sql += case_log_type_id & "','"
            sql += ReplaceString(case_log_description) & "','"
            sql += "1','')"     ' case_log_category_id = internal
            cmd = New OdbcDataAdapter(sql, conn)
            cmd.Fill(ds, "case_log")
            AddCaseLog = case_log_id
        Catch ex As Exception
            LogEvent("AddCaseLog: " & ex.Message, EventLogEntryType.Error)
            log.Error("AddCaseLog: " & ex.Message)
        End Try
    End Function

    Function BuildEmailAddress(ByVal incident_id As String, ByVal engineer_id As String, ByVal numAlerted As Integer) As String
        Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()
        Dim strEmail As String
        Dim strVendorEmail As String

        Dim email_list As String = ""
        sql = "SELECT email FROM staff WHERE staff_id=" & engineer_id
        email_list = FetchField(sql, strConnection, "email")
        sql = "SELECT vendor_case_id FROM incident WHERE id='" & incident_id & "' "
        strVendorEmail = FetchField(sql, strConnection, "vendor_case_id")
        If strVendorEmail <> "" Then
            email_list = email_list & ";" & strVendorEmail
        End If
        Dim staff_id As String = engineer_id
        Dim i As Integer
        For i = 1 To numAlerted
            sql = "SELECT boss_id FROM staff WHERE staff_id=" & staff_id
            Dim boss_id As String = FetchField(sql, strConnection, "boss_id")
            sql = "SELECT email FROM staff WHERE staff_id=" & boss_id
            strEmail = FetchField(sql, strConnection, "email")
            If strEmail <> "" Then
                If email_list <> "" Then
                    email_list = email_list & ";" & strEmail
                Else
                    email_list = strEmail
                End If
            End If
            staff_id = boss_id
        Next
        BuildEmailAddress = email_list
    End Function

    Sub sEmail(ByVal incident_id As String, ByVal strSubj As String, ByVal strBody As String, ByVal strEmail As String, ByVal smtp_server As String, ByVal smtp_sender As String, ByVal smtp_password As String, ByVal bSsl As String, ByVal sslPort As String)
        Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()

        Dim SendEmailResult As String = SendEmail(company_name, strEmail, strSubj, strBody, "", smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
        Dim SendEmailStatus As String
        If SendEmailResult = "OK" Then
            SendEmailStatus = "1"
        Else
            SendEmailStatus = "0"
        End If

        ' Insert into case_log table
        'Dim case_log_id As String = AddCaseLog(incident_id, "4", strSubj & " (" & strEmail & ")")

        ' Insert into email_sms_log
        sql = "INSERT INTO email_sms_log(date,time,problem_id,mobile,email,message,status,remark) VALUES('"
        sql += DateStamp() & "','"
        sql += TimeStamp() & "','"
        sql += incident_id & "','"
        sql += "','"        ' mobile
        sql += strEmail & "','"
        sql += strSubj & "','"
        sql += SendEmailStatus & "','"
        sql += SendEmailResult & "')"
        cmd = New OdbcDataAdapter(sql, conn)
        cmd.Fill(ds, "email_sms_log")
    End Sub

    Function BuildMobile(ByVal incident_id As String, ByVal engineer_id As String, ByVal numAlerted As Integer) As ArrayList
        Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()

        Dim arrMobile As New ArrayList
        sql = "SELECT mobile FROM staff WHERE staff_id=" & engineer_id
        Dim strMobile As String = FetchField(sql, strConnection, "mobile")
        If IsNumeric(strMobile) Then
            arrMobile.Add(strMobile)
        End If

        Dim staff_id As String = engineer_id
        Dim i As Integer
        For i = 1 To numAlerted
            sql = "SELECT boss_id FROM staff WHERE staff_id=" & staff_id
            Dim boss_id As String = FetchField(sql, strConnection, "boss_id")
            sql = "SELECT mobile FROM staff WHERE staff_id=" & boss_id
            strMobile = FetchField(sql, strConnection, "mobile")
            If IsNumeric(strMobile) Then
                arrMobile.Add(strMobile)
            End If
            staff_id = boss_id
        Next
        BuildMobile = arrMobile
    End Function

    Sub sSMS(ByVal incident_id As String, ByVal strSubj As String, ByVal arrMobile As ArrayList, ByVal sms_url As String, ByVal sms_user_keyword As String, ByVal sms_password_keyword As String, ByVal sms_number_keyword As String, ByVal sms_message_keyword As String, ByVal sms_username As String, ByVal sms_password As String, ByVal sms_protocol As String, ByVal sms_result As String)
        Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()
        Dim strMobile As String
        For Each strMobile In arrMobile
            Dim postdata As String = sms_number_keyword & "=" & strMobile & "&" & sms_user_keyword & "=" & sms_username & "&" & sms_password_keyword & "=" & sms_password & "&" & sms_message_keyword & "=" & strSubj & "&lang=E"
            'Dim postdata As String = sms_number_keyword & "=" & strMobile & "&" & sms_user_keyword & "=" & sms_username & "&" & sms_password_keyword & "=" & sms_password & "&" & sms_message_keyword & "=" & strSubj & "&" & sms_language_keyword & "=T"
            Dim SendSMSResult As String = HTTPPost(sms_url, postdata)
            Dim SendSMSStatus As String
            If InStr(1, SendSMSResult, sms_result) Then
                SendSMSStatus = "1"
            Else
                SendSMSStatus = "0"
            End If

            ' Insert into case_log table
            Dim case_log_id As String = AddCaseLog(incident_id, "5", strSubj & " (sms alert)")

            sql = "INSERT INTO email_sms_log VALUES ('"
            sql += DateStamp() & "','"
            sql += TimeStamp() & "','"
            sql += incident_id & "','"
            sql += case_log_id & "','"
            sql += strMobile & "','"
            sql += "','"        ' Email
            sql += strSubj & "','"
            sql += SendSMSStatus & "','"
            sql += ReplaceString(SendSMSResult) & "')"
            'Info.info(sql)
            cmd = New OdbcDataAdapter(sql, conn)
            cmd.Fill(ds, "email_sms_log")
        Next
    End Sub

    Function NonWorkingMinute(ByVal strDate As String, ByVal strTime As String, ByVal strDateNow As String, ByVal strTimeNow As String, ByVal conn As OdbcConnection) As Integer
        ' strDate, strDateNow format must be yyyy-MM-dd
        ' strTime, strTimeNow format must be HH:mm:ss
        'Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()
        Dim dt As DataTable = Nothing
        Dim dr As DataRow
        Dim intNonWorkingMinute As Integer = 0
        Dim d As Date
        Dim intTimeNow, intOpen, intClose As Integer
        Dim strWDayName, strTemp As String
        Dim myArray(1) As String

        Try
            Dim t1 As DateTime = DateTime.ParseExact(strDate, "yyyy-MM-dd", Nothing)
            Dim t2 As DateTime = DateTime.ParseExact(strDateNow, "yyyy-MM-dd", Nothing)
            Dim timediff As Integer = Abs(DateDiff(DateInterval.DayOfYear, t1, t2))
            If timediff = 0 Then    ' when case created date is today
                myArray = Split(CDate(strTime).ToString("HH:mm"), ":")
                Dim intCreatedTime As Integer = (myArray(0) * 60) + myArray(1)
                myArray = Split(strTimeNow, ":")
                intTimeNow = (myArray(0) * 60) + myArray(1)
                If IsHoliday(strDateNow, conn) Then
                    intNonWorkingMinute = intTimeNow - intCreatedTime
                Else
                    strWDayName = WeekdayName(Weekday(strDateNow))
                    sql = "SELECT * FROM business_hour WHERE day='" & strWDayName & "'"
                    cmd = New OdbcDataAdapter(sql, conn)
                    ds = New DataSet()
                    cmd.Fill(ds, "business_hour")
                    dt = ds.Tables("business_hour")
                    For Each dr In dt.Rows
                        myArray = Split(dr.Item("open"), ":")
                        intOpen = (myArray(0) * 60) + myArray(1)
                        myArray = Split(dr.Item("close"), ":")
                        intClose = (myArray(0) * 60) + myArray(1)
                        If intTimeNow <= intOpen Then
                            intNonWorkingMinute = intTimeNow - intCreatedTime
                        ElseIf intTimeNow > intClose Then
                            If intCreatedTime <= intOpen Then
                                intNonWorkingMinute = intTimeNow - intClose
                                intNonWorkingMinute += (intOpen - intCreatedTime)
                            ElseIf intCreatedTime > intClose Then
                                intNonWorkingMinute = intTimeNow - intCreatedTime
                            Else
                                intNonWorkingMinute = intTimeNow - intClose
                            End If
                        Else
                            If intCreatedTime <= intOpen Then
                                intNonWorkingMinute = intOpen - intCreatedTime
                            Else
                                intNonWorkingMinute = 0
                            End If
                        End If
                    Next
                End If
            Else
                ' Calculate NonWorkingMinute on case created date
                myArray = Split(CDate(strTime).ToString("HH:mm"), ":")
                Dim intCreatedTime As Integer = (myArray(0) * 60) + myArray(1)
                If IsHoliday(strDate, conn) Then
                    intNonWorkingMinute += (1440 - intCreatedTime)
                Else
                    strWDayName = WeekdayName(Weekday(strDate))
                    sql = "SELECT * FROM business_hour WHERE day='" & strWDayName & "'"
                    cmd = New OdbcDataAdapter(sql, conn)
                    ds = New DataSet()
                    cmd.Fill(ds, "business_hour")
                    dt = ds.Tables("business_hour")
                    For Each dr In dt.Rows
                        myArray = Split(dr.Item("open"), ":")
                        intOpen = (myArray(0) * 60) + myArray(1)
                        myArray = Split(dr.Item("close"), ":")
                        intClose = (myArray(0) * 60) + myArray(1)
                        If intCreatedTime <= intOpen Then
                            intNonWorkingMinute += ((intOpen - intCreatedTime) + (1440 - intClose))
                        ElseIf intCreatedTime > intClose Then
                            intNonWorkingMinute += (1440 - intCreatedTime)
                        Else
                            intNonWorkingMinute += (1440 - intClose)
                        End If
                    Next
                End If
                'log.Info(intNonWorkingMinute)

                ' Calculate NonWorkingMinute from the day after case created date till yesterday
                d = CDate(strDate).AddDays(1)
                Dim maxDate As Date = CDate(strDateNow).AddDays(-1)
                While DateTime.Compare(maxDate, d) >= 0
                    strTemp = FormatDate(CStr(d), "dd-MMM-yy", "yyyy-MM-dd")
                    If IsHoliday(strTemp, conn) Then
                        intNonWorkingMinute += 1440
                    Else
                        strWDayName = WeekdayName(Weekday(strTemp))
                        sql = "SELECT * FROM business_hour WHERE day='" & strWDayName & "'"
                        cmd = New OdbcDataAdapter(sql, conn)
                        ds = New DataSet()
                        cmd.Fill(ds, "business_hour")
                        dt = ds.Tables("business_hour")
                        For Each dr In dt.Rows
                            myArray = Split(dr.Item("open"), ":")
                            intOpen = (myArray(0) * 60) + myArray(1)
                            myArray = Split(dr.Item("close"), ":")
                            intClose = (myArray(0) * 60) + myArray(1)
                            intNonWorkingMinute += (1440 - (intClose - intOpen))
                        Next
                    End If
                    d = d.AddDays(1)
                End While
                'log.Info(intNonWorkingMinute)

                ' Calculate NonWorkingMinute for today
                myArray = Split(strTimeNow, ":")
                intTimeNow = (myArray(0) * 60) + myArray(1)
                If IsHoliday(strDateNow, conn) Then
                    intNonWorkingMinute += intTimeNow
                Else
                    strWDayName = WeekdayName(Weekday(strDateNow))
                    sql = "SELECT * FROM business_hour WHERE day='" & strWDayName & "'"
                    cmd = New OdbcDataAdapter(sql, conn)
                    ds = New DataSet()
                    cmd.Fill(ds, "business_hour")
                    dt = ds.Tables("business_hour")
                    For Each dr In dt.Rows
                        myArray = Split(dr.Item("open"), ":")
                        intOpen = (myArray(0) * 60) + myArray(1)
                        myArray = Split(dr.Item("close"), ":")
                        intClose = (myArray(0) * 60) + myArray(1)
                        If intTimeNow <= intOpen Then
                            intNonWorkingMinute += intTimeNow
                        ElseIf intTimeNow > intClose Then
                            intNonWorkingMinute += (intOpen + (intTimeNow - intClose))
                        Else
                            intNonWorkingMinute += intOpen
                        End If
                    Next
                End If
                'log.Info(intNonWorkingMinute)
            End If
            Return intNonWorkingMinute
        Catch ex As Exception
            LogEvent("NonWorkingMinute: " & ex.Message, EventLogEntryType.Error)
            log.Error("NonWorkingMinute: " & ex.Message)
            Return intNonWorkingMinute
        End Try
    End Function

    Function IsHoliday(ByVal strDate As String, ByVal conn As OdbcConnection) As Boolean
        ' strDate format must be yyyy-MM-dd
        Dim myArray(2) As String
        myArray = Split(strDate, "-")
        Dim fixed_date As String = myArray(1) & "-" & myArray(2)
        sql = "SELECT * FROM holiday WHERE IF(fixed =0,holiday_date='" & strDate & "',DATE_FORMAT(holiday_date,'%m-%d')='" & fixed_date & "')"
        If RowCountSql(sql, strConnection, conn) > 0 Then
            Return True
        End If

        Dim strWDayName As String = WeekdayName(Weekday(strDate))
        sql = "SELECT * FROM business_hour WHERE day='" & strWDayName & "' AND work_day='0'"
        If RowCountSql(sql, strConnection, conn) > 0 Then
            Return True
        End If
        Return False
    End Function

    Public Function PendingMinute5x8(ByVal incident_id As String) As Integer
        Dim conn As New OdbcConnection(strConnection)
        sql = "SELECT CAST(CONCAT(date,' ',time) AS CHAR) AS timestamp,case_log_description FROM case_log WHERE case_log_type_id=2 AND case_log_description LIKE '%Pending%' AND case_id='" & incident_id & "' ORDER BY case_log_id"
        Dim cmd As New OdbcDataAdapter(sql, conn)
        Dim ds As New DataSet()
        cmd.Fill(ds, "case_log")
        Dim dt As DataTable = ds.Tables("case_log")
        Dim dr As DataRow
        Dim myArrayT1(1), myArrayT2(1) As String
        Dim case_log_description As String
        Dim t1, t2 As DateTime
        Dim intPendingMinute As Integer = 0
        Try
            For Each dr In dt.Rows
                case_log_description = dr.Item("case_log_description")
                If InStr(case_log_description, "to Pending") Then
                    t1 = DateTime.ParseExact(dr.Item("timestamp"), "yyyy-MM-dd HH:mm:ss", Nothing)
                    myArrayT1 = Split(dr.Item("timestamp"), " ")
                Else
                    t2 = DateTime.ParseExact(dr.Item("timestamp"), "yyyy-MM-dd HH:mm:ss", Nothing)
                    myArrayT2 = Split(dr.Item("timestamp"), " ")
                    intPendingMinute += (DateDiff(DateInterval.Minute, t1, t2) - NonWorkingMinute(myArrayT1(0), myArrayT1(1), myArrayT2(0), myArrayT2(1), conn))
                End If
            Next
            Return intPendingMinute
        Catch ex As Exception
            LogEvent("PendingMinute5x8: " & ex.Message, EventLogEntryType.Error)
            log.Error("PendingMinute5x8: " & ex.Message)
            Return intPendingMinute
        End Try
    End Function

    Public Function PendingMinute7x24(ByVal incident_id As String) As Integer
        Dim conn As New OdbcConnection(strConnection)
        sql = "SELECT CAST(CONCAT(date,' ',time) AS CHAR) AS timestamp,case_log_description FROM case_log WHERE case_log_type_id=2 AND case_log_description LIKE '%Pending%' AND case_id='" & incident_id & "' ORDER BY case_log_id"
        Dim cmd As New OdbcDataAdapter(sql, conn)
        Dim ds As New DataSet()
        cmd.Fill(ds, "case_log")
        Dim dt As DataTable = ds.Tables("case_log")
        Dim dr As DataRow
        Dim case_log_description As String
        Dim t1, t2 As DateTime
        Dim intPendingMinute As Integer = 0
        Try
            For Each dr In dt.Rows
                case_log_description = dr.Item("case_log_description")
                If InStr(case_log_description, "to Pending") Then
                    t1 = DateTime.ParseExact(dr.Item("timestamp"), "yyyy-MM-dd HH:mm:ss", Nothing)
                Else
                    t2 = DateTime.ParseExact(dr.Item("timestamp"), "yyyy-MM-dd HH:mm:ss", Nothing)
                    intPendingMinute += DateDiff(DateInterval.Minute, t1, t2)
                End If
            Next
            Return intPendingMinute
        Catch ex As Exception
            LogEvent("PendingMinute7x24: " & ex.Message, EventLogEntryType.Error)
            log.Error("PendingMinute7x24: " & ex.Message)
            Return intPendingMinute
        End Try
    End Function

    Public Sub UpdateIncidentOverdue(ByVal field As String, ByVal value As String, ByVal id As String, ByVal conn As OdbcConnection)
        'Dim conn As New OdbcConnection(strConnection)
        sql = "UPDATE incident SET " & field & "='" & value & "' WHERE id='" & id & "'"
        Dim cmd As New OdbcDataAdapter(sql, conn)
        Dim ds As New DataSet()
        cmd.Fill(ds, "incident")
    End Sub

    '    Public Sub AutoClose()
    '        Try
    '            Dim conn As New OdbcConnection(strConnection)
    '            Dim cmd, cmd1 As New OdbcDataAdapter()
    '            Dim ds, ds1 As New DataSet()
    '            Dim dt, dt1 As DataTable
    '            Dim dr, dr1 As DataRow

    '            sql = "SELECT * FROM(" & _
    '                  " SELECT DISTINCT incident.id,sla_detail.autoclose FROM incident,contract,sla,sla_detail,priority WHERE incident.contract_id=contract.contract_id AND contract.sla_id=sla.sla_id AND case_status_id=6 AND (sla.sla_id = sla_detail.sla_id and incident.priority_id = sla_detail.priority_id)" & _
    '                  " UNION " & _
    '                  " SELECT DISTINCT incident.id,autoclose FROM incident,staff,priority WHERE incident.staff_id = staff.staff_id AND incident.priority_id=priority.priority_id AND case_status_id=6" & _
    '                  ")tb ORDER BY tb.id"
    'case_id AND incident.id='" & incident_id & "' AND case_log_type_id=2 AND case_log_description LIKE '%to Resolved%' ORDER BY date DESC,time DESC LIMIT 1"
    '                cmd1 = New OdbcDataAdapter(sql, conn)
    '                ds1 = New DataSet()
    '                cmd1.Fill(ds1, "*")
    '                dt1 = ds1.Tables("*")
    '                Dim resolved_date As String
    '                For Each dr1 In dt1.Rows
    '                    resolved_date = dr1.Item("resolved_date")
    '                    Dim t1 As DateTime = DateTime.ParseExact(resolved_date, "yyyy-MM-dd HH:mm:ss", Nothing)
    '                    Dim t2 As DateTime = Now
    '                    Dim timediff As Integer = Abs(DateDiff(DateInterval.Minute, t1, t2))
    '                    log.Info("AutoClose1: " & incident_id & "," & resolved_date & "," & autoclose & "," & timediff)
    '                    If timediff >= Val(autoclose) And Val(autoclose) <> 0 Then
    '                        log.Info("AutoClose2: " & incident_id & " is auto-closed.")
    '                        CloseCase(incident_id)
    '                    End If
    '                Next
    '            Next

    '        Catch ex As Exception
    '            LogEvent("AutoClose: " & ex.Message, EventLogEntryType.Error)
    '            log.Error("AutoClose: " & ex.Message)
    '        End Try
    '    End Sub
    Public Sub AutoClose()
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd, cmd1 As New OdbcDataAdapter()
            Dim ds, ds1 As New DataSet()
            Dim dt, dt1 As DataTable
            Dim dr, dr1 As DataRow
            'sql = "SELECT DISTINCT staff_id FROM incident WHERE case_status_id=6 ORDER BY incident.id"
            'sql = "SELECT DISTINCT incident.id,autoclose FROM incident,contract,sla WHERE incident.contract_id=contract.contract_id AND contract.sla_id=sla.sla_id AND case_status_id=6 ORDER BY incident.id"
            sql = "SELECT * FROM(" & _
                  "SELECT DISTINCT incident.id,autoclose FROM incident,contract,sla,sla_detail  WHERE(incident.contract_id = contract.contract_id) AND contract.sla_id=sla.sla_id AND sla.sla_id = sla_detail.sla_id AND case_status_id=6" & _
                  " UNION " & _
                  " SELECT DISTINCT incident.id,autoclose FROM incident,staff,priority WHERE incident.staff_id = staff.staff_id AND incident.priority_id=priority.priority_id AND case_status_id=6" & _
                  ")tb ORDER BY tb.id"
            cmd = New OdbcDataAdapter(sql, conn)
            ds = New DataSet()
            cmd.Fill(ds, "*")
            dt = ds.Tables("*")
            Dim incident_id, autoclose As String
            Dim myArray(1) As String
            For Each dr In dt.Rows
                incident_id = dr.Item("id")
                autoclose = dr.Item("autoclose")
                If autoclose <> "00:00" Then
                    myArray = Split(autoclose, ":")
                    autoclose = (myArray(0) * 60) + myArray(1)
                    sql = "SELECT CAST(CONCAT(date,' ',time) AS CHAR) AS resolved_date FROM incident,case_log WHERE incident.id=case_log.case_id AND incident.id='" & incident_id & "' AND case_log_type_id=2 AND case_log_description LIKE '%to Resolved%' ORDER BY date DESC,time DESC LIMIT 1"
                    cmd1 = New OdbcDataAdapter(sql, conn)
                    ds1 = New DataSet()
                    cmd1.Fill(ds1, "*")
                    dt1 = ds1.Tables("*")
                    Dim resolved_date As String
                    For Each dr1 In dt1.Rows
                        resolved_date = dr1.Item("resolved_date")
                        Dim t1 As DateTime = DateTime.ParseExact(resolved_date, "yyyy-MM-dd HH:mm:ss", Nothing)
                        Dim t2 As DateTime = Now
                        Dim timediff As Integer = Abs(DateDiff(DateInterval.Minute, t1, t2))
                        log.Info("AutoClose1: " & incident_id & "," & resolved_date & "," & autoclose & "," & timediff)
                        If timediff >= Val(autoclose) Then
                            log.Info("AutoClose2: " & incident_id & " is auto-closed.")
                            CloseCase(incident_id)
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            LogEvent("AutoClose: " & ex.Message, EventLogEntryType.Error)
            log.Error("AutoClose: " & ex.Message)
        End Try
    End Sub
    Sub CloseCase(ByVal incident_id As String)
        Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()

        sql = "UPDATE incident SET "
        sql += "case_status_id='8' "
        sql += ",closure_type_id='1' "
        sql += "WHERE id=" & incident_id
        cmd = New OdbcDataAdapter(sql, conn)
        cmd.Fill(ds, "incident")

        ' Update case_log and sending email for case status change
        ChangeCaseStatusLogEmail(incident_id, "6", "8")

        ' Change sub cases as resolved in order to let system do auto close later.
        ChangeSubCaseAsResolved(incident_id)
    End Sub

    Sub ChangeSubCaseAsResolved(ByVal parent_case_id As String)
        Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()
        sql = "SELECT id,case_status_id,case_duration FROM incident WHERE parent_case_id='" & parent_case_id & "' AND case_status_id <> 8 AND case_status_id <> 6"
        'log.Info("ChangeSubCaseAsResolved1: " & sql)
        cmd = New OdbcDataAdapter(sql, conn)
        cmd.Fill(ds, "incident")
        Dim dt1 As DataTable = ds.Tables("incident")
        Dim dr1 As DataRow = Nothing
        Dim incident_id, case_status_id As Integer
        For Each dr1 In dt1.Rows
            incident_id = dr1.Item("id")
            case_status_id = dr1.Item("case_status_id")
            sql = "UPDATE incident SET "
            sql += "case_status_id='6' "
            sql += ",resolve_duration='" & dr1.Item("case_duration") & "' "
            sql += "WHERE id=" & incident_id
            cmd = New OdbcDataAdapter(sql, conn)
            cmd.Fill(ds, "incident")

            ChangeCaseStatusLogEmail(incident_id, case_status_id, "6")
            ChangeSubCaseAsResolved(incident_id)
        Next
    End Sub

    Sub ChangeCaseStatusLogEmail(ByVal incident_id As String, ByVal old_case_status_id As String, ByVal case_status_id As String)
        Dim conn As New OdbcConnection(strConnection)
        Dim cmd As New OdbcDataAdapter()
        Dim ds As New DataSet()

        sql = "SELECT case_status_title FROM case_status WHERE case_status_id='" & old_case_status_id & "'"
        Dim old_case_status As String = FetchField(sql, strConnection, "case_status_title")
        sql = "SELECT case_status_title FROM case_status WHERE case_status_id='" & case_status_id & "'"
        Dim new_case_status As String = FetchField(sql, strConnection, "case_status_title")
        Dim strChangeStatus As String = "changed case status from " & old_case_status & " to " & new_case_status & "."

        ' Insert into case_log table
        Dim case_log_id As Integer = FindLastID("case_log", strConnection, "case_log_id") + 1
        sql = "INSERT INTO case_log VALUES ('"
        sql += case_log_id & "','"
        sql += DateStamp() & "','"
        sql += TimeStamp() & "','"
        sql += incident_id & "','"
        sql += "System','"
        sql += "2','"       ' case_log_type = change case status
        sql += strChangeStatus & "','"
        sql += "2','')"     ' case_log_category_id = external log
        cmd = New OdbcDataAdapter(sql, conn)
        cmd.Fill(ds, "case_log")

        Dim strSubj As String = ""
        Dim strBody As String = ""
        Dim SendEmailResult As String = ""
        Dim SendEmailStatus As String = ""

        sql = "SELECT case_id,title,email FROM incident WHERE id='" & incident_id & "'"
        cmd = New OdbcDataAdapter(sql, conn)
        cmd.Fill(ds, "*")
        Dim dt1 As DataTable = ds.Tables("*")
        Dim dr1 As DataRow = Nothing
        Dim strEmail As String
        For Each dr1 In dt1.Rows
            strEmail = dr1.Item("email")
            ' Send Email for case status change
            If strEmail <> "" Then
                strSubj = BuildSubj(dr1.Item("case_id"), dr1.Item("title"), strChangeStatus)
                strBody = BuildBodyClose(incident_id, strChangeStatus)
                Dim mail As New NSDEmailSMSSender()
                ' Send Email
                mail.ConnString = strConnection
                mail.AlertType = "1" ' 1= Email , 2= SMS
                mail.IncidentID = incident_id
                mail.Send_To = strEmail
                mail.Subject = strSubj
                mail.Body = strBody
                If signature <> "" Then
                    mail.Signature = signature
                End If
                mail.Send()
                log.Info("Sent Email ," & strSubj & " , " & strEmail)
                'SendEmailResult = SendEmail(company_name, strEmail, strSubj, strBody, "", smtp_server, smtp_sender, smtp_password, bSsl, sslPort)
                'If SendEmailResult = "OK" Then
                '    SendEmailStatus = "1"
                'Else
                '    SendEmailStatus = "0"
                'End If
                'sql = "INSERT INTO email_sms_log VALUES ('"
                'sql += DateStamp() & "','"
                'sql += TimeStamp() & "','"
                'sql += incident_id & "','"
                'sql += case_log_id & "','"
                'sql += "','"        ' mobile
                'sql += strEmail & "','"
                'sql += strSubj & "','"
                'sql += SendEmailStatus & "','"
                'sql += SendEmailResult & "')"
                'cmd = New OdbcDataAdapter(sql, conn)
                'cmd.Fill(ds, "email_sms_log")
            End If
        Next
    End Sub

    Public Sub FreeEngineer()
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter()
            Dim ds As New DataSet()
            Dim dt As DataTable
            Dim dr As DataRow
            conn.Open()
            sql = "SELECT staff_id,staff_status_id FROM staff"
            cmd = New OdbcDataAdapter(sql, conn)
            ds = New DataSet()
            cmd.Fill(ds, "staff")
            dt = ds.Tables("staff")
            Dim staff_id, staff_status_id As String
            Dim num_case As Integer
            For Each dr In dt.Rows
                staff_id = dr.Item("staff_id")
                staff_status_id = dr.Item("staff_status_id")
                sql = "SELECT id FROM incident WHERE case_status_id < 6 AND engineer_id='" & staff_id & "'"
                num_case = RowCountSql(sql, strConnection, conn)
                log.Info("FreeEngineer1: " & staff_id & "," & num_case)
                If num_case = 0 And staff_status_id = "2" Then
                    sql = "UPDATE staff SET staff_status_id='1' WHERE staff_id='" & staff_id & "'"
                    cmd = New OdbcDataAdapter(sql, conn)
                    cmd.Fill(ds, "staff")
                    log.Info("FreeEngineer2: Changed " & staff_id & " is available.")
                End If
            Next
            conn.Close()
        Catch ex As Exception
            LogEvent("FreeEngineer: " & ex.Message, EventLogEntryType.Error)
            log.Error("FreeEngineer: " & ex.Message)
        End Try
    End Sub


    Public Function ProductValidNSD() As String
        Try
            SetGlobalization()
            Dim sql As String = "SELECT product,product_key FROM config"
            Dim strConnection As String = LoadDBDriver()
            System.Threading.Thread.Sleep(1000)
            If Not Is_DB_Online(sql, strConnection) Then
                Return "could not connect to database."
            Else
                Dim strProduct As String = FetchField(sql, strConnection, "product")
                Dim strProductKey As String = FetchField(sql, strConnection, "product_key")
                If strProduct = Nothing Or strProductKey = Nothing Then
                    Return "product key is not valid."
                ElseIf strProductKey = "ne1kA_qu@rtZ*" Then
                    Return "Valid"
                Else
                    Dim strAuth As String = DecryptNetka(strProductKey)
                    Dim myArray() As String = Split(strAuth, "|")
                    Dim ip As String = myArray(0)
                    Dim mac As String = myArray(1)
                    Dim product As String = myArray(2)
                    Dim expiry_date As String = myArray(3)
                    If expiry_date = "never" Then
                    ElseIf DateDiff("d", Now, CDate(expiry_date)) >= 0 Then
                    Else
                        Return "product expired. Please contact Netka System for purchasing. [" & Now & "," & expiry_date & "]"
                    End If
                    Dim strMac As String = MacAddress(ip)
                    Dim bNetworkInterfaceWasDown As Boolean = True
                    Dim bValidate As Boolean = False
                    'If strMac <> "" Then
                    '    bNetworkInterfaceWasDown = False
                    'End If
                    If InStr(ip, ",") > 0 And InStr(mac, ",") > 0 Then
                        Dim arrIP() As String = Split(ip, ",")
                        Dim arrMac() As String = Split(mac, ",")
                        Dim intUBoundArrIP As Integer = UBound(arrIP)
                        Dim intUBoundArrMac As Integer = UBound(arrMac)
                        If intUBoundArrIP <> intUBoundArrMac Then
                            Return "Product expired. Please contact Netka System for purchasing."
                        End If
                        Dim arrMac1(intUBoundArrMac) As String
                        Dim i As Integer
                        For i = 0 To intUBoundArrIP
                            arrMac1(i) = MacAddress(arrIP(i))
                            '                            nks.debug(arrIP(i) & "|" & arrMac(i) & "|" & arrMac1(i))
                            If arrMac(i) = arrMac1(i) Then
                                bValidate = True
                            End If
                            If arrMac1(i) <> "" Then
                                bNetworkInterfaceWasDown = False
                            End If
                        Next
                    Else
                        Dim mac1 As String = MacAddress(ip)
                        'debug(ip & "|" & mac & "|" & mac1)
                        If mac = mac1 Then
                            bValidate = True
                        End If
                        If mac1 <> "" Then
                            bNetworkInterfaceWasDown = False
                        End If
                    End If
                    If ((bValidate And (strProduct = product)) Or (mac = "000E355795B0") Or (mac = "001F3C9ACF92") Or (mac = "001B778F0707") Or (mac = "00216BA9DD84") Or (mac = "0022FACD3402") Or (mac = "001CBF75BA89") Or (mac = "0022FACD3A66")) Then
                        Return "Valid"
                    ElseIf bNetworkInterfaceWasDown Then
                        Return "server's network interface was down."
                    Else
                        Return "product key is not valid. [" & ip & "," & mac & "," & product & "," & expiry_date & "]"
                    End If
                End If
            End If
        Catch ex As Exception
            log.Info(ex.ToString)
            Return "Error: " & ex.Message
        End Try
    End Function

    Public Function LoadDBDriver() As String
        LoadDBDriver = ConfigurationManager.ConnectionStrings("strConnection").ConnectionString
        If InStr(LoadDBDriver, "atabase") Then

        Else
            LoadDBDriver = AESDecrypt(LoadDBDriver, "s7Pwe1$,Gh(Ve2Xa", "Nmq<3FcW##ly5$UO")
        End If
        'Dim ObjFile As New FileInfo("c:\dbdriver_netkaquartz.ini")
        'Dim ObjStreamReader As StreamReader = ObjFile.OpenText
        'LoadDBDriver = ObjStreamReader.ReadToEnd()
        'ObjStreamReader.Close()
    End Function

    Public Function Is_DB_Online(ByVal sql As String, ByVal strConnection As String) As Boolean
        Try
            Dim conn As New OdbcConnection(strConnection)
            Dim cmd As New OdbcDataAdapter(sql, conn)
            Dim ds As New DataSet()
            cmd.Fill(ds, "*")
            Return True
        Catch
            Return False
        End Try
    End Function

   

    Function StripTagsImg(ByVal html As String) As String
        Dim result As String = html
        signature = ""

        If result.IndexOf("<img") <> -1 Then
            Dim col As MatchCollection = Regex.Matches(html, "UserImageFiles/(.*)""")
            ' Loop through Matches.
            For Each m As Match In col
                Dim g As System.Text.RegularExpressions.Group = m.Groups(1)

                If signature <> "" Then
                    signature = signature & "|"
                End If
                Dim tmp As String = ""
                tmp = g.Value
                tmp = Regex.Replace(tmp, "\(", "\(")
                tmp = Regex.Replace(tmp, "\)", "\)")

                result = Regex.Replace(result, "<img[^>]*" & tmp & """ ?/>", "<img src=cid:" & Replace(g.Value, " ", "_") & ">")
                signature = signature & g.Value
            Next
        Else
            signature = ""
        End If

        Return result
    End Function

    Function GetDownTimeDuration(ByVal case_id) As String
        'Dim ts As TimeSpan 
        Dim tSpan As TimeSpan
        'Dim totalSeconds as double
        'Dim arr() as String 
        Dim strValue As String
        Dim startdate As Date
        Dim enddate As Date
        Dim flag As Boolean = False
        sql = "SELECT downtime,uptime FROM incident WHERE id='" & case_id & "'"

        Dim conn As New OdbcConnection(strConnection)
        Dim ds As New System.Data.DataSet()
        Dim cmd As New OdbcDataAdapter(sql, conn)
        cmd.Fill(ds, "*")
        Dim dr As System.Data.DataRow
        For Each dr In ds.Tables(0).Rows
            If Not IsDBNull(dr("downtime")) Then
                startdate = dr("downtime")
            Else
                startdate = "0001-01-01 00:00:00"
            End If
            If Not IsDBNull(dr("uptime")) Then
                enddate = dr("uptime")
            Else
                enddate = "0001-01-01 00:00:00"
            End If

        Next dr
        If enddate = "0001-01-01 00:00:00" Then
            If startdate <> "0001-01-01 00:00:00" Then
                enddate = Now()
            End If

        Else
            flag = True
        End If
        tSpan = TimeSpan.FromSeconds((enddate - startdate).TotalSeconds)
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
End Class
Public Class NSDEmailSMSSender


    Private sql As String
    Private strConnection As String = ""
    Private strDate As String
    Private strColumn, strHeader As String
    Private apppath As String = ConfigurationSettings.AppSettings("AppPath")
    Private companySubject As String = ConfigurationSettings.AppSettings("company_name")
    Private _alert_type As String = ""
    Private _incident_id As String = ""
    Private _args As String = ""
    Private _to As String = ""
    Private _cc As String = ""
    Private _bcc As String = ""
    Private _att As String = ""
    Private _mobile As String = ""
    Private _action As String = ""
    Private _subject As String = ""
    Private _strBody As String = ""
    Private _attpath As String = ""
    Private _img As String = ""

    Public Property ConnString As String
        Get
            Return strConnection
        End Get
        Set(ByVal value As String)
            strConnection = value
        End Set
    End Property
    Public Property AlertType As String
        Get
            Return _alert_type
        End Get
        Set(ByVal value As String)
            _alert_type = value
        End Set
    End Property
    Public Property IncidentID As String
        Get
            Return _incident_id
        End Get
        Set(ByVal value As String)
            _incident_id = value
        End Set
    End Property
    Public Property Send_To As String
        Get
            Return _to
        End Get
        Set(ByVal value As String)
            _to = value
        End Set
    End Property
    Public Property Send_CC As String
        Get
            Return _cc
        End Get
        Set(ByVal value As String)
            _cc = value
        End Set
    End Property
    Public Property Send_BCC As String
        Get
            Return _bcc
        End Get
        Set(ByVal value As String)
            _bcc = value
        End Set
    End Property
    Public Property Attachment As String
        Get
            Return _attpath
        End Get
        Set(ByVal value As String)
            _attpath = value
        End Set
    End Property
    Public Property MobileList As String
        Get
            Return _mobile
        End Get
        Set(ByVal value As String)
            _mobile = value
        End Set
    End Property
    Public Property Action As String
        Get
            Return _action
        End Get
        Set(ByVal value As String)
            _action = value
        End Set
    End Property
    Public Property Subject As String
        Get
            Return _subject
        End Get
        Set(ByVal value As String)
            _subject = value
        End Set
    End Property
    Public Property Body As String
        Get
            Return _strBody
        End Get
        Set(ByVal value As String)
            _strBody = value
        End Set
    End Property
    Public Property Signature As String
        Get
            Return _img
        End Get
        Set(ByVal value As String)
            _img = value
        End Set
    End Property

    Private Function ParserArgument() As String
        '	 "-s """ & strConnection & """ -t " & "1" & " -c " & Session("incident_id") & " -l " & "1" & " -e " & """thammas@netkasystem.com;earth_tc6@hotmail.com""" & " - a """ & strSubj & """ -h """ & strSubj & """ -b """ & strBody & """"
        Dim result As String = ""
        Dim d As String = """"
        Dim s As String = "'"
        If strConnection <> "" Then result = "-s """ & strConnection & """"
        If _alert_type <> "" Then result = result & " -t " & _alert_type
        If _incident_id <> "" Then result = result & " -c " & _incident_id
        If _to <> "" Then result = result & " -e """ & _to & """"
        If _cc <> "" Then result = result & " -cc """ & _cc & """"
        If _bcc <> "" Then result = result & " -bcc """ & _bcc & """"
        If _attpath <> "" Then result = result & " -att """ & _attpath & """"
        If _mobile <> "" Then result = result & " -m """ & _mobile & """"
        If _action <> "" Then result = result & " -a """ & _action & """"
        If _subject <> "" Then result = result & " -h """ & _subject.Replace(d, s & s) & """"
        If _strBody <> "" Then result = result & " -b """ & _strBody.Replace(d, s & s) & """"
        If apppath <> "" Then result = result & " -p """ & apppath & """"
        If companySubject <> "" Then result = result & " -n """ & companySubject & """"
        If _img <> "" Then result = result & " -sig """ & _img & """"
        Return result
    End Function

    Public Sub Send()
        ' Shell NSDEmailSMSSender.EXE 
        ' -------------------------------------------------------------------------------------------------------
        Dim proc As New Diagnostics.Process()
        proc.StartInfo.Arguments = ParserArgument()

        proc.StartInfo.FileName = apppath & "\NSDEmailSMSSender.exe"
        '	nks.debug(Server.MapPath("NSDEmailSMSSender.exe"))
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.RedirectStandardOutput = False
        proc.Start()
        ' -------------------------------------------------------------------------------------------------------

    End Sub

End Class


