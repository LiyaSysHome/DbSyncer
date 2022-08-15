Imports System.IO
Imports System.Data
Imports System.Data.SqlClient





Public Class Syncer

    Dim DATE_UPDATE_STATUS As Boolean = True

    Public Shared Function GetConnection_Remote_Sql_String() As String
        Dim constr As String

        constr = "Server=liyasys-pc\sql_2016;Database=Liyasys_Syncer;user id =sa; password =123456;Integrated Security=False;Connect Timeout=60"

        GetConnection_Remote_Sql_String = constr
    End Function

    Public Shared Function HaveInternetConnection() As Boolean

        Try
            Return My.Computer.Network.Ping("www.google.com")
        Catch
            Return False
        End Try

    End Function
    Private Sub START()

      

        Get_Actions_from_Process_File()

        If Common_Procedures.CREATE_DATABASE_STATUS = True Then
            pnl_1.Visible = True
            pnl_2.Visible = False
            pnl_3.Visible = False
            pnl_4.Visible = False

            GET_TABLES_FROM_DATABASE()


        End If

        If Common_Procedures.FIELD_CHECK_STATUS = True Then
            pnl_1.Visible = False
            pnl_2.Visible = True
            pnl_3.Visible = False
            pnl_4.Visible = False

            GET_EXTRA_FIELDS_FROM_DATABASE()
        End If

        If Common_Procedures.NEW_DATA_STATUS = True Then
            pnl_1.Visible = False
            pnl_2.Visible = False
            pnl_3.Visible = False
            pnl_4.Visible = True
            GET_DATA_FROM_TRIGGER()
        End If


        If Common_Procedures.FORM_VISIBLE_STATUS = False Then

            Me.Hide()
            Me.Width = 140
            Me.Height = 25
        Else
            Me.Show()
            Me.Width = 140
            Me.Height = 25
        End If

    End Sub

    Private Sub btn_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Load.Click

        START()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        START()

    End Sub

    Private Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        If Common_Procedures.OLD_DATA_STATUS = True And DATE_UPDATE_STATUS = True Then
            DATE_UPDATE_STATUS = False

            pnl_1.Visible = False
            pnl_2.Visible = False
            pnl_3.Visible = True
            pnl_4.Visible = False

            GET_OLD_DATA_FROM_OLD_DATABASE()
        End If

    End Sub


    Private Sub Get_ServerDetails()

        Dim pth As String, ConStr As String
        Dim a() As String
        Dim fs As FileStream
        Dim r As StreamReader
        Dim w As StreamWriter

        Try

            If InStr(1, Trim(LCase(Application.StartupPath)), "\bin\debug") > 0 Then
                Common_Procedures.AppPath = Replace(Trim(LCase(Application.StartupPath)), "\bin\debug", "")
            Else
                Common_Procedures.AppPath = Application.StartupPath
            End If

            pth = Trim(Common_Procedures.AppPath) & "\connection.ini"

            Common_Procedures.ServerName = ""
            Common_Procedures.ServerPassword = ""
            Common_Procedures.ServerWindowsLogin = ""
            Common_Procedures.ServerDataBaseLocation_InExTernalUSB = ""

            Common_Procedures.ConnectionString_CompanyGroupdetails = ""
            Common_Procedures.ConnectionString_Master = ""
            Common_Procedures.Connection_String = ""
            Common_Procedures.DataBaseName = ""

            If File.Exists(pth) = False Then
                fs = New FileStream(pth, FileMode.Create)
                w = New StreamWriter(fs)
                w.WriteLine(SystemInformation.ComputerName & "\sql_2016,liya467")
                w.Close()
                fs.Close()
                w.Dispose()
                fs.Dispose()
            End If

            ConStr = ""
            If File.Exists(pth) = True Then
                fs = New FileStream(pth, FileMode.Open)
                r = New StreamReader(fs)
                ConStr = r.ReadLine
                r.Close()
                fs.Close()
                r.Dispose()
                fs.Dispose()
            End If

            If Trim(ConStr) <> "" Then
                a = Split(ConStr, ",")
                If UBound(a) >= 0 Then Common_Procedures.ServerName = Trim(a(0))
                If UBound(a) >= 1 Then Common_Procedures.ServerPassword = Trim(a(1))
                If UBound(a) >= 2 Then Common_Procedures.ServerWindowsLogin = Trim(a(2))
                If UBound(a) >= 3 Then
                    If Trim(a(3)) <> "" Then
                        If InStr(1, Trim(UCase(a(3))), "PLIT") > 0 And InStr(1, Trim(UCase(a(3))), "COMPANYGROUP") > 0 Then
                            Common_Procedures.CompanyDetailsDataBaseName = Trim(a(3))
                        End If
                    End If
                End If
                If UBound(a) >= 4 Then Common_Procedures.ServerDataBaseLocation_InExTernalUSB = Trim(a(4))
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "DOES NOT OPEN...", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub Get_Actions_from_Process_File()

        Dim pth As String, ConStr As String
        Dim a() As String
        Dim fs As FileStream
        Dim r As StreamReader
        Dim w As StreamWriter

        Try

            If InStr(1, Trim(LCase(Application.StartupPath)), "\bin\debug") > 0 Then
                Common_Procedures.AppPath = Replace(Trim(LCase(Application.StartupPath)), "\bin\debug", "")
            Else
                Common_Procedures.AppPath = Application.StartupPath
            End If

            pth = Trim(Common_Procedures.AppPath) & "\process.ini"

            Common_Procedures.DATABASE_NAME = ""
            Common_Procedures.CREATE_DATABASE_STATUS = False
            Common_Procedures.FIELD_CHECK_STATUS = False
            Common_Procedures.OLD_DATA_STATUS = False
            Common_Procedures.NEW_DATA_STATUS = False
            Common_Procedures.FORM_VISIBLE_STATUS = False

            If File.Exists(pth) = False Then
                fs = New FileStream(pth, FileMode.Create)
                w = New StreamWriter(fs)
                w.WriteLine("PLIT_BILLING_1,YES,UNSHOW")
                w.Close()
                fs.Close()
                w.Dispose()
                fs.Dispose()
            End If

            ConStr = ""
            If File.Exists(pth) = True Then
                fs = New FileStream(pth, FileMode.Open)
                r = New StreamReader(fs)
                ConStr = r.ReadLine
                r.Close()
                fs.Close()
                r.Dispose()
                fs.Dispose()
            End If

            If Trim(ConStr) <> "" Then
                a = Split(ConStr, ",")

                If UBound(a) >= 0 Then Common_Procedures.DATABASE_NAME = Trim(a(0))
                'If UBound(a) >= 1 Then Common_Procedures.CREATE_DATABASE_STATUS = IIf(Trim(a(1)) = "YES", True, False)
                'If UBound(a) >= 2 Then Common_Procedures.FIELD_CHECK_STATUS = IIf(Trim(a(2)) = "YES", True, False)
                'If UBound(a) >= 3 Then Common_Procedures.OLD_DATA_STATUS = IIf(Trim(a(3)) = "YES", True, False)
                If UBound(a) >= 1 Then Common_Procedures.NEW_DATA_STATUS = IIf(Trim(a(1)) = "YES", True, False)
                If UBound(a) >= 2 Then Common_Procedures.FORM_VISIBLE_STATUS = IIf(Trim(a(2)) = "SHOW", True, False)

                'If UBound(a) >= 4 Then Common_Procedures.ServerDataBaseLocation_InExTernalUSB = Trim(a(4))
            End If




        Catch ex As Exception
            MessageBox.Show(ex.Message, "DOES NOT OPEN...", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Syncer_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Dim cn1 As SqlClient.SqlConnection
        Dim da2 As New SqlClient.SqlDataAdapter
        Dim dt2 As New DataTable
        Dim n As Integer = 0
        Dim CompgrpCondt As String = ""
        Dim DBName As String = ""
        Dim ShowSTS As Boolean = False



        cn1 = New SqlClient.SqlConnection(Common_Procedures.ConnectionString_CompanyGroupdetails)

        cn1.Open()

        GET_CONNECTION()

        Timer1.Enabled = True

    End Sub
    Private Sub GET_CONNECTION()
        Dim cmd_mysql As New SqlClient.SqlCommand
        Dim cmd As New SqlClient.SqlCommand
        Dim IdNo As Integer
        Dim da As New SqlClient.SqlDataAdapter
        Dim dt As New DataTable
        Dim da1 As New SqlClient.SqlDataAdapter
        Dim dt1 As New DataTable
        Dim query As String = ""

        Try


            IdNo = 1
CHANGE_ID:
            Common_Procedures.CompGroupIdNo = 0
            Common_Procedures.CompGroupName = ""
            Common_Procedures.CompGroupFnRange = ""

            Common_Procedures.Connection_String = ""
            Common_Procedures.DataBaseName = ""

            If Val(IdNo) <> 0 Then


                Common_Procedures.CompGroupIdNo = Val(IdNo)


                Common_Procedures.DataBaseName = Common_Procedures.DATABASE_NAME


                Common_Procedures.Connection_String = Common_Procedures.Create_Sql_ConnectionString(Common_Procedures.DataBaseName)

            Else
                IdNo = IdNo + 1
                GoTo CHANGE_ID

                'MessageBox.Show("Set Database ID ", "INVALID DATABASE ID...", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error in Get trigger details", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub



    Private Sub Syncer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        If Not My.Computer.Network.IsAvailable Then
            Me.Close()
        End If

        Me.Width = 140
        Me.Height = 25

        Me.Location() = New Point(Screen.PrimaryScreen.WorkingArea.Width - Me.Width, Screen.PrimaryScreen.WorkingArea.Top) ' Me.Height)

        Get_Actions_from_Process_File()



        lbl_message_1.Text = ""
        lbl_message_2.Text = ""
        lbl_message_3.Text = ""
        lbl_message_4.Text = ""


        Get_ServerDetails()



        If Trim(Common_Procedures.ServerName) = "" Then
            MessageBox.Show("Invalid Connection File Details", "INVALID SERVER DETAILS...", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End If



        Common_Procedures.ConnectionString_Master = ""
        Common_Procedures.ConnectionString_CompanyGroupdetails = ""
        Common_Procedures.Connection_String = ""

        Common_Procedures.ConnectionString_Master = Common_Procedures.Create_Sql_ConnectionString("master")
        Common_Procedures.ConnectionString_CompanyGroupdetails = Common_Procedures.Create_Sql_ConnectionString(Common_Procedures.CompanyDetailsDataBaseName)




    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Application.Exit()
    End Sub


    Private Sub GET_DATA_FROM_TRIGGER()
        Dim cn1 As SqlClient.SqlConnection
        Dim cmd As New SqlClient.SqlCommand

        Dim con_my As SqlClient.SqlConnection
        Dim cmd_mysql As New SqlCommand
        Dim trans_mysql As SqlTransaction

        Dim Nr As Integer

        Dim da As New SqlClient.SqlDataAdapter
        Dim dt As New DataTable

        Dim da1 As New SqlClient.SqlDataAdapter
        Dim dt1 As New DataTable
        Dim query As String = ""
        Dim Column_DataType As String = ""
        Dim Null_sts As Boolean = False
        Dim Date1_sts As Boolean = False
        Dim Date1 As DateTime

        ProgressBar4.Value = 0
        Try

            cn1 = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            cn1.Open()

            con_my = New SqlConnection(GetConnection_Remote_Sql_String)
            con_my.Open()


            trans_mysql = con_my.BeginTransaction
            cmd_mysql.Transaction = trans_mysql


            da = New SqlClient.SqlDataAdapter("select * from Trigger_Head where TABLE_NAME in (select tablename from Syncer_Table_Details where SyncStatus = 1) ORDER BY AUTO_SLNO asc ", cn1)
            dt = New DataTable
            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                For I = 0 To dt.Rows.Count - 1

                    Dim Primary_Key1 As String = Replace(Trim(dt.Rows(I)("Primary_key_1").ToString), "~", "'")
                    Dim Primary_Key2 As String = Replace(Trim(dt.Rows(I)("Primary_key_2").ToString), "~", "'")
                    Dim Primary_Key3 As String = Replace(Trim(dt.Rows(I)("Primary_key_3").ToString), "~", "'")


                    'check primary key is date
                    If InStr(Trim(UCase(Primary_Key1)), "DATE") > 0 Then
                        Date1_sts = True
                    Else
                        Date1_sts = False
                    End If


                    If Trim(dt.Rows(I)("Primary_Key_1_Value").ToString) <> "" Then
                        If Date1_sts = True Then
                            Primary_Key1 = Primary_Key1 & "=@Date1"

                            Dim temp1 As String = Replace(Trim(dt.Rows(I)("Primary_Key_1_Value").ToString), "~", "")

                            Date1 = Convert.ToDateTime(temp1)
                        Else
                            Primary_Key1 = Primary_Key1 & "=" & Replace(Trim(dt.Rows(I)("Primary_Key_1_Value").ToString), "~", "'")
                        End If

                    End If





                    If UCase(Trim(dt.Rows(I)("Command").ToString)) = "DELETE" Then

                        query = " delete from " & Trim(dt.Rows(I)("TABLE_NAME").ToString) & "  " & IIf(Trim(Primary_Key1) <> "", " WHERE " & Trim(Primary_Key1), " ") & IIf(Trim(Primary_Key2) <> "", " AND " & Trim(Primary_Key2), " ") & IIf(Trim(Primary_Key3) <> "", " AND " & Trim(Primary_Key3), " ")

                        cmd_mysql.Connection = con_my

                        cmd_mysql.CommandType = CommandType.Text

                        If Date1_sts = True Then

                            If IsDate(Date1) Then
                                cmd_mysql.Parameters.Clear()
                                cmd_mysql.Parameters.AddWithValue("@Date1", Date1)
                            End If
                        End If

                        cmd_mysql.CommandText = query
                        Nr = cmd_mysql.ExecuteNonQuery()

                        ' If Nr > 0 Then
                        cmd.Connection = cn1
                        cmd.CommandText = "delete from Trigger_Head where Command = '" & Trim(dt.Rows(I)("Command").ToString) & "' AND AUTO_SLNO = " & Str(Val(dt.Rows(I)("AUTO_SLNO").ToString))
                        cmd.ExecuteNonQuery()
                        'Else
                        ' Continue For
                        ' End If


                    ElseIf UCase(Trim(dt.Rows(I)("Command").ToString)) = "INSERT" Then

                        Dim DateStringArray(50) As String
                        Dim DateArray(50) As DateTime

                        query = "INSERT INTO "

                        Dim qry As String = "select * from " & Trim(dt.Rows(I)("Table_Name").ToString) & IIf(Trim(Primary_Key1) <> "", " WHERE " & Trim(Primary_Key1), " ") & IIf(Trim(Primary_Key2) <> "", " AND " & Trim(Primary_Key2), " ") & IIf(Trim(Primary_Key3) <> "", " AND " & Trim(Primary_Key3), " ")

                        cmd.Connection = cn1

                        cmd.CommandText = qry
                        da1 = New SqlClient.SqlDataAdapter()
                        da1.SelectCommand = cmd

                        If Date1_sts = True Then
                            If IsDate(Date1) Then
                                cmd.Parameters.Clear()
                                cmd.Parameters.AddWithValue("@Date1", Date1)
                            End If
                        End If

                        'da1 = New SqlClient.SqlDataAdapter("select * from " & Trim(dt.Rows(I)("Table_Name").ToString) & IIf(Trim(Primary_Key1) <> "", " WHERE " & Trim(Primary_Key1), " ") & IIf(Trim(Primary_Key2) <> "", " AND " & Trim(Primary_Key2), " ") & IIf(Trim(Primary_Key3) <> "", " AND " & Trim(Primary_Key3), " "), cn1)
                        dt1 = New DataTable
                        da1.Fill(dt1)

                        If dt1.Rows.Count > 0 Then
                            'ROWS
                            For j = 0 To dt1.Rows.Count - 1

                                'COLUMN
                                If dt1.Columns.Count > 0 Then

                                    query = query & "" & " " & Trim(dt.Rows(I)("TABLE_NAME").ToString) & " ( "    '  Trim(dt.Rows(I)("Table_Name").ToString) & " ( "
                                    'LOOP FIELDS
                                    For clm = 0 To dt1.Columns.Count - 1

                                        Column_DataType = UCase(Get_Column_dataType(Trim(dt.Rows(I)("Table_Name").ToString), Trim(dt1.Columns(clm).ColumnName)))

                                        If Column_DataType = "TIMESTAMP" Then

                                            If clm = Val(dt1.Columns.Count - 1) Then
                                                query = query & " ) "
                                            End If

                                            Continue For

                                        End If

                                        If clm = Val(dt1.Columns.Count - 1) Then
                                            query = query & Trim(dt1.Columns(clm).ColumnName) & " ) "
                                        Else
                                            query = query & Trim(dt1.Columns(clm).ColumnName) & " , "
                                        End If

                                    Next clm

                                    query = query & " VALUES ( "

                                    'LOOP FIELDS VALUES
                                    For clm = 0 To dt1.Columns.Count - 1

                                        Column_DataType = UCase(Get_Column_dataType(Trim(dt.Rows(I)("Table_Name").ToString), Trim(dt1.Columns(clm).ColumnName)))

                                        Dim STR As String = Trim(dt1.Rows(j)(clm).ToString)

                                        If STR = Nothing Then
                                            Null_sts = True
                                        Else
                                            Null_sts = False
                                        End If





                                        If Column_DataType = "VARCHAR" Or Column_DataType = "NVARCHAR" Or Column_DataType = "CHAR" Or Column_DataType = "TEXT" Then
                                            If Null_sts = True Then
                                                STR = "''"
                                            Else
                                                STR = "'" & STR & "' "
                                            End If

                                        ElseIf Column_DataType = "INT" Or Column_DataType = "TINYINT" Or Column_DataType = "SMALLINT" Or Column_DataType = "BIGINT" Or Column_DataType = "NUMERIC" Or Column_DataType = "DECIMAL" Then
                                            If Null_sts = True Then
                                                STR = 0
                                            Else
                                                STR = "" & STR & ""
                                            End If

                                        ElseIf Column_DataType = "IMAGE" Then
                                            If Null_sts = True Then
                                                STR = "''"
                                            Else
                                                STR = "'" & STR & "'"
                                            End If

                                        ElseIf Column_DataType = "SMALLDATETIME" Then
                                            If Null_sts = True Then
                                                STR = "'2000-01-01 00:00:00.000'"
                                            Else

                                                DateStringArray(clm) = "@DateValue" & Convert.ToString(clm)
                                                DateArray(clm) = Convert.ToDateTime(STR)
                                                STR = "@DateValue" & Convert.ToString(clm)
                                            End If


                                        ElseIf Column_DataType = "DATETIME" Then
                                            If Null_sts = True Then
                                                STR = "'2000-01-01 00:00:00.000'"
                                            Else

                                                DateStringArray(clm) = "@DateValue" & Convert.ToString(clm)
                                                DateArray(clm) = Convert.ToDateTime(STR)
                                                STR = "@DateValue" & Convert.ToString(clm)
                                            End If

                                        ElseIf Column_DataType = "TIMESTAMP" Then
                                            STR = ""
                                            Continue For

                                        End If




                                        If clm = Val(dt1.Columns.Count - 1) Then
                                            query = query & " " & STR & " )"
                                        Else
                                            query = query & " " & STR & "  , "
                                        End If
                                    Next clm
                                End If
                            Next j


                            cmd_mysql.Connection = con_my


                            cmd_mysql.CommandType = CommandType.Text


                            If Get_Table_Identity_Status(Trim(dt.Rows(I)("TABLE_NAME").ToString)) = True Then

                                cmd_mysql.CommandText = "SET IDENTITY_INSERT " & Trim(dt.Rows(I)("TABLE_NAME").ToString) & " ON"
                                cmd_mysql.ExecuteNonQuery()
                            End If

                            cmd_mysql.Parameters.Clear()
                            For clm = 0 To dt1.Columns.Count - 1
                                If DateStringArray(clm) <> "" Then
                                    cmd_mysql.Parameters.AddWithValue(DateStringArray(clm), DateArray(clm))
                                End If
                            Next
                            If Date1_sts = True Then

                                If IsDate(Date1) Then
                                    cmd_mysql.Parameters.AddWithValue("@Date1", Date1)
                                End If
                            End If


                            Nr = 0
                            cmd_mysql.CommandText = query
                            Nr = cmd_mysql.ExecuteNonQuery()

                            If Get_Table_Identity_Status(Trim(dt.Rows(I)("TABLE_NAME").ToString)) = True Then

                                cmd_mysql.CommandText = "SET IDENTITY_INSERT " & Trim(dt.Rows(I)("TABLE_NAME").ToString) & " OFF"
                                cmd_mysql.ExecuteNonQuery()
                            End If

                            If Nr > 0 Then
                                cmd.Connection = cn1
                                cmd.CommandText = "delete from Trigger_Head where Command = '" & Trim(dt.Rows(I)("Command").ToString) & "' AND AUTO_SLNO = " & Str(Val(dt.Rows(I)("AUTO_SLNO").ToString))
                                cmd.ExecuteNonQuery()

                            End If




                            'DELETE TRIGGER HEAD - USED DATA

                        Else
                            Continue For
                        End If


                    ElseIf UCase(Trim(dt.Rows(I)("Command").ToString)) = "UPDATE" Then


                        Dim DateStringArray(50) As String
                        Dim DateArray(50) As DateTime

                        query = "UPDATE  "

                        Dim qry As String = "select * from " & Trim(dt.Rows(I)("Table_Name").ToString) & IIf(Trim(Primary_Key1) <> "", " WHERE " & Trim(Primary_Key1), " ") & IIf(Trim(Primary_Key2) <> "", " AND " & Trim(Primary_Key2), " ") & IIf(Trim(Primary_Key3) <> "", " AND " & Trim(Primary_Key3), " ")

                        cmd.Connection = cn1

                        cmd.CommandText = qry
                        da1 = New SqlClient.SqlDataAdapter()
                        da1.SelectCommand = cmd

                        If Date1_sts = True Then
                            If IsDate(Date1) Then
                                cmd.Parameters.Clear()
                                cmd.Parameters.AddWithValue("@Date1", Date1)
                            End If
                        End If



                        'da1 = New SqlClient.SqlDataAdapter("select * from " & Trim(dt.Rows(I)("Table_Name").ToString) & IIf(Trim(Primary_Key1) <> "", " WHERE " & Trim(Primary_Key1), " ") & IIf(Trim(Primary_Key2) <> "", " AND " & Trim(Primary_Key2), " ") & IIf(Trim(Primary_Key3) <> "", " AND " & Trim(Primary_Key3), " "), cn1)
                        dt1 = New DataTable
                        da1.Fill(dt1)

                        If dt1.Rows.Count > 0 Then
                            'ROWS
                            For j = 0 To dt1.Rows.Count - 1

                                'COLUMN
                                If dt1.Columns.Count > 0 Then

                                    Dim Table_Identity_sts As Boolean = False
                                    If Get_Table_Identity_Status(Trim(dt.Rows(I)("TABLE_NAME").ToString)) = True Then
                                        Table_Identity_sts = True
                                    End If

                                    query = query & "" & " " & Trim(dt.Rows(I)("TABLE_NAME").ToString) & " SET  "    '  Trim(dt.Rows(I)("Table_Name").ToString) & " SET  "

                                    'LOOP FIELDS AND VALUES
                                    For clm = 0 To dt1.Columns.Count - 1


                                        'skip identity column for updates
                                        If Table_Identity_sts = True Then
                                            If Get_Column_Identity_Status(Trim(dt.Rows(I)("TABLE_NAME").ToString), Trim(dt1.Columns(clm).ColumnName)) = True Then
                                                Continue For
                                            End If
                                        End If

                                        Column_DataType = UCase(Get_Column_dataType(Trim(dt.Rows(I)("Table_Name").ToString), Trim(dt1.Columns(clm).ColumnName)))

                                        Dim STR As String = Trim(dt1.Rows(j)(clm).ToString)

                                        If STR = Nothing Then
                                            Null_sts = True
                                        Else
                                            Null_sts = False
                                        End If

                                        If Column_DataType = "VARCHAR" Or Column_DataType = "NVARCHAR" Or Column_DataType = "CHAR" Or Column_DataType = "TEXT" Then
                                            If Null_sts = True Then
                                                STR = "''"
                                            Else
                                                STR = "'" & STR & "' "
                                            End If

                                        ElseIf Column_DataType = "INT" Or Column_DataType = "TINYINT" Or Column_DataType = "SMALLINT" Or Column_DataType = "BIGINT" Or Column_DataType = "NUMERIC" Or Column_DataType = "DECIMAL" Then
                                            If Null_sts = True Then
                                                STR = 0
                                            Else
                                                STR = "" & STR & ""
                                            End If

                                        ElseIf Column_DataType = "IMAGE" Then
                                            If Null_sts = True Then
                                                STR = "''"
                                            Else
                                                STR = "'" & STR & "'"
                                            End If

                                        ElseIf Column_DataType = "SMALLDATETIME" Then
                                            If Null_sts = True Then
                                                STR = "'2000-01-01 00:00:00.000'"
                                            Else

                                                DateStringArray(clm) = "@DateValue" & Convert.ToString(clm)
                                                DateArray(clm) = Convert.ToDateTime(STR)
                                                STR = "@DateValue" & Convert.ToString(clm)
                                            End If


                                        ElseIf Column_DataType = "DATETIME" Then
                                            If Null_sts = True Then
                                                STR = "'2000-01-01 00:00:00.000'"
                                            Else

                                                DateStringArray(clm) = "@DateValue" & Convert.ToString(clm)
                                                DateArray(clm) = Convert.ToDateTime(STR)
                                                STR = "@DateValue" & Convert.ToString(clm)
                                            End If


                                        ElseIf Column_DataType = "TIMESTAMP" Then


                                            If clm = Val(dt1.Columns.Count - 1) Then
                                                query = Microsoft.VisualBasic.Left(query, Microsoft.VisualBasic.Len(query) - 1)
                                            End If

                                            STR = ""
                                            Continue For


                                            'If Null_sts = True Then
                                            '    STR = "'2000-01-01 00:00:00.000'"
                                            'Else
                                            '    STR = "'" & STR & "'"
                                            'End If
                                        End If

                                        If clm = Val(dt1.Columns.Count - 1) Then
                                            query = query & Trim(dt1.Columns(clm).ColumnName) & " = " & STR & "  "
                                        Else
                                            query = query & Trim(dt1.Columns(clm).ColumnName) & " = " & STR & " , "
                                        End If

                                    Next clm

                                    If Primary_Key1 <> "" Then query = query & " WHERE  " & Primary_Key1
                                    If Primary_Key2 <> "" Then query = query & " AND  " & Primary_Key2
                                    If Primary_Key3 <> "" Then query = query & " AND  " & Primary_Key3

                                End If
                            Next j


                            cmd_mysql.Connection = con_my

                            cmd_mysql.CommandType = CommandType.Text


                            cmd_mysql.Parameters.Clear()
                            For clm = 0 To dt1.Columns.Count - 1
                                If DateStringArray(clm) <> "" Then
                                    cmd_mysql.Parameters.AddWithValue(DateStringArray(clm), DateArray(clm))
                                End If
                            Next
                            If Date1_sts = True Then

                                If IsDate(Date1) Then
                                    cmd_mysql.Parameters.AddWithValue("@Date1", Date1)
                                End If
                            End If



                            Nr = 0
                            cmd_mysql.CommandText = query
                            Nr = cmd_mysql.ExecuteNonQuery()


                            'DELETE TRIGGER HEAD - USED DATA
                            If Nr > 0 Then
                                cmd.Connection = cn1
                                cmd.CommandText = "delete from Trigger_Head where Command = '" & Trim(dt.Rows(I)("Command").ToString) & "' AND AUTO_SLNO = " & Str(Val(dt.Rows(I)("AUTO_SLNO").ToString))
                                cmd.ExecuteNonQuery()

                            End If

                        End If



                    End If

                    ProgressBar4.Minimum = 0
                    ProgressBar4.Maximum = dt.Rows.Count - 1

                    ProgressBar4.Value = I


                    lbl_message_4.Text = I & "  => " & dt.Rows(I).Item("TABLE_NAME").ToString & " Updated..."
                    lbl_message_4.Refresh()



                Next

                trans_mysql.Commit()
                cmd_mysql.Dispose()

                lbl_message_4.Text = "Data updated Successfully.."
                lbl_message_4.Refresh()

            End If
            dt.Clear()

            dt.Dispose()
            da.Dispose()






        Catch ex1 As Exception

            Try
                lbl_message_4.Text = ex1.Message
                lbl_message_4.Refresh()

                trans_mysql.Rollback()

            Catch ex As Exception
                If InStr(LCase(ex1.Message), "network") > 0 Then
                    Timer1.Enabled = False

                    LBL_TITLE.BackColor = Color.Red

                    '   MessageBox.Show("NO INTERNET CONNECTION FOUND", "NO NETWORK FOUND", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    Me.Close()

                Else
                    MessageBox.Show(ex1.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            End Try

        End Try
    End Sub
    Private Sub GET_TABLES_FROM_DATABASE()
        Dim CON As SqlConnection
        Dim da_TABLE As New SqlClient.SqlDataAdapter
        Dim dt_TABLE As New DataTable

        Dim da_COLUMN As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN As New DataTable

        Dim da_KEY As New SqlClient.SqlDataAdapter
        Dim dt_KEY As New DataTable

        Dim da_CONS As New SqlClient.SqlDataAdapter
        Dim dt_CONS As New DataTable

        Dim Query As String = ""
        Dim Sub_Query As String = ""

        Dim con_my As SqlConnection
        Dim cmd_mysql As New SqlCommand
        Dim trans_mysql As SqlTransaction
        Dim Nr As Integer = 0
        Dim LOOP_TABLE_INDX As Integer = 0
        Dim Count As Integer = 0

        ProgressBar1.Value = 0

        Try
            con_my = New SqlConnection(GetConnection_Remote_Sql_String)
            con_my.Open()


            trans_mysql = con_my.BeginTransaction
            cmd_mysql.Transaction = trans_mysql



            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()

            LOOP_TABLE_INDX = 0

LOOP_NEXT_TABLE:
            'GET ALL TABLE IN DATABASE
            da_TABLE = New SqlClient.SqlDataAdapter("select * from information_schema.tables ", CON)
            dt_TABLE = New DataTable
            da_TABLE.Fill(dt_TABLE)

            If dt_TABLE.Rows.Count > 0 Then
                For TABLE_INDX = 0 To dt_TABLE.Rows.Count - 1

                    If LOOP_TABLE_INDX <= dt_TABLE.Rows.Count - 1 Then
                        If TABLE_INDX > LOOP_TABLE_INDX Then
                            LOOP_TABLE_INDX = TABLE_INDX
                        Else
                            TABLE_INDX = LOOP_TABLE_INDX
                        End If
                    Else
                        GoTo END_TASK
                    End If



                    ProgressBar1.Minimum = 0
                    ProgressBar1.Maximum = dt_TABLE.Rows.Count - 1

                    ProgressBar1.Value = TABLE_INDX


                    lbl_message_1.Text = TABLE_INDX & "  => " & dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString & " Created..."
                    lbl_message_1.Refresh()

                    If Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) <> "" Then

                        Query = ""
                        Query = "CREATE TABLE " & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & " ("
                        'GET ALL COLUMN IN TABLE
                        da_COLUMN = New SqlClient.SqlDataAdapter("SELECT *  FROM information_schema.COLUMNS  WHERE TABLE_NAME ='" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "'", CON)
                        dt_COLUMN = New DataTable
                        da_COLUMN.Fill(dt_COLUMN)
                        If dt_COLUMN.Rows.Count > 0 Then
                            For COLUMN_INDX = 0 To dt_COLUMN.Rows.Count - 1
                                If Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) <> "" Then

                                    Sub_Query = ""

                                    If UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "INT" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "TINYINT" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "SMALLINT" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "BIGINT" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString) & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "VARCHAR" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "NVARCHAR" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "CHAR" Then

                                        Dim MAX_CHAR As Integer = Val(dt_COLUMN.Rows(COLUMN_INDX).Item("CHARACTER_MAXIMUM_LENGTH").ToString)
                                        If MAX_CHAR > 500 Then
                                            MAX_CHAR = 500
                                        ElseIf MAX_CHAR < 100 Then
                                            MAX_CHAR = 100
                                        End If


                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString) & "(" & MAX_CHAR & ")" & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "TEXT" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString) & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "IMAGE" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " IMAGE " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")


                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "NUMERIC" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " NUMERIC" & "(" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("NUMERIC_PRECISION").ToString) & "," & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("NUMERIC_SCALE").ToString) & ")" & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "DECIMAL" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " DECIMAL" & "(" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("NUMERIC_PRECISION").ToString) & "," & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("NUMERIC_SCALE").ToString) & ")" & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "SMALLDATETIME" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " SMALLDATETIME "

                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "DATETIME" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " DATETIME "

                                    ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "TIMESTAMP" Then

                                        Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " TIMESTAMP "

                                    Else
                                        Sub_Query = ""
                                    End If






                                    If COLUMN_INDX = Val(dt_COLUMN.Rows.Count - 1) Then
                                        Query = Query & Sub_Query & "  "
                                    Else
                                        Query = Query & Sub_Query & " , "
                                    End If


                                End If

                            Next COLUMN_INDX


                            da_KEY = New SqlClient.SqlDataAdapter("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + QUOTENAME(CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND TABLE_NAME = '" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "' AND TABLE_SCHEMA = 'DBO'", CON)
                            dt_KEY = New DataTable
                            da_KEY.Fill(dt_KEY)
                            If dt_KEY.Rows.Count > 0 Then
                                For KEY_INDX = 0 To dt_KEY.Rows.Count - 1
                                    If Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) <> "" Then
                                        'CHECKING PRIMARY KEY COLUMN
                                        If Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) <> "" And Sub_Query <> "" Then
                                            Sub_Query = Sub_Query & ", PRIMARY KEY (" & Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) & " )"
                                        End If
                                    End If

                                Next KEY_INDX
                            End If
                            dt_KEY.Clear()
                            dt_KEY.Dispose()
                            da_KEY.Dispose()

                            'CHECKING UNIQUE KEY
                            da_KEY = New SqlClient.SqlDataAdapter("select CCU.CONSTRAINT_NAME, CCU.COLUMN_NAME from INFORMATION_SCHEMA.TABLE_CONSTRAINTS as TC inner join INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE as CCU     on TC.CONSTRAINT_CATALOG = CCU.CONSTRAINT_CATALOG    and TC.CONSTRAINT_SCHEMA = CCU.CONSTRAINT_SCHEMA    and TC.CONSTRAINT_NAME = CCU.CONSTRAINT_NAME where  TC.TABLE_NAME = '" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "' and TC.CONSTRAINT_TYPE = 'UNIQUE'", CON)
                            dt_KEY = New DataTable
                            da_KEY.Fill(dt_KEY)
                            If dt_KEY.Rows.Count > 0 Then
                                For KEY_INDX = 0 To dt_KEY.Rows.Count - 1
                                    If Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) <> "" Then
                                        'CHECKING PRIMARY KEY COLUMN
                                        If Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) <> "" And Sub_Query <> "" Then
                                            Sub_Query = Sub_Query & ", UNIQUE KEY (" & Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) & " )"
                                        End If
                                    End If

                                Next KEY_INDX
                            End If
                            dt_KEY.Clear()
                            dt_KEY.Dispose()
                            da_KEY.Dispose()



                            Query = Query & Sub_Query & "  "


                            Query = Query & ") "

                            '=============================================
                            'execute query to mysql database
                            '=============================================
                            cmd_mysql.Connection = con_my
                            cmd_mysql.CommandType = CommandType.Text

                            Nr = 0
                            cmd_mysql.CommandText = Query
                            Nr = cmd_mysql.ExecuteNonQuery()
                            '=============================================

                            '  Count = Count + 1

                            ' lbl_print_message.Text = lbl_print_message.Text & vbCrLf & Count & " ---> " & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & " Created "


                        End If
                        dt_COLUMN.Clear()
                        dt_COLUMN.Dispose()
                        da_COLUMN.Dispose()
                    End If



                Next TABLE_INDX

            End If
            dt_TABLE.Clear()
            dt_TABLE.Dispose()
            da_TABLE.Dispose()


            Common_Procedures.CREATE_DATABASE_STATUS = False

            lbl_message_1.Text = "Created Successfully"
            lbl_message_1.Refresh()

            '  MessageBox.Show("Created Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
END_TASK:




        Catch ex As Exception



            If InStr(UCase(Trim(ex.Message)), "ALREADY") > 0 Then
                LOOP_TABLE_INDX = LOOP_TABLE_INDX + 1
                GoTo LOOP_NEXT_TABLE
            End If

            lbl_message_1.Text = ex.Message
            lbl_message_1.Refresh()

            'MessageBox.Show(ex, "Error", "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try



    End Sub

    Private Sub GET_EXTRA_FIELDS_FROM_DATABASE()
        Dim CON As SqlConnection
        Dim da_TABLE As New SqlClient.SqlDataAdapter
        Dim dt_TABLE As New DataTable

        Dim da_TABLE_MY As New SqlClient.SqlDataAdapter
        Dim dt_TABLE_MY As New DataTable

        Dim da_COLUMN As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN As New DataTable

        Dim da_COLUMN_MY As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN_MY As New DataTable

        Dim da_KEY As New SqlClient.SqlDataAdapter
        Dim dt_KEY As New DataTable

        Dim Query As String = ""
        Dim Sub_Query As String = ""

        Dim con_my As SqlConnection
        Dim cmd_mysql As New SqlCommand
        Dim trans_mysql As SqlTransaction
        Dim Nr As Integer = 0
        Dim LOOP_TABLE_INDX As Integer = 0
        Dim Count As Integer = 0

        ProgressBar2.Value = 0

        Try
            con_my = New SqlConnection(GetConnection_Remote_Sql_String)
            con_my.Open()


            trans_mysql = con_my.BeginTransaction
            cmd_mysql.Transaction = trans_mysql


            cmd_mysql.Connection = con_my
            cmd_mysql.CommandType = CommandType.Text

            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()

            LOOP_TABLE_INDX = 0

LOOP_NEXT_TABLE:
            'GET ALL TABLE IN DATABASE
            da_TABLE = New SqlClient.SqlDataAdapter("select * from information_schema.tables ", CON)
            dt_TABLE = New DataTable
            da_TABLE.Fill(dt_TABLE)

            If dt_TABLE.Rows.Count > 0 Then
                For TABLE_INDX = 0 To dt_TABLE.Rows.Count - 1

                    If LOOP_TABLE_INDX <= dt_TABLE.Rows.Count - 1 Then
                        If TABLE_INDX > LOOP_TABLE_INDX Then
                            LOOP_TABLE_INDX = TABLE_INDX
                        Else
                            TABLE_INDX = LOOP_TABLE_INDX
                        End If
                    Else
                        GoTo END_TASK
                    End If




                    If Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) <> "" Then


                        'GET ALL COLUMN IN TABLE

                        da_COLUMN = New SqlClient.SqlDataAdapter("SELECT *  FROM information_schema.COLUMNS  WHERE TABLE_NAME ='" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "'", CON)
                        dt_COLUMN = New DataTable
                        da_COLUMN.Fill(dt_COLUMN)
                        If dt_COLUMN.Rows.Count > 0 Then
                            For COLUMN_INDX = 0 To dt_COLUMN.Rows.Count - 1
                                If Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) <> "" Then


                                    da_TABLE_MY = New SqlDataAdapter("SELECT * FROM information_schema.COLUMNS WHERE COLUMN_NAME = '" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & "' AND TABLE_NAME ='" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "' and  TABLE_SCHEMA = '" & Trim(Common_Procedures.DATABASE_NAME) & "' ", con_my)
                                    dt_TABLE_MY = New DataTable
                                    da_TABLE_MY.Fill(dt_TABLE_MY)
                                    If dt_TABLE_MY.Rows.Count = 0 Then


                                        Sub_Query = "ALTER TABLE " & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & " ADD COLUMN  "


                                        If UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "INT" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "TINYINT" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "SMALLINT" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "BIGINT" Then

                                            Sub_Query = Sub_Query & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & " " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString) & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                        ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "VARCHAR" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "NVARCHAR" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "CHAR" Then

                                            Dim MAX_CHAR As Integer = Val(dt_COLUMN.Rows(COLUMN_INDX).Item("CHARACTER_MAXIMUM_LENGTH").ToString)
                                            If MAX_CHAR > 500 Then
                                                MAX_CHAR = 500
                                            ElseIf MAX_CHAR < 100 Then
                                                MAX_CHAR = 100
                                            End If

                                            Sub_Query = Sub_Query & " `" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & "` " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString) & "(" & MAX_CHAR & ")" & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                        ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "TEXT" Then

                                            Sub_Query = Sub_Query & " `" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & "` " & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString) & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                        ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "IMAGE" Then

                                            Sub_Query = Sub_Query & " `" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & "` TEXT " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                        ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "NUMERIC" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "DECIMAL" Then

                                            Sub_Query = Sub_Query & " `" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & "` DECIMAL" & "(" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("NUMERIC_PRECISION").ToString) & "," & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("NUMERIC_SCALE").ToString) & ")" & " " & IIf(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("IS_NULLABLE").ToString) = "YES", " NULL ", " NOT NULL ")

                                        ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "SMALLDATETIME" Or UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "DATETIME" Then

                                            Sub_Query = Sub_Query & " `" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & "` DATE "

                                        ElseIf UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)) = "TIMESTAMP" Then

                                            Sub_Query = Sub_Query & " `" & Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMN_NAME").ToString) & "` TIMESTAMP "
                                        Else
                                            Sub_Query = ""
                                        End If



                                        If Sub_Query <> "" Then
                                            Nr = 0
                                            cmd_mysql.CommandText = Sub_Query
                                            Nr = cmd_mysql.ExecuteNonQuery()
                                        End If



                                    End If

                                    dt_COLUMN_MY.Clear()
                                    dt_COLUMN_MY.Dispose()
                                    da_TABLE_MY.Dispose()

                                End If

                            Next COLUMN_INDX


                            ''CHECKING PRIMARY KEY
                            'da_KEY = New SqlClient.SqlDataAdapter("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + QUOTENAME(CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND TABLE_NAME = '" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "' AND TABLE_SCHEMA = 'DBO'", CON)
                            'dt_KEY = New DataTable
                            'da_KEY.Fill(dt_KEY)
                            'If dt_KEY.Rows.Count > 0 Then
                            '    For KEY_INDX = 0 To dt_KEY.Rows.Count - 1
                            '        If Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) <> "" Then
                            '            'CHECKING PRIMARY KEY COLUMN
                            '            If Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) <> "" And Sub_Query <> "" Then
                            '                Sub_Query = Sub_Query & ", PRIMARY KEY (`" & Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) & "` )"
                            '            End If
                            '        End If

                            '    Next KEY_INDX
                            'End If
                            'dt_KEY.Clear()
                            'dt_KEY.Dispose()
                            'da_KEY.Dispose()



                        End If
                        dt_COLUMN.Clear()
                        dt_COLUMN.Dispose()
                        da_COLUMN.Dispose()
                    End If

                    ProgressBar2.Minimum = 0
                    ProgressBar2.Maximum = dt_TABLE.Rows.Count - 1

                    ProgressBar2.Value = TABLE_INDX


                    lbl_message_2.Text = TABLE_INDX & "  => " & dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString & " Created..."
                    lbl_message_2.Refresh()


                Next TABLE_INDX

            End If
            dt_TABLE.Clear()
            dt_TABLE.Dispose()
            da_TABLE.Dispose()



            Common_Procedures.CREATE_DATABASE_STATUS = False

            lbl_message_2.Text = "Field Created Successfully"
            lbl_message_2.Refresh()
            'MessageBox.Show("Field Created Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
END_TASK:




        Catch ex As Exception
            lbl_message_2.Text = ex.Message

            lbl_message_2.Refresh()

            If InStr(UCase(Trim(ex.Message)), "ALREADY") > 0 Then
                LOOP_TABLE_INDX = LOOP_TABLE_INDX + 1
                GoTo LOOP_NEXT_TABLE
            End If

            MessageBox.Show(ex, "Error", "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try



    End Sub
    Private Sub GET_OLD_DATA_FROM_DATABASE(ByVal table_name As String)
        Dim CON As SqlClient.SqlConnection
        Dim con_my As SqlConnection
        Dim cmd_mysql As New SqlCommand
        Dim trans_mysql As SqlTransaction
        Dim cmd As New SqlClient.SqlCommand
        Dim IdNo As Integer, Nr As Integer
        'Dim DBPartName As String
        Dim table_name_for_msg As String = ""
        Dim da As New SqlClient.SqlDataAdapter
        Dim dt As New DataTable

        Dim da1 As New SqlClient.SqlDataAdapter
        Dim dt1 As New DataTable
        Dim query As String = ""


        ProgressBar3.Value = 0
        Try


            con_my = New SqlConnection(GetConnection_Remote_Sql_String)
            con_my.Open()



            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()


            trans_mysql = con_my.BeginTransaction
            cmd_mysql.Transaction = trans_mysql


            cmd_mysql.Connection = con_my

            cmd_mysql.CommandType = CommandType.Text

            Dim tab_name_cont As String = ""
            tab_name_cont = "Table_Name ='" & table_name & "' "

            da = New SqlClient.SqlDataAdapter("select * from information_schema.tables  WHERE " & IIf(Trim(tab_name_cont) <> "", tab_name_cont & " and ", "") & " TABLE_CATALOG ='" & Common_Procedures.DATABASE_NAME & "'", CON)
            dt = New DataTable
            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                For I = 0 To dt.Rows.Count - 1


                    table_name_for_msg = Trim(dt.Rows(I)("Table_Name").ToString)

                    query = ""
                    ' COLUMNS FROM TABLE
                    da1 = New SqlClient.SqlDataAdapter("select * from information_schema.COLUMNS WHERE TABLE_NAME= '" & Trim(dt.Rows(I)("Table_Name").ToString) & "' AND TABLE_NAME NOT LIKE '%TEMP%'  ", CON)
                    dt1 = New DataTable
                    da1.Fill(dt1)

                    If dt1.Rows.Count > 0 Then

                        query = "INSERT INTO " & Trim(dt.Rows(I)("Table_Name").ToString) & " ("


                        For RW = 0 To dt1.Rows.Count - 1

                            'skip some column for error coming
                            If UCase(Trim(dt1.Rows(RW).Item("COLUMN_NAME").ToString)) = "UPDATETIME" Then
                                Continue For
                            End If
                            '=============

                            If RW = Val(dt1.Rows.Count - 1) Then
                                query = query & " " & Trim(dt1.Rows(RW).Item("COLUMN_NAME").ToString) & " ) "
                            Else
                                query = query & " " & Trim(dt1.Rows(RW).Item("COLUMN_NAME").ToString) & " , "
                            End If

                        Next RW
                    End If
                    dt1.Clear()
                    dt1.Dispose()
                    da1.Dispose()



                    If query = "" Then
                        Continue For
                    End If

                    Dim SUB_QUERY As String = ""

                    ' VALUES FROM TABLE
                    da1 = New SqlClient.SqlDataAdapter("select * from " & Trim(dt.Rows(I)("Table_Name").ToString) & "  ", CON)
                    dt1 = New DataTable
                    da1.Fill(dt1)
                    If dt1.Rows.Count > 0 Then

                        Nr = 0
                        cmd_mysql.CommandText = "DELETE FROM " & Trim(dt.Rows(I)("Table_Name").ToString)
                        Nr = cmd_mysql.ExecuteNonQuery()




                        For j = 0 To dt1.Rows.Count - 1


                            SUB_QUERY = query & " VALUES ( "

                            If dt1.Columns.Count > 0 Then

                                'LOOP FIELDS VALUES
                                For clm = 0 To dt1.Columns.Count - 1

                                    Dim STR As String = Trim(dt1.Rows(j)(clm).ToString)


                                    If STR Is Nothing Or STR = "" Then
                                        If InStr(LCase(Trim(Common_Procedures.Get_DataType_From_Colum_Name(CON, Trim(dt.Rows(I)("Table_Name").ToString), clm + 1))), "datetime") > 0 Then
                                            STR = Format(Convert.ToDateTime("1/1/2000"), "yyyy-MM-dd")
                                        Else
                                            STR = " "
                                        End If

                                    End If


                                    If (InStr(Trim(STR), "-") > 0 Or InStr(Trim(STR), "/") > 0) And (InStr(Trim(STR), ":") > 0 Or InStr(Trim(STR), "00:00:00") > 0) And Microsoft.VisualBasic.Len(Trim(STR)) >= 10 Then

                                        STR = Format(Convert.ToDateTime(STR), "yyyy-MM-dd")

                                    ElseIf InStr(Trim(STR), "0x") > 0 Or InStr(Trim(STR), "System.Byte[]") > 0 Then
                                        'skip some column values for error coming
                                        Continue For
                                    End If

                                    '
                                    STR = Replace(STR, "'", "`")

                                    If STR Is Nothing Then
                                        STR = vbNull
                                    End If

                                    If clm = Val(dt1.Columns.Count - 1) Then
                                        SUB_QUERY = SUB_QUERY & "'" & STR & "' )"
                                    Else
                                        SUB_QUERY = SUB_QUERY & "'" & STR & "' , "
                                    End If
                                Next clm




                                Nr = 0
                                cmd_mysql.CommandText = SUB_QUERY
                                Nr = cmd_mysql.ExecuteNonQuery()



                            End If
                        Next j



                    End If
                    dt1.Clear()
                    dt1.Dispose()
                    da1.Dispose()


                    'progress bar

                    ProgressBar3.Minimum = 0
                    ProgressBar3.Maximum = dt.Rows.Count - 1

                    ProgressBar3.Value = I


                    lbl_message_3.Text = I & "  => " & Trim(dt.Rows(I)("Table_Name").ToString) & " Updated..."
                    lbl_message_3.Refresh()

                Next I


                trans_mysql.Commit()
                cmd_mysql.Dispose()

                DATE_UPDATE_STATUS = False

                Timer2.Enabled = False

                lbl_message_3.Text = "Updated Successfully.."
                lbl_message_3.Refresh()

                'MessageBox.Show("DATA UPDATED SUCCESSFULLY", "UPDATED", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If


        Catch ex1 As Exception
            Try
                lbl_message_3.Text = "(" & table_name_for_msg & " )" & ex1.Message
                lbl_message_3.Refresh()


                trans_mysql.Rollback()

            Catch ex As Exception


                If InStr(LCase(ex1.Message), "unable to connect to any of the specified mysql hosts.") > 0 Then
                    Timer1.Enabled = False

                    lbl_message_3.Text = "NO INTERNET CONNECTION FOUND"
                    lbl_message_3.Refresh()

                    'MessageBox.Show("NO INTERNET CONNECTION FOUND", "NO NETWORK FOUND", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    Me.Close()

                Else
                    lbl_message_3.Text = ex1.Message
                    lbl_message_3.Refresh()
                    'MessageBox.Show(ex1.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            End Try








        End Try
    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim file As String = "whatever file you want"
        If System.IO.File.Exists(Trim(Common_Procedures.AppPath) & "\process.ini") = True Then
            Process.Start(Trim(Common_Procedures.AppPath) & "\process.ini")
        Else
            MsgBox("File Does Not Exist!!", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub GET_OLD_DATA_FROM_OLD_DATABASE()
        Dim CON As SqlConnection
        Dim da_TABLE As New SqlClient.SqlDataAdapter
        Dim dt_TABLE As New DataTable

        Dim da_COLUMN As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN As New DataTable

        Dim da_KEY As New SqlClient.SqlDataAdapter
        Dim dt_KEY As New DataTable

        Dim da_CONS As New SqlClient.SqlDataAdapter
        Dim dt_CONS As New DataTable

        Dim Query As String = ""
        Dim Sub_Query As String = ""

        Dim con_my As SqlConnection
        Dim cmd_mysql As New SqlCommand
        Dim trans_mysql As SqlTransaction
        Dim Nr As Integer = 0
        Dim LOOP_TABLE_INDX As Integer = 0
        Dim Count As Integer = 0

        ProgressBar1.Value = 0

        Try
            con_my = New SqlConnection(GetConnection_Remote_Sql_String)
            con_my.Open()


            trans_mysql = con_my.BeginTransaction
            cmd_mysql.Transaction = trans_mysql



            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()

            LOOP_TABLE_INDX = 0

LOOP_NEXT_TABLE:
            'GET ALL TABLE IN DATABASE
            da_TABLE = New SqlClient.SqlDataAdapter("select * from information_schema.tables ", CON)
            dt_TABLE = New DataTable
            da_TABLE.Fill(dt_TABLE)

            If dt_TABLE.Rows.Count > 0 Then
                For TABLE_INDX = 0 To dt_TABLE.Rows.Count - 1

                    If Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString()) <> "" Then
                        GET_OLD_DATA_FROM_DATABASE(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString())

                        lbl_message_1.Text = Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString())
                    End If

                Next TABLE_INDX

            End If
            dt_TABLE.Clear()
            dt_TABLE.Dispose()
            da_TABLE.Dispose()


            Common_Procedures.CREATE_DATABASE_STATUS = False

            lbl_message_1.Text = "Created Successfully"
            lbl_message_1.Refresh()


END_TASK:




        Catch ex As Exception



            If InStr(UCase(Trim(ex.Message)), "ALREADY") > 0 Then
                LOOP_TABLE_INDX = LOOP_TABLE_INDX + 1
                GoTo LOOP_NEXT_TABLE
            End If

            lbl_message_1.Text = ex.Message
            lbl_message_1.Refresh()


        End Try



    End Sub
    Private Function Get_Column_dataType(ByVal Table_name, ByVal column_name) As String

        Dim CON As SqlConnection
        Dim da_TABLE As New SqlClient.SqlDataAdapter
        Dim dt_TABLE As New DataTable

        Dim da_COLUMN As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN As New DataTable

        Dim Column_DataType As String = ""

        Try

            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()


            If Trim(Table_name) <> "" Then

                da_COLUMN = New SqlClient.SqlDataAdapter("SELECT DATA_TYPE  FROM information_schema.COLUMNS  WHERE TABLE_NAME ='" & Trim(Table_name) & "' and COLUMN_NAME = '" & Trim(column_name) & "'", CON)
                dt_COLUMN = New DataTable
                da_COLUMN.Fill(dt_COLUMN)
                If dt_COLUMN.Rows.Count > 0 Then
                    For COLUMN_INDX = 0 To dt_COLUMN.Rows.Count - 1
                        If Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString) <> "" Then

                            Column_DataType = Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("DATA_TYPE").ToString)

                        End If

                    Next COLUMN_INDX

                End If
                dt_COLUMN.Clear()
                dt_COLUMN.Dispose()
                da_COLUMN.Dispose()
            End If




        Catch ex As Exception

            Return Column_DataType

            MessageBox.Show(ex, "Error", "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        Return Column_DataType

    End Function
    Private Function Get_Table_Identity_Status(ByVal Table_name) As Boolean

        Dim CON As SqlConnection
        Dim da_TABLE As New SqlClient.SqlDataAdapter
        Dim dt_TABLE As New DataTable

        Dim da_COLUMN As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN As New DataTable

        Dim Column_Identity_sts As Boolean = False

        Try

            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()


            If Trim(Table_name) <> "" Then

                da_COLUMN = New SqlClient.SqlDataAdapter("SELECT * FROM  SYS.IDENTITY_COLUMNS  WHERE OBJECT_NAME(OBJECT_ID) = '" & Trim(Table_name) & "' ", CON)
                dt_COLUMN = New DataTable
                da_COLUMN.Fill(dt_COLUMN)

                Column_Identity_sts = False
                If dt_COLUMN.Rows.Count > 0 Then
                    Column_Identity_sts = True
                End If

                dt_COLUMN.Clear()
                dt_COLUMN.Dispose()
                da_COLUMN.Dispose()
            End If




        Catch ex As Exception

            Return Column_Identity_sts

            MessageBox.Show(ex, "Error", "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        Return Column_Identity_sts

    End Function

    Private Function Get_Column_Identity_Status(ByVal Table_name, ByVal column_name) As Boolean

        Dim CON As SqlConnection
        Dim da_TABLE As New SqlClient.SqlDataAdapter
        Dim dt_TABLE As New DataTable

        Dim da_COLUMN As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN As New DataTable

        Dim Column_Identity_sts As Boolean = False

        Try

            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()


            If Trim(Table_name) <> "" Then

                da_COLUMN = New SqlClient.SqlDataAdapter("SELECT   OBJECT_NAME(OBJECT_ID) AS TABLENAME, NAME AS COLUMNNAME, SEED_VALUE,              INCREMENT_VALUE,              LAST_VALUE,            IS_NOT_FOR_REPLICATION FROM     SYS.IDENTITY_COLUMNS     WHERE OBJECT_NAME(OBJECT_ID) = '" & Trim(Table_name) & "'", CON)
                dt_COLUMN = New DataTable
                da_COLUMN.Fill(dt_COLUMN)
                If dt_COLUMN.Rows.Count > 0 Then
                    For COLUMN_INDX = 0 To dt_COLUMN.Rows.Count - 1
                        If Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMNNAME").ToString) <> "" Then

                            Column_Identity_sts = False
                            If UCase(column_name) = UCase(Trim(dt_COLUMN.Rows(COLUMN_INDX).Item("COLUMNNAME").ToString)) Then
                                Column_Identity_sts = True
                            End If
                        End If
                    Next COLUMN_INDX
                End If
                dt_COLUMN.Clear()
                dt_COLUMN.Dispose()
                da_COLUMN.Dispose()
            End If




        Catch ex As Exception

            Return Column_Identity_sts

            MessageBox.Show(ex, "Error", "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        Return Column_Identity_sts

    End Function
    Private Sub GET_TABLES_FROM_DATABASE_FOR_CREATE_TRIGGER()
        Dim CON As SqlConnection
        Dim da_TABLE As New SqlClient.SqlDataAdapter
        Dim dt_TABLE As New DataTable

        Dim da_COLUMN As New SqlClient.SqlDataAdapter
        Dim dt_COLUMN As New DataTable

        Dim da_KEY As New SqlClient.SqlDataAdapter
        Dim dt_KEY As New DataTable

        Dim da_CONS As New SqlClient.SqlDataAdapter
        Dim dt_CONS As New DataTable

        Dim Query As String = ""
        Dim Sub_Query As String = ""
        Dim Column_1 As String = ""
        Dim Column_2 As String = ""
        Dim Column_3 As String = ""

        Dim con_my As SqlConnection
        Dim cmd_mysql As New SqlCommand
        Dim trans_mysql As SqlTransaction
        Dim Nr As Integer = 0
        Dim LOOP_TABLE_INDX As Integer = 0
        Dim Count As Integer = 0

        Dim Skip_Tables As String = ""


        Skip_Tables = "~EntryTemp~" & _
                      "~EntryTemp_Simple~" & _
                      "~EntryTempSub~" & _
                      "~ReportTemp~" & _
                      "~ReportTempSub~" & _
                      "~Trigger_Head~"


        Timer1.Stop()
        Timer2.Stop()

        Try
            ' con_my = New SqlConnection(GetConnection_Remote_Sql_String)
            con_my = New SqlClient.SqlConnection(Common_Procedures.Connection_String)

            con_my.Open()


            trans_mysql = con_my.BeginTransaction
            cmd_mysql.Transaction = trans_mysql



            CON = New SqlClient.SqlConnection(Common_Procedures.Connection_String)
            CON.Open()

            LOOP_TABLE_INDX = 0

LOOP_NEXT_TABLE:
            'GET ALL TABLE IN DATABASE
            da_TABLE = New SqlClient.SqlDataAdapter("select * from information_schema.tables ", CON)
            dt_TABLE = New DataTable
            da_TABLE.Fill(dt_TABLE)

            If dt_TABLE.Rows.Count > 0 Then
                For TABLE_INDX = 0 To dt_TABLE.Rows.Count - 1

                    If LOOP_TABLE_INDX <= dt_TABLE.Rows.Count - 1 Then
                        If TABLE_INDX > LOOP_TABLE_INDX Then
                            LOOP_TABLE_INDX = TABLE_INDX
                        Else
                            TABLE_INDX = LOOP_TABLE_INDX
                        End If
                    Else
                        GoTo END_TASK
                    End If

                    If TABLE_INDX = 202 Then
                        Column_1 = ""
                    End If

                    If Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) <> "" And InStr(Trim(Skip_Tables), "~" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "~") = 0 Then

                        Column_1 = ""
                        Column_2 = ""
                        Column_3 = ""

                        Query = ""
                        Query = " CREATE TRIGGER Trig_" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & " ON dbo." & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "" & _
                                        " FOR INSERT, UPDATE, DELETE " & _
                                        " AS "

                        'GET ALL COLUMN IN TABLE
                        da_COLUMN = New SqlClient.SqlDataAdapter("SELECT *  FROM information_schema.COLUMNS  WHERE TABLE_NAME ='" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "'", CON)
                        dt_COLUMN = New DataTable
                        da_COLUMN.Fill(dt_COLUMN)
                        If dt_COLUMN.Rows.Count > 0 Then


                            'CHECKING PRIMARY KEY COLUMN

                            da_KEY = New SqlClient.SqlDataAdapter("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + QUOTENAME(CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND TABLE_NAME = '" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "' AND TABLE_SCHEMA = 'DBO' AND CONSTRAINT_NAME like 'PK%'", CON)
                            dt_KEY = New DataTable
                            da_KEY.Fill(dt_KEY)
                            If dt_KEY.Rows.Count > 0 Then
                                For KEY_INDX = 0 To dt_KEY.Rows.Count - 1

                                    Sub_Query = ""
                                    'CHECKING PRIMARY KEY COLUMN
                                    If Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) <> "" And Query <> "" Then

                                        Sub_Query = " " & Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) '& " = " & " CAST (D." & Trim(dt_KEY.Rows(KEY_INDX).Item("COLUMN_NAME").ToString) & " as VARCHAR )" & " "

                                    End If
                                    If KEY_INDX = 0 Then
                                        Column_1 = Sub_Query
                                    ElseIf KEY_INDX = 1 Then
                                        Column_2 = Sub_Query
                                    ElseIf KEY_INDX = 3 Then
                                        Column_3 = Sub_Query
                                    Else
                                        Exit For
                                    End If

                                Next KEY_INDX
                            End If
                            dt_KEY.Clear()
                            dt_KEY.Dispose()
                            da_KEY.Dispose()



                            Sub_Query = ""
                            Sub_Query = "  IF EXISTS ( SELECT 0 FROM Deleted )" & _
                                            " BEGIN" & _
                                                " IF EXISTS ( SELECT 0 FROM Inserted ) " & _
                                                    " BEGIN " & _
                                                        " INSERT  INTO dbo.Trigger_Head( Command ,Table_Name ,Primary_key_1 ,Primary_key_2 ,Primary_key_3 ,Modified_Date)" & _
                                                                                " SELECT  'UPDATE' ,'" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "' ,	" & IIf(Trim(Column_1) <> "", "'" & Trim(Column_1) & "  = ~'  + CAST (D." & Trim(Column_1) & " as VARCHAR ) +'~' ,", "'',") & _
                                                                                                                        IIf(Trim(Column_2) <> "", "'" & Trim(Column_2) & "  = ~'  + CAST (D." & Trim(Column_2) & " as VARCHAR ) +'~' ,", "'',") & _
                                                                                                                        IIf(Trim(Column_3) <> "", "'" & Trim(Column_3) & "  = ~'  + CAST (D." & Trim(Column_3) & " as VARCHAR ) +'~' ,", "'',") & "   GETDATE()  FROM    Inserted D " & _
                                                     " END" & _
                                                " ELSE" & _
                                                     " BEGIN" & _
                                                        " INSERT  INTO dbo.Trigger_Head ( Command ,Table_Name ,Primary_key_1 ,Primary_key_2 ,Primary_key_3 ,Modified_Date) " & _
                                                                                "SELECT  'DELETE' ,'" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "'  , " & IIf(Trim(Column_1) <> "", "'" & Trim(Column_1) & "  = ~'  + CAST (D." & Trim(Column_1) & " as VARCHAR ) +'~' ,", "'',") & _
                                                                                                                    IIf(Trim(Column_2) <> "", "'" & Trim(Column_2) & "  = ~'  + CAST (D." & Trim(Column_2) & " as VARCHAR ) +'~' ,", "'',") & _
                                                                                                                    IIf(Trim(Column_3) <> "", "'" & Trim(Column_3) & "  = ~'  + CAST (D." & Trim(Column_3) & " as VARCHAR ) +'~' ,", "'',") & "  GETDATE()  FROM    Deleted D" & _
                                                    " End" & _
                                            " END" & _
                                        " ELSE" & _
                                            " BEGIN" & _
                                                " INSERT  INTO dbo.Trigger_Head ( Command ,Table_Name ,Primary_key_1 ,Primary_key_2 ,Primary_key_3 ,Modified_Date   )" & _
                                                                        "SELECT  'INSERT' ,'" & Trim(dt_TABLE.Rows(TABLE_INDX).Item("TABLE_NAME").ToString) & "' , " & IIf(Trim(Column_1) <> "", "'" & Trim(Column_1) & "  = ~'  + CAST (D." & Trim(Column_1) & " as VARCHAR ) +'~' ,", "'',") & _
                                                                                                                        IIf(Trim(Column_2) <> "", "'" & Trim(Column_2) & "  = ~'  + CAST (D." & Trim(Column_2) & " as VARCHAR ) +'~' ,", "'',") & _
                                                                                                                        IIf(Trim(Column_3) <> "", "'" & Trim(Column_3) & "  = ~'  + CAST (D." & Trim(Column_3) & " as VARCHAR ) +'~' ,", "'',") & "  GETDATE()  FROM    Inserted D " & _
                                            " END" & _
                                            " "

                            Query = Query & "  " & Sub_Query


                            '=============================================
                            'execute query to mssql database
                            '=============================================
                            cmd_mysql.Connection = con_my
                            cmd_mysql.CommandType = CommandType.Text

                            Nr = 0
                            cmd_mysql.CommandText = Query
                            Nr = cmd_mysql.ExecuteNonQuery()
                            '=============================================


                        End If
                        dt_COLUMN.Clear()
                        dt_COLUMN.Dispose()
                        da_COLUMN.Dispose()
                    End If



                Next TABLE_INDX

                trans_mysql.Commit()

                cmd_mysql.Dispose()

            End If
            dt_TABLE.Clear()
            dt_TABLE.Dispose()
            da_TABLE.Dispose()





            MessageBox.Show("Created Successfully", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
END_TASK:

            Timer1.Stop()
            Timer2.Stop()

            Application.Exit()


        Catch ex As Exception



            If InStr(UCase(Trim(ex.Message)), "ALREADY") > 0 Then
                LOOP_TABLE_INDX = LOOP_TABLE_INDX + 1
                GoTo LOOP_NEXT_TABLE
            Else
                trans_mysql.Rollback()
                Application.Exit()
            End If

            MessageBox.Show(ex, "Error", "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try



    End Sub

End Class
