Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Management
Imports System.IO.Ports
Imports System.Security.Cryptography.X509Certificates
Imports System.Net
Imports System.Net.Security
Imports System.Net.Mail

Public Class Common_Procedures
   


    Public Shared CompanyDetailsDataBaseName As String = "LiyaSys_Syncer_CompanyGroup_Details"
    Public Shared Connection_String As String
    Public Shared ConnectionString_CompanyGroupdetails As String
    Public Shared ConnectionString_Master As String
    Public Shared ServerName As String
    Public Shared ServerPassword As String
    Public Shared ServerWindowsLogin As String
    Public Shared ServerDataBaseLocation_InExTernalUSB As String

    Public Shared CompGroupIdNo As Integer
    Public Shared CompGroupName As String
    Public Shared CompGroupFnRange As String

    Public Shared DataBaseName As String

    Public Shared First_Opened_Today As Boolean = False
    
    Public Shared FnYearCode As String
    Public Shared CompIdNo As Integer
    Public Shared Company_FromDate As Date
    Public Shared Company_ToDate As Date
    Public Shared AppPath As String
    
    Public Shared DATABASE_NAME As String
    Public Shared CREATE_DATABASE_STATUS As Boolean
    Public Shared FIELD_CHECK_STATUS As Boolean
    Public Shared OLD_DATA_STATUS As Boolean
    Public Shared NEW_DATA_STATUS As Boolean
    Public Shared FORM_VISIBLE_STATUS As Boolean

    Public Shared Function Create_Sql_ConnectionString(ByVal DBName As String) As String
        Dim myConnection_string As String
        Dim pth As String
        Dim fs As FileStream
        Dim w As StreamWriter
        Dim sInpIP As String = ""

        If Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "IP" Or Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "SIP" Or Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "DIP" Then
            If Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "DIP" Then
                If Common_Procedures.First_Opened_Today = True Then
                    sInpIP = InputBox("Enter Server System IP address :", "FOR CORRECT SERVER SYSTEM IP ADDRESS..", Trim(Common_Procedures.ServerName))

                    If Trim(sInpIP) <> "" Then
                        pth = Trim(Common_Procedures.AppPath) & "\connection.ini"

                        If File.Exists(pth) = True Then
                            File.Delete(pth)
                        End If

                        Common_Procedures.ServerName = Trim(sInpIP)

                        fs = New FileStream(pth, FileMode.Create)
                        w = New StreamWriter(fs)
                        w.WriteLine(Trim(Common_Procedures.ServerName) & "," & Trim(Common_Procedures.ServerPassword) & ",DIP")
                        w.Close()
                        fs.Close()
                        w.Dispose()
                        fs.Dispose()

                        Common_Procedures.First_Opened_Today = False

                    End If
                End If
            End If

            myConnection_string = "Data Source=" & Trim(Common_Procedures.ServerName) & ",1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False;Connect Timeout=60"
            
        ElseIf Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "WIN" Then
            myConnection_string = "Data Source=" & Trim(Common_Procedures.ServerName) & ";Initial Catalog=" & Trim(DBName) & ";Integrated Security=True;Connect Timeout=60"

        Else
            myConnection_string = "Data Source=" & Trim(Common_Procedures.ServerName) & ";Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False;Connect Timeout=60"

        End If


      

        Create_Sql_ConnectionString = Trim(myConnection_string)

    End Function
    Public Shared Function Get_DataType_From_Colum_Name(ByVal Cn1 As SqlClient.SqlConnection, ByVal Table_name As String, ByVal ordinal_position As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal vCmpGrpIdNo As Integer = 0) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim Data_type As String = ""
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select data_type from information_schema.COLUMNS WHERE ordinal_position = " & ordinal_position & " and TABLE_NAME= '" & Trim(Table_name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        Data_type = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                Data_type = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()




        Get_DataType_From_Colum_Name = Data_type

    End Function
    Public Shared Function get_Company_DataBaseName(ByVal CompIdNo As Integer) As String
        Dim DbNm As String = ""
        Dim S As String = ""

        DbNm = ""

        If Trim(CompanyDetailsDataBaseName) <> "" Then

            S = Replace(Trim(LCase(CompanyDetailsDataBaseName)), "_companygroup_details", "")

            DbNm = Trim(S) & "_" & Trim(Val(CompIdNo))

        End If

        get_Company_DataBaseName = Trim(DbNm)

    End Function
End Class