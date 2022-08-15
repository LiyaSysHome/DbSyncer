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


    Public Shared Own_Or_Tie As String = "OWN"


    Public Shared CompanyDetailsDataBaseName As String = "PLit_Billing_CompanyGroup_Details"
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

    Public Shared FnRange As String
    Public Shared FnYearCode As String
    Public Shared CompIdNo As Integer
    Public Shared Company_FromDate As Date
    Public Shared Company_ToDate As Date
    Public Shared AppPath As String
    Public Shared MRP_saving As Single

    Public Shared Print_OR_Preview_Status As Integer

    Public Shared BillAdj_Amt As Single = 0

    Public Shared VoucherType As String = ""
    Public Shared Voucher_CR_Name As String = ""
    Public Shared Voucher_CR_or_DR As String = ""
    Public Shared Voucher_Code As String = ""
    Public Shared Voucher_DR_Name As String = ""

    Public Shared Password_Input As String = ""
    Public Shared Sales_Or_Service As String = ""
    Public Shared SalesEntryType As String = ""

    Public Shared First_Opened_Today As Boolean = False
    Public Shared MDI_LedType As String
    Public Shared Last_Closed_FormName As String
    Public Shared GST_and_VAT_Entry_Status As Boolean = False
    Public Shared DriveVolumeSerialName As String = ""
    Public Shared Selected_Company_Idno As Integer
    Public Shared Selected_Company_Name As String
    Public Shared Company_ReSelection_Sts As Boolean = False
    Public Shared MousePositionX As Integer = 0
    Public Shared MousePositionY As Integer = 0


    Public Shared GotFocus_BackColor As Drawing.Color = Color.FromArgb(146, 211, 248)
    Public Shared GotFocus_ForeColor As Drawing.Color = Color.Black

    Public Shared LostFocus_BackColor As Drawing.Color = Color.White
    Public Shared LostFocus_ForeColor As Drawing.Color = Color.Black

    Public Shared OneDrive_Path As String
    Public Shared OneDrive_Status As Boolean
    Public Shared account_flip_status As Boolean
    Public Shared vShowEntrance_Status_ForCC As Boolean = False
    Public Shared vShowEntrance_Status As Boolean = False
    Public Shared SerialPort1 As New SerialPort

    Public Shared DATABASE_NAME As String
    Public Shared CREATE_DATABASE_STATUS As Boolean
    Public Shared FIELD_CHECK_STATUS As Boolean
    Public Shared OLD_DATA_STATUS As Boolean
    Public Shared NEW_DATA_STATUS As Boolean
    Public Shared FORM_VISIBLE_STATUS As Boolean




    Public Structure Report_ComboDetails
        Dim PKey As String
        Dim TableName As String
        Dim Selection_FieldName As String
        Dim Return_FieldName As String
        Dim Condition As String
        Dim Display_Name As String
        Dim BlankFieldCondition As String
        Dim CtrlType_Cbo_OR_Txt As String
    End Structure
    Public Shared RptCboDet(10) As Report_ComboDetails

    Public Structure Encryption_DeEncryption_Pass_Salt_Phrase
        Dim passPhrase As String
        Dim saltValue As String
    End Structure
    Public Shared Entrance_SQL_PassWord As Encryption_DeEncryption_Pass_Salt_Phrase
    Public Shared UserCreation_AcPassWord As Encryption_DeEncryption_Pass_Salt_Phrase
    Public Shared UserCreation_UnAcPassWord As Encryption_DeEncryption_Pass_Salt_Phrase
    Public Shared SoftWareRegister As Encryption_DeEncryption_Pass_Salt_Phrase

    Public Structure SettingsDetails
        Dim CustomerCode As String
        Dim CustomerDBCode As String
        Dim CompanyName As String
        Dim SoftWare_UserType As String
        Dim Sdd As Date
        Dim AutoBackUp_Date As Date

        Dim SMS_Provider_SenderID As String
        Dim SMS_Provider_Key As String
        Dim SMS_Provider_RouteID As String
        Dim SMS_Provider_Type As String

        Dim Email_Address As String
        Dim Email_Password As String
        Dim Email_Host As String
        Dim Email_Port As Integer

        Dim EntrySelection_Combine_AllCompany As Integer
        Dim InvoicePrint_Format As String
        Dim Jurisdiction As String
        Dim Report_Show_CurrentDate_IN_ToDate As Integer
        Dim Report_Show_CurrentDate_IN_FromDate As Integer

        Dim NegativeStock_Restriction As Integer
        Dim Printing_Show_PrintDialogue As Integer
        Dim OnSave_MoveTo_NewEntry_Status As Integer
        Dim PAYROLLENTRY_Attendance_In_Hours_Status As Integer
        Dim PreviousEntryDate_ByDefault As Integer
        Dim Payroll_Status As Integer
        Dim Company_Selection_For_All_Entry As Integer
        Dim SW_Validity_Date As Date
        Dim OneDrive_Backup_Enabled As Boolean
        Dim Show_Opening_Stock_in_all_accounts_common As Integer
        Dim Show_Points As Boolean
        Dim sales_rate_Edit_Rights As Boolean
        Dim Cost_Rate_in_sales_entry As Boolean
        Dim sales_entry_trasport_text_enabled As Boolean
        Dim Quantity_as_0_decimal As Boolean
        Dim Quantity_as_3_decimal As Boolean
        Dim Quantity_as_2_decimal As Boolean
        Dim Thermal_3_Inch As Boolean
        Dim ItemName_Tamil As Boolean
        Dim ItemName_Telugu As Boolean
        Dim Barcode_Enabled As Boolean
        Dim Label_printer_TVS_LP46_NEO As Boolean
        Dim Label_printer_TSCTE244 As Boolean
        Dim PRN_FILE_TVS_LP_46_COLUMN_2 As Boolean
        Dim PRN_FILE_TVS_LP_46_COLUMN_3 As Boolean
        Dim PRN_FILE_TSCTE244_COLUMN_2 As Boolean
        Dim PRN_FILE_TSCTE244_COLUMN_3 As Boolean
        Dim import_Excel_data As Boolean
        Dim HSN_code_In_Thermal_PrintOut As Boolean
        Dim Price_List_WholesaleRate As Boolean

    End Structure
    Public Shared settings As SettingsDetails

    Public Enum CommonLedger As Integer
        Cash_Ac = 1
        Weaving_Wages_Ac = 2
        Sizing_Charges_Ac = 3
        Godown_Ac = 4
        Transport_Charges_Ac = 7
        TDS_Charges_Ac = 8
        Freight_Charges_Ac = 9
        Salary_Ac = 10
        DD_COMMISSION_Ac = 11
        Stock_In_Hand_Ac = 12
        Profit_Loss_Ac = 13
        RATE_DIFFERENCE_Ac = 14
        CASH_DISCOUNT_Ac = 15
        Agent_Commission_Ac = 16
        Discount_Ac = 17
        Conversion_Bill_Charges_Ac = 18
        Processing_Charges_Ac = 19
        Vat_Ac = 20
        Purchase_Ac = 21
        Sales_Ac = 22
        Service_Replacement_Ac = 52
    End Enum
    Public Structure MasterReturnDetails
        Dim Form_Name As String
        Dim Control_Name As String
        Dim Master_Type As String
        Dim Return_Value As String
    End Structure
    Public Shared Master_Return As MasterReturnDetails

    Public Structure Report_InputDetails
        Dim ReportName As String
        Dim ReportGroupName As String
        Dim ReportHeading As String
        Dim ReportInputs As String
        Dim IsGridReport As Boolean
        Dim Date1 As Date
        Dim Date2 As Date
        Dim IdNo1 As Integer
        Dim IdNo2 As Integer
        Dim Name1 As String
        Dim Name2 As String
    End Structure
    Public Shared RptInputDet As Report_InputDetails

    Public Structure UserDetails
        Dim IdNo As Integer
        Dim Name As String
        Dim Type As String
        Dim AccountPassword As String
        Dim UnAccountPassword As Date
    End Structure
    Public Shared User As UserDetails

    Public Structure UserRightsDetails
        Dim Ledger_Creation As String
        Dim Area_Creation As String
        Dim Item_Creation As String
        Dim ItemGroup_Creation As String
        Dim Unit_Creation As String
        Dim Category_Creation As String
        Dim Variety_Creation As String
        Dim Waste_Creation As String
        Dim Size_Creation As String
        Dim Transport_Creation As String
        Dim Scheme_Master As String
        Dim Ledger_OpeningBalance As String
        Dim Opening_Stock As String
        Dim Salesman_Creation As String
        Dim Bill_Entry As String
        Dim Price_List As String


        Dim Purchase_Entry As String
        Dim Sales_Entry As String
        Dim Sales_Entry_Rate As String
        Dim Tax_Sales_Entry As String
        Dim Labour_Sales_Entry As String
        Dim Delivery_entry As String
        Dim sales_Quotation_Entry As String

        Dim WasteSales_Entry As String

        Dim Knotting_Entry As String
        Dim Knotting_Invoice_Entry As String

        Dim Invoice_Saara_Entry As String
        Dim Delivery_Saara_Entry As String
        Dim Bill_Entry_Saara As String

        Dim Printing_Invoice_Entry As String
        Dim Printing_Order_Entry As String
        Dim Printing_Order_Program_Entry As String

        ' VOUCHERS
        Dim mnu_Voucher_Main As String
        Dim Voucher_Entry As String
        Dim Voucher_Purchase_Entry As String
        Dim Voucher_Sales_Entry As String
        Dim Voucher_Payment_Entry As String
        Dim Voucher_Receipt_Entry As String
        Dim Voucher_Contra_Entry As String
        Dim Voucher_Journal_Entry As String
        Dim Voucher_CreditNote_Entry As String
        Dim Voucher_DebitNote_Entry As String
        Dim Voucher_Petti_Cash_Entry As String
        Dim Voucher_General_Other_PurchaseSales_Entry As String


        ' ACCOUNTS
        Dim mnu_Accounts_Main As String
        Dim Accounts_Ledger_DateWise_Report As String
        Dim Accounts_Ledger_MonthWise_Report As String
        Dim Accounts_SigleLedger_DateWise_Report As String
        Dim Accounts_SigleLedger As String
        Dim Accounts_Ledger_DateRange_Report As String
        Dim Accounts_GroupLedger_OnRange As String
        Dim Accounts_DayOn_Book As String
        Dim Accounts_Opening_TB As String
        Dim Accounts_General_TB As String
        Dim Accounts_Group_TB As String
        Dim Accounts_Final_TB As String
        Dim Accounts_Profit_And_Loss As String
        Dim Accounts_Balance_Sheet As String
        Dim Accounts_Outstanding_BillsWise As String
        Dim Accounts_Outstanding_BillsWise_Simple As String
        Dim Accounts_Outstanding_MonthWise As String
        Dim Accounts_Outstanding_DayWise As String
        Dim Accounts_Customer_Bills As String
        Dim Accounts_VoucherBills As String

        Dim Accounts_Profit_Loss As String
        Dim Accounts_BalanceSheet As String
        Dim Accounts_CustomerBills As String

        Dim Report_Purchase_Register As String
        Dim Report_Sales_Register As String
        Dim Report_Stock_Register As String
        Dim Report_Minimum_Stock_Register As String
        Dim Report_Knotting_Reports As String

        Dim Accounts_Ledger_Report As String
        Dim Accounts_GroupLedger_Report As String
        Dim Accounts_DayBook As String
        Dim Accounts_AllLedger As String
        Dim Accounts_TB As String

        'Dim Accounts_Profit_Loss As String
        'Dim Accounts_BalanceSheet As String
        'Dim Accounts_CustomerBills As String

        'Dim Report_Purchase_Register As String
        'Dim Report_Sales_Register As String
        'Dim Report_Stock_Register As String
        'Dim Report_Minimum_Stock_Register As String
        'Dim Report_Knotting_Reports As String

        ' DISTRIBUTION ENTRY
        Dim Distribution_Purchase_Order_Entry As String
        Dim Distribution_Purchase_Entry As String
        Dim Distribution_Purchase_Return_Entry As String
        Dim Distribution_Sales_Order_Entry As String
        Dim Distribution_Sales_Brand1_Entry As String
        Dim Distribution_Sales_Brand2_Entry As String
        Dim Distribution_Sales_Brand3_Entry As String
        Dim Distribution_Sales_Return_Entry As String
        Dim Distribution_Checking_Inward_Entry As String
        Dim Distribution_Checking_Outward_Entry As String
        Dim Distribution_Simple_Dc As String
        Dim Distribution_Simple_Dc_Return As String
        Dim Distribution_Promotion_Dc_Entry As String
        Dim Distribution_Promotion_Dc_Return_Entry As String
        Dim Distribution_Display_Material_Dc As String
        Dim Distribution_Service_Receipt_Entry As String
        Dim Distribution_Service_Replacement_Entry As String
        Dim Distribution_Credit_Note As String
        Dim Distribution_Transfer_Details As String

        '--- REPORTS
        Dim Reports_Distribution_Purchase_Order As String
        Dim Reports_Distribution_Purchase_Order_Return As String
        Dim Reports_Distribution_Purchase_Register As String
        Dim Reports_Distribution_Sales_Order_Register As String
        Dim Reports_Distribution_Sales_Order_Return As String
        Dim Reports_Distribution_Sales_Register As String
        Dim Reports_Distribution_Sales_Summary As String
        Dim Reports_Distribution_Item_Despatch_Report As String
        Dim Reports_Distribution_Sample_Dc_Register As String
        Dim Reports_Distribution_Sample_Dc_Return_Register As String
        Dim Reports_Distribution_Promotion_Dc_Register As String
        Dim Reports_Distribution_Promotion_Dc_Return_Register As String
        Dim Reports_Distribution_Display_Material_Register As String
        Dim Reports_Distribution_Service_Receipt_Register As String
        Dim Reports_Distribution_Service_Replacement_Register As String
        Dim Reports_Distribution_Stock_Despatch_Register As String
        Dim Reports_Distribution_Stock_Details_Register As String
        Dim Reports_Distribution_Stock_Summary_Register As String
        Dim Reports_Distribution_Stock_Summary_With_Itemgroup_Register As String
        Dim Reports_Distribution_Stock_Summary_Hsn_Code_Register As String
        Dim Reports_Distribution_Gst_Return_Register As String
        Dim Reports_Distribution_Master_Register As String
        Dim Reports_Distribution_Mis_Register As String



    End Structure

    Public Shared UR As UserRightsDetails

    'Public Enum DriveType As Integer
    '    Unknown = 0
    '    NoRoot = 1
    '    Removable = 2
    '    Localdisk = 3
    '    Network = 4
    '    CD = 5
    '    RAMDrive = 6
    'End Enum

    Public Shared Sub Print_To_PrintDocument(ByRef e As System.Drawing.Printing.PrintPageEventArgs, ByVal PrintText As String, ByVal Xaxis As Decimal, ByVal Yaxis As Decimal, ByVal AlignMent As Integer, ByVal DataWidth As Decimal, ByVal DataFont As Font, Optional ByVal BrushColor As Brush = Nothing)
        Dim X As Decimal, Y As Decimal
        Dim strWidth As Decimal, strHeight As Decimal = 0
        Dim vbrushcolor As Brush

        strWidth = e.Graphics.MeasureString(PrintText, DataFont).Width
        'strHeight = e.Graphics.MeasureString(PrintText, DataFont).Height

        If AlignMent = 1 Then
            X = Xaxis - strWidth

        ElseIf AlignMent = 2 Then
            If DataWidth > strWidth Then
                X = Xaxis + (DataWidth - strWidth) / 2
            Else
                X = Xaxis
            End If

        Else
            X = Xaxis

        End If
        Y = Yaxis

        If IsNothing(BrushColor) = False Then
            vbrushcolor = BrushColor


        Else
            vbrushcolor = Brushes.Black
        End If

        e.Graphics.DrawString(PrintText, DataFont, vbrushcolor, X, Y)

    End Sub

    Public Shared Sub Print_To_PrintDocument_Red(ByRef e As System.Drawing.Printing.PrintPageEventArgs, ByVal PrintText As String, ByVal Xaxis As Decimal, ByVal Yaxis As Decimal, ByVal AlignMent As Integer, ByVal DataWidth As Decimal, ByVal DataFont As Font)
        Dim X As Decimal, Y As Decimal
        Dim strWidth As Decimal, strHeight As Decimal = 0

        strWidth = e.Graphics.MeasureString(PrintText, DataFont).Width
        'strHeight = e.Graphics.MeasureString(PrintText, DataFont).Height

        If AlignMent = 1 Then
            X = Xaxis - strWidth

        ElseIf AlignMent = 2 Then
            If DataWidth > strWidth Then
                X = Xaxis + (DataWidth - strWidth) / 2
            Else
                X = Xaxis
            End If

        Else
            X = Xaxis

        End If
        Y = Yaxis

        e.Graphics.DrawString(PrintText, DataFont, Brushes.Red, X, Y)

    End Sub

    Public Shared Sub Print_To_PrintDocument_Green(ByRef e As System.Drawing.Printing.PrintPageEventArgs, ByVal PrintText As String, ByVal Xaxis As Decimal, ByVal Yaxis As Decimal, ByVal AlignMent As Integer, ByVal DataWidth As Decimal, ByVal DataFont As Font)
        Dim X As Decimal, Y As Decimal
        Dim strWidth As Decimal, strHeight As Decimal = 0

        strWidth = e.Graphics.MeasureString(PrintText, DataFont).Width
        'strHeight = e.Graphics.MeasureString(PrintText, DataFont).Height

        If AlignMent = 1 Then
            X = Xaxis - strWidth

        ElseIf AlignMent = 2 Then
            If DataWidth > strWidth Then
                X = Xaxis + (DataWidth - strWidth) / 2
            Else
                X = Xaxis
            End If

        Else
            X = Xaxis

        End If
        Y = Yaxis

        e.Graphics.DrawString(PrintText, DataFont, Brushes.Green, X, Y)

    End Sub

    Public Shared Function Rupees_Converstion(ByVal amt As Single) As String
        Dim A1 As String = ""
        Dim s2 As String = ""
        Dim s3 As String = ""
        Dim Ps1 As Single
        Dim i, j As Integer
        Dim d(100) As String
        Dim Wrd(6) As String
        Dim Sum As Integer

        d(1) = "One"
        d(2) = "Two"
        d(3) = "Three"
        d(4) = "Four"
        d(5) = "Five"
        d(6) = "Six"
        d(7) = "Seven"
        d(8) = "Eight"
        d(9) = "Nine"
        d(10) = "Ten"
        d(11) = "Eleven"
        d(12) = "Twelve"
        d(13) = "Thirteen"
        d(14) = "Fourteen"
        d(15) = "Fifteen"
        d(16) = "Sixteen"
        d(17) = "Seventeen"
        d(18) = "Eighteen"
        d(19) = "Ninteen"
        d(20) = "Twenty"
        d(30) = "Thirty"
        d(40) = "Forty"
        d(50) = "Fifty"
        d(60) = "Sixty"
        d(70) = "Seventy"
        d(80) = "Eighty"
        d(90) = "Ninety"
        Wrd(1) = ""
        Wrd(2) = " Hundred "
        Wrd(3) = " Thousand "
        Wrd(4) = " Lakhs "
        Wrd(5) = " Crores "
        s3 = ""
        Ps1 = Val(Right$(Trim(Format(amt, "###########0.00")), 2))
        If Ps1 <> 0 Then If (Ps1 Mod 10 = 0) Or Ps1 <= 20 Then s3 = d(Ps1) + " Paise" Else s3 = d(Int(Ps1 / 10) * 10) + " " + d(Ps1 Mod 10) + " Paise"
        If Ps1 > 0 Then amt = amt - (Ps1 / 100)
        Do While amt > 0
            i = i + 1
            Sum = amt Mod (IIf((i = 2), 10, 100))
            amt = Int(amt / (IIf((i = 2), 10, 100)))
            If Sum <> 0 Then j = j + 1
            A1 = IIf((j = 2), "And ", "")
            If Sum <> 0 Then If (Sum Mod 10 = 0) Or Sum <= 20 Then s2 = d(Sum) + Wrd(i) + A1 + s2 Else s2 = d(Int(Sum / 10) * 10) + " " + d(Sum Mod 10) + Wrd(i) + A1 + s2
        Loop
        Rupees_Converstion = Trim(s2) + IIf((Len(Trim(s2)) > 0) And (Len(Trim(s3)) > 0), " Rupees And ", " Rupees ") + s3 + " Only"
    End Function

    Public Shared Function Currency_Format(ByVal Value As Double) As String
        Dim s1 As String = ""
        Dim s2 As String = ""
        Dim k As String = ""

        If Value >= 0 Then k = "" Else k = "-"

        s1 = Trim(Format(Math.Abs(Value), "############0.00"))

        Select Case Len(s1)
            Case Is < 9
                s2 = Format(Val(s1), "##,##0.00")
            Case 9, 10
                s2 = Left$(s1, Len(s1) - 8) & "," & Mid$(s1, Len(s1) - 7, 2) & "," & Right$(s1, 6)
            Case 11, 12
                s2 = Left$(s1, Len(s1) - 10) & "," & Mid$(s1, Len(s1) - 9, 2) & "," & Mid$(s1, Len(s1) - 7, 2) & "," & Right$(s1, 6)
            Case 13, 14
                s2 = Left$(s1, Len(s1) - 12) & "," & Mid$(s1, Len(s1) - 11, 2) & "," & Mid$(s1, Len(s1) - 9, 2) & "," & Mid$(s1, Len(s1) - 7, 2) & "," & Right$(s1, 6)
            Case Is > 14
                s2 = Left$(s1, Len(s1) - 14) & "," & Mid$(s1, Len(s1) - 13, 2) & "," & Mid$(s1, Len(s1) - 11, 2) & "," & Mid$(s1, Len(s1) - 9, 2) & "," & Mid$(s1, Len(s1) - 7, 2) & "," & Right$(s1, 6)
        End Select
        Currency_Format = k & Trim(s2)
    End Function

    Public Shared Function get_VoucherType(ByVal VouName As String) As String
        Select Case Trim(LCase(VouName))
            Case "purc"
                get_VoucherType = "Purchase"
            Case "sale"
                get_VoucherType = "Sales"
            Case "rcpt"
                get_VoucherType = "Receipt"
            Case "pymt"
                get_VoucherType = "Payment"
            Case "cntr"
                get_VoucherType = "Contra"
            Case "jrnl"
                get_VoucherType = "Journal"
            Case "crnt"
                get_VoucherType = "Credit Note"
            Case "dbnt"
                get_VoucherType = "Debit Note"
            Case "csrp"
                get_VoucherType = "Cash Receipt"
            Case "cspy"
                get_VoucherType = "Cash Payment"
            Case "ptcs"
                get_VoucherType = "Petti Cash"
            Case "chrt"
                get_VoucherType = "Cheque Return"
            Case Else
                get_VoucherType = ""
        End Select
    End Function

    Public Shared Function Remove_NonCharacters(ByVal Txt As String) As String
        Dim S As String
        Dim I As Integer
        Dim k As Integer

        S = ""
        For I = 1 To Len(Txt)
            k = Asc(Mid(Txt, I, 1))
            If k = 45 Or k = 47 Or (k >= 48 And k <= 57) Or (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Or k = 95 Then
                S = S & Chr(k)
            End If
        Next
        Remove_NonCharacters = S
    End Function

    Public Shared Sub Control_Focus(ByVal Ka As Integer, ByVal Ctrl As Object)
        If Ka = 13 Or Ka = 40 Then SendKeys.Send("{TAB}")
        If Ka = 38 Then SendKeys.Send("+{TAB}")
        If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is ComboBox Then SendKeys.Send("{HOME}+{END}")
    End Sub

    Public Shared Function OrderBy_CodeToValue(ByVal Code As String) As Double
        Dim c As String = ""
        Dim k As Single = 0

        If Val(Code) = 0 Then
            OrderBy_CodeToValue = 0
            Exit Function
        End If

        c = Replace(Code, Val(Code), "")
        k = 0
        If Trim(c) <> "" Then k = Format((Asc(UCase(c)) - 64) / 100, "##############0.00")

        OrderBy_CodeToValue = Format(Val(Code) + k, "##############0.00")

    End Function

    Public Shared Function OrderBy_ValueToCode(ByVal value As Double) As String
        Dim c As String = ""
        Dim k As Single = 0

        If Val(value) = 0 Then
            OrderBy_ValueToCode = ""
            Exit Function
        End If

        k = Format(Val(value), "#####0.00") - Int(Val(value))

        c = ""
        If Val(k) > 0 Then c = UCase(Chr(64 + k))

        OrderBy_ValueToCode = Int(Val(value)) & c

    End Function

    Public Shared Function Accept_NumericOnly(ByVal KeyAscii_Value As Integer) As Integer
        Accept_NumericOnly = 0
        If (KeyAscii_Value >= 48 And KeyAscii_Value <= 57) Or KeyAscii_Value = 45 Or KeyAscii_Value = 46 Or KeyAscii_Value = 13 Or KeyAscii_Value = 8 Or KeyAscii_Value = 9 Then
            Accept_NumericOnly = KeyAscii_Value
        End If
    End Function

    Public Function Accept_AlphaNumericOnly(ByVal KeyAscii_Value As Integer) As Integer
        Accept_AlphaNumericOnly = 0
        If (KeyAscii_Value <> 39 And (KeyAscii_Value >= 32 And KeyAscii_Value <= 57)) Or (KeyAscii_Value >= 65 And KeyAscii_Value <= 90) Or (KeyAscii_Value >= 97 And KeyAscii_Value <= 122) Or KeyAscii_Value = 13 Or KeyAscii_Value = 8 Or KeyAscii_Value = 9 Or KeyAscii_Value = 92 Then
            Accept_AlphaNumericOnly = KeyAscii_Value
        End If
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


    Public Shared Function CompanyGroup_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCompanyGroup_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCompanyGroup_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select CompanyGroup_IdNo from " & Trim(Common_Procedures.CompanyDetailsDataBaseName) & "..CompanyGroup_Head where CompanyGroup_Name = '" & Trim(vCompanyGroup_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vCompanyGroup_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCompanyGroup_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        CompanyGroup_NameToIdNo = Val(vCompanyGroup_ID)

    End Function

    Public Shared Function CompanyGroup_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCompanyGroup_ID As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCompanyGroup_Nm As String

        Da = New SqlClient.SqlDataAdapter("select CompanyGroup_Name from  " & Trim(Common_Procedures.CompanyDetailsDataBaseName) & "..CompanyGroup_Head where CompanyGroup_IdNo = " & Str(Val(vCompanyGroup_ID)), Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vCompanyGroup_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCompanyGroup_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        CompanyGroup_IdNoToName = Trim(vCompanyGroup_Nm)

    End Function


    Public Shared Function Company_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCompany_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCompany_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Company_IdNo from Company_Head where Company_Name = '" & Trim(vCompany_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vCompany_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCompany_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Company_NameToIdNo = Val(vCompany_ID)

    End Function

    Public Shared Function Company_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCompany_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCompany_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Company_Name from Company_Head where Company_IdNo = " & Str(Val(vCompany_ID)), Cn1)
        Da.Fill(Dt)

        vCompany_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCompany_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Company_IdNoToName = Trim(vCompany_Nm)

    End Function

    Public Shared Function Company_ShortNameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCompany_ShtNm As String, Optional ByVal vCmpGrpIdNo As Integer = 0) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCompany_ID As Integer
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select Company_IdNo from " & Trim(vDB_Name) & "Company_Head where Company_ShortName = '" & Trim(vCompany_ShtNm) & "'", Cn1)
        Dt = New DataTable
        Da.Fill(Dt)

        vCompany_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCompany_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Company_ShortNameToIdNo = Val(vCompany_ID)

    End Function

    Public Shared Function Company_IdNoToShortName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCompany_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCompany_ShtNm As String

        Da = New SqlClient.SqlDataAdapter("select Company_ShortName from Company_Head where Company_IdNo = " & Str(Val(vCompany_ID)), Cn1)
        Da.Fill(Dt)

        vCompany_ShtNm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCompany_ShtNm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Company_IdNoToShortName = Trim(vCompany_ShtNm)

    End Function

    Public Shared Function AccountsGroup_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vAccountsGroup_Nm As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vAccountsGroup_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select AccountsGroup_IdNo from AccountsGroup_Head where AccountsGroup_Name = '" & Trim(vAccountsGroup_Nm) & "'", Cn1)
        Da.Fill(Dt)

        vAccountsGroup_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vAccountsGroup_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        AccountsGroup_NameToIdNo = Val(vAccountsGroup_ID)

    End Function

    Public Shared Sub get_SMS_Provider_Details(ByVal Cn1 As SqlClient.SqlConnection, ByVal CompIDNo As Integer, ByRef SMS_SenderID As String, ByRef SMS_Key As String, ByRef SMS_RouteID As String, ByRef SMS_Type As String)
        Dim Da1 As New SqlClient.SqlDataAdapter
        Dim Dt1 As New DataTable
        Dim S_SenderID As String = ""
        Dim S_Key As String = ""
        Dim S_RouteID As String = ""
        Dim S_Type As String = ""

        Try
            If Val(CompIDNo) = 0 Then
                SMS_SenderID = Trim(Common_Procedures.settings.SMS_Provider_SenderID)
                SMS_Key = Trim(Common_Procedures.settings.SMS_Provider_Key)
                SMS_RouteID = Trim(Common_Procedures.settings.SMS_Provider_RouteID)
                SMS_Type = Trim(Common_Procedures.settings.SMS_Provider_Type)

            Else

                S_SenderID = ""
                S_Key = ""
                S_RouteID = ""
                S_Type = ""

                Da1 = New SqlClient.SqlDataAdapter("select * from company_head where company_idno = " & Str(Val(CompIDNo)), Cn1)
                Dt1 = New DataTable
                Da1.Fill(Dt1)
                If Dt1.Rows.Count > 0 Then
                    If IsDBNull(Dt1.Rows(0).Item("SMS_Provider_SenderID").ToString) = False Then
                        S_SenderID = Trim(Dt1.Rows(0).Item("SMS_Provider_SenderID").ToString)
                    End If
                    If IsDBNull(Dt1.Rows(0).Item("SMS_Provider_Key").ToString) = False Then
                        S_Key = Trim(Dt1.Rows(0).Item("SMS_Provider_Key").ToString)
                    End If
                    If IsDBNull(Dt1.Rows(0).Item("SMS_Provider_RouteID").ToString) = False Then
                        S_RouteID = Trim(Dt1.Rows(0).Item("SMS_Provider_RouteID").ToString)
                    End If
                    If IsDBNull(Dt1.Rows(0).Item("SMS_Provider_Type").ToString) = False Then
                        S_Type = Trim(Dt1.Rows(0).Item("SMS_Provider_Type").ToString)
                    End If
                End If
                Dt1.Clear()

                If Trim(S_SenderID) <> "" And Trim(S_Key) <> "" And Trim(S_RouteID) <> "" Then
                    SMS_SenderID = Trim(S_SenderID)
                    SMS_Key = Trim(S_Key)
                    SMS_RouteID = Trim(S_RouteID)
                    SMS_Type = Trim(S_Type)

                Else
                    SMS_SenderID = Trim(Common_Procedures.settings.SMS_Provider_SenderID)
                    SMS_Key = Trim(Common_Procedures.settings.SMS_Provider_Key)
                    SMS_RouteID = Trim(Common_Procedures.settings.SMS_Provider_RouteID)
                    SMS_Type = Trim(Common_Procedures.settings.SMS_Provider_Type)

                End If

            End If

            Dt1.Dispose()
            Da1.Dispose()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "ERROR GETTING SMS PROVIDER DETAILS...", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Shared Function AccountsGroup_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vAccountsGroup_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vAccountsGroup_Nm As String

        Da = New SqlClient.SqlDataAdapter("select AccountsGroup_Name from AccountsGroup_Head where AccountsGroup_IdNo = " & Str(Val(vAccountsGroup_ID)), Cn1)
        Da.Fill(Dt)

        vAccountsGroup_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vAccountsGroup_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        AccountsGroup_IdNoToName = Trim(vAccountsGroup_Nm)

    End Function


    Public Shared Function Ledger_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vLed_IdNo As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vLed_Name As String

        Da = New SqlClient.SqlDataAdapter("select Ledger_Name from Ledger_Head where Ledger_IdNo = " & Str(Val(vLed_IdNo)), Cn1)
        Da.Fill(Dt)

        vLed_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vLed_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Ledger_IdNoToName = Trim(vLed_Name)

    End Function

    Public Shared Function Ledger_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vLed_Name As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vLed_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Ledger_IdNo from Ledger_Head where Ledger_Name = '" & Trim(vLed_Name) & "'", Cn1)
        Da.Fill(Dt)

        vLed_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vLed_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Ledger_NameToIdNo = Val(vLed_ID)

    End Function
    Public Shared Function State_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSte_IdNo As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vste_Name As String

        Da = New SqlClient.SqlDataAdapter("select State_Name from State_Head where State_IdNo = " & Str(Val(vSte_IdNo)), Cn1)
        Da.Fill(Dt)

        vste_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vste_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        State_IdNoToName = Trim(vste_Name)

    End Function

    Public Shared Function State_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSte_Name As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSte_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select State_IdNo from State_Head where State_Name = '" & Trim(vSte_Name) & "'", Cn1)
        Da.Fill(Dt)

        vSte_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSte_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        State_NameToIdNo = Val(vSte_ID)

    End Function
    Public Shared Function Salesman_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSte_IdNo As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vste_Name As String

        Da = New SqlClient.SqlDataAdapter("select Salesman_Name from Salesman_Head where Salesman_Idno = " & Str(Val(vSte_IdNo)), Cn1)
        Da.Fill(Dt)

        vste_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vste_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Salesman_IdNoToName = Trim(vste_Name)

    End Function

    Public Shared Function Salesman_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSte_Name As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSte_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Salesman_Idno from Salesman_Head where Salesman_Name = '" & Trim(vSte_Name) & "'", Cn1)
        Da.Fill(Dt)

        vSte_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSte_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Salesman_NameToIdNo = Val(vSte_ID)

    End Function
    Public Shared Function Ledger_AlaisNameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vLed_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vLed_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Ledger_IdNo from Ledger_AlaisHead where Ledger_DisplayName = '" & Trim(vLed_Name) & "' Order by Ledger_IdNo", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vLed_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vLed_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Ledger_AlaisNameToIdNo = Val(vLed_ID)

    End Function
    Public Shared Function Ledger_IdnoToAlaisName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vLed_idno As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vLed_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Ledger_DisplayName from Ledger_AlaisHead where Ledger_IdNo  = " & Val(vLed_idno) & " Order by Ledger_DisplayName", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vLed_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vLed_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Ledger_IdnoToAlaisName = Trim(vLed_Nm)

    End Function

    Public Shared Function AccountsGroup_NameToCode(ByVal Cn1 As SqlClient.SqlConnection, ByVal vAccountsGroup_Nm As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vAccountsGroup_Code As Integer

        Da = New SqlClient.SqlDataAdapter("select Parent_Idno from AccountsGroup_Head where AccountsGroup_Name = '" & Trim(vAccountsGroup_Nm) & "'", Cn1)
        Da.Fill(Dt)

        vAccountsGroup_Code = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vAccountsGroup_Code = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        AccountsGroup_NameToCode = Val(vAccountsGroup_Code)

    End Function

    Public Shared Function AccountsGroup_IdNoToCode(ByVal Cn1 As SqlClient.SqlConnection, ByVal vAccountsGroup_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vAccountsGroup_Code As String

        Da = New SqlClient.SqlDataAdapter("select Parent_Idno from AccountsGroup_Head where AccountsGroup_IdNo = " & Str(Val(vAccountsGroup_ID)), Cn1)
        Da.Fill(Dt)

        vAccountsGroup_Code = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vAccountsGroup_Code = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        AccountsGroup_IdNoToCode = Trim(vAccountsGroup_Code)

    End Function

    Public Shared Function Item_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItem_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal vCmpGrpIdNo As Integer = 0) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItem_ID As Integer
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select Item_IdNo from " & Trim(vDB_Name) & "Item_Head where Item_Name = '" & Trim(vItem_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vItem_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItem_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Item_NameToIdNo = Val(vItem_ID)

    End Function
    Public Shared Function SerialName_ToSerialId(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSerial_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal vCmpGrpIdNo As Integer = 0) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItem_ID As Integer
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select serial_idno from " & Trim(vDB_Name) & "Serial_Head where  Serial_Name = '" & Trim(vSerial_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vItem_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItem_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        SerialName_ToSerialId = Val(vItem_ID)

    End Function
    Public Shared Function Item_IdToItemGroupIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItem_Idno As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal vCmpGrpIdNo As Integer = 0) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItem_ID As Integer
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select ItemGroup_IdNo from " & Trim(vDB_Name) & "Item_Head where Item_IdNo = " & Val(vItem_Idno) & "", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vItem_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItem_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Item_IdToItemGroupIdNo = Val(vItem_ID)

    End Function
    Public Shared Function Item_CodeToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItem_Code As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItem_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Item_IdNo from Item_Head where Item_Code = '" & Trim(vItem_Code) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vItem_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItem_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Item_CodeToIdNo = Val(vItem_ID)

    End Function

    Public Shared Function Item_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItem_ID As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItem_Name As String

        Da = New SqlClient.SqlDataAdapter("select Item_Name from Item_Head where Item_IdNo = " & Str(Val(vItem_ID)), Cn1)
        Dt = New DataTable
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vItem_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItem_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Item_IdNoToName = Trim(vItem_Name)

    End Function

    Public Shared Function Unit_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vUnit_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal vCmpGrpIdNo As Integer = 0) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vUnit_ID As Integer
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select Unit_IdNo from " & Trim(vDB_Name) & "Unit_Head where Unit_Name = '" & Trim(vUnit_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vUnit_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vUnit_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Unit_NameToIdNo = Val(vUnit_ID)

    End Function


    Public Shared Function Unit_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vUnit_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vUnit_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Unit_Name from Unit_Head where Unit_IdNo = " & Str(Val(vUnit_ID)), Cn1)
        Da.Fill(Dt)

        vUnit_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vUnit_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Unit_IdNoToName = Trim(vUnit_Nm)

    End Function

    Public Shared Function ItemGroup_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItemGroup_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItemGroup_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select ItemGroup_IdNo from ItemGroup_Head where ItemGroup_Name = '" & Trim(vItemGroup_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)


        vItemGroup_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItemGroup_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        ItemGroup_NameToIdNo = Val(vItemGroup_ID)

    End Function

    Public Shared Function Item_NameToItemGroupIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItem_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItemGroup_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select ItemGroup_IdNo from Item_Head where Item_Name = '" & Trim(vItem_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)


        vItemGroup_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItemGroup_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Item_NameToItemGroupIdNo = Val(vItemGroup_ID)

    End Function
    Public Shared Function ItemHSNCODE_ToGroupIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItemGroup_HSNCode As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItemGroup_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select ItemGroup_IdNo from ItemGroup_Head where Item_HSN_Code = '" & Trim(vItemGroup_HSNCode) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)


        vItemGroup_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItemGroup_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        ItemHSNCODE_ToGroupIdNo = Val(vItemGroup_ID)

    End Function
    Public Shared Function Price_List_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vPrice_List_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vPrice_List_Name As String

        Da = New SqlClient.SqlDataAdapter("select Price_List_Name from Price_List_Head where Price_List_IdNo = " & Str(Val(vPrice_List_ID)), Cn1)
        Da.Fill(Dt)

        vPrice_List_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vPrice_List_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Price_List_IdNoToName = Trim(vPrice_List_Name)

    End Function
    Public Shared Function Price_List_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vPrice_List_Name As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vPrice_List_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Price_List_IdNo from Price_List_Head where Price_List_Name = '" & Trim(vPrice_List_Name) & "'", Cn1)
        Da.Fill(Dt)

        vPrice_List_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vPrice_List_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Price_List_NameToIdNo = Val(vPrice_List_ID)

    End Function

    Public Shared Function ItemGroup_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItemGroup_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vItemGroup_Name As String

        Da = New SqlClient.SqlDataAdapter("select ItemGroup_Name from ItemGroup_Head where ItemGroup_IdNo = " & Str(Val(vItemGroup_ID)), Cn1)
        Da.Fill(Dt)

        vItemGroup_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vItemGroup_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        ItemGroup_IdNoToName = Trim(vItemGroup_Name)

    End Function

    Public Shared Function Variety_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vVariety_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vVariety_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Variety_IdNo from Variety_Head where Variety_Name = '" & Trim(vVariety_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vVariety_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vVariety_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Variety_NameToIdNo = Val(vVariety_ID)

    End Function

    Public Shared Function Variety_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vVariety_ID As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vVariety_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Variety_Name from Variety_Head where Variety_IdNo = " & Str(Val(vVariety_ID)), Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vVariety_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vVariety_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Variety_IdNoToName = Trim(vVariety_Nm)

    End Function

    Public Shared Function Area_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vArea_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vArea_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Area_IdNo from Area_Head where Area_Name = '" & Trim(vArea_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vArea_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vArea_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Area_NameToIdNo = Val(vArea_ID)

    End Function

    Public Shared Function Area_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vArea_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vArea_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Area_Name from Area_Head where Area_IdNo = " & Str(Val(vArea_ID)), Cn1)
        Da.Fill(Dt)

        vArea_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vArea_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Area_IdNoToName = Trim(vArea_Nm)

    End Function
    Public Shared Function Machine_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vMachine_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vMachine_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Machine_IdNo from Machine_Head where Machine_Name = '" & Trim(vMachine_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vMachine_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vMachine_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Machine_NameToIdNo = Val(vMachine_ID)

    End Function

    Public Shared Function Machine_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vMachine_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vMachine_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Machine_Name from Machine_Head where Machine_IdNo = " & Str(Val(vMachine_ID)), Cn1)
        Da.Fill(Dt)

        vMachine_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vMachine_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Machine_IdNoToName = Trim(vMachine_Nm)

    End Function
    Public Shared Function Cetegory_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCetegory_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCetegory_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Cetegory_IdNo from Cetegory_Head where Cetegory_Name = '" & Trim(vCetegory_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vCetegory_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCetegory_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Cetegory_NameToIdNo = Val(vCetegory_ID)

    End Function

    Public Shared Function Cetegory_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCetegory_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCetegory_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Cetegory_Name from Cetegory_Head where Cetegory_IdNo = " & Str(Val(vCetegory_ID)), Cn1)
        Da.Fill(Dt)

        vCetegory_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCetegory_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Cetegory_IdNoToName = Trim(vCetegory_Nm)

    End Function
    Public Shared Function Colour_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vColour_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vColour_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Colour_IdNo from Colour_Head where Colour_Name = '" & Trim(vColour_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vColour_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vColour_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Colour_NameToIdNo = Val(vColour_ID)

    End Function

    Public Shared Function Colour_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vColour_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vColour_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Colour_Name from Colour_Head where Colour_IdNo = " & Str(Val(vColour_ID)), Cn1)
        Da.Fill(Dt)

        vColour_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vColour_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Colour_IdNoToName = Trim(vColour_Nm)

    End Function

    Public Shared Function Gender_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vGender_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vGender_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Gender_IdNo from Gender_Head where Gender_Name = '" & Trim(vGender_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vGender_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vGender_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Gender_NameToIdNo = Val(vGender_ID)

    End Function

    Public Shared Function Gender_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vGender_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vGender_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Gender_Name from Gender_Head where Gender_IdNo = " & Str(Val(vGender_ID)), Cn1)
        Da.Fill(Dt)

        vGender_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vGender_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Gender_IdNoToName = Trim(vGender_Nm)

    End Function
    Public Shared Function Style_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vStyle_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vStyle_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Style_IdNo from Style_Head where Style_Name = '" & Trim(vStyle_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vStyle_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vStyle_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Style_NameToIdNo = Val(vStyle_ID)

    End Function

    Public Shared Function Style_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vStyle_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vStyle_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Style_Name from Style_Head where Style_IdNo = " & Str(Val(vStyle_ID)), Cn1)
        Da.Fill(Dt)

        vStyle_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vStyle_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Style_IdNoToName = Trim(vStyle_Nm)

    End Function
    Public Shared Function Sleeve_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSleeve_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSleeve_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Sleeve_IdNo from Sleeve_Head where Sleeve_Name = '" & Trim(vSleeve_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vSleeve_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSleeve_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Sleeve_NameToIdNo = Val(vSleeve_ID)

    End Function

    Public Shared Function Sleeve_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSleeve_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSleeve_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Sleeve_Name from Sleeve_Head where Sleeve_IdNo = " & Str(Val(vSleeve_ID)), Cn1)
        Da.Fill(Dt)

        vSleeve_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSleeve_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Sleeve_IdNoToName = Trim(vSleeve_Nm)

    End Function
    Public Shared Function Design_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vDesign_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vDesign_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Design_IdNo from Design_Head where Design_Name = '" & Trim(vDesign_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vDesign_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vDesign_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Design_NameToIdNo = Val(vDesign_ID)

    End Function

    Public Shared Function Design_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vDesign_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vDesign_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Design_Name from Design_Head where Design_IdNo = " & Str(Val(vDesign_ID)), Cn1)
        Da.Fill(Dt)

        vDesign_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vDesign_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Design_IdNoToName = Trim(vDesign_Nm)

    End Function
    Public Shared Function Waste_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vWaste_Name As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vWaste_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Waste_IdNo from Waste_Head where Waste_Name = '" & Trim(vWaste_Name) & "'", Cn1)
        Da.Fill(Dt)

        vWaste_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vWaste_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Waste_NameToIdNo = Val(vWaste_ID)

    End Function

    Public Shared Function Waste_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vWaste_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vWaste_Name As String

        Da = New SqlClient.SqlDataAdapter("select Waste_Name from Waste_Head where Waste_IdNo = " & Str(Val(vWaste_ID)), Cn1)
        Da.Fill(Dt)

        vWaste_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vWaste_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Waste_IdNoToName = Trim(vWaste_Name)

    End Function

    Public Shared Function Transport_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vTransport_Nm As String) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vTransport_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Transport_IdNo from Transport_Head where Transport_Name = '" & Trim(vTransport_Nm) & "'", Cn1)
        Da.Fill(Dt)

        vTransport_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vTransport_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Transport_NameToIdNo = Val(vTransport_ID)

    End Function

    Public Shared Function Transport_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vTransport_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vTransport_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Transport_Name from Transport_Head where Transport_IdNo = " & Str(Val(vTransport_ID)), Cn1)
        Da.Fill(Dt)

        vTransport_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vTransport_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Transport_IdNoToName = Trim(vTransport_Nm)

    End Function

    Public Shared Function Size_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSize_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSize_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Size_IdNo from Size_Head where Size_Name = '" & Trim(vSize_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vSize_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSize_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Size_NameToIdNo = Val(vSize_ID)

    End Function

    Public Shared Function Size_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSize_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSize_Name As String

        Da = New SqlClient.SqlDataAdapter("select Size_Name from Size_Head where Size_IdNo = " & Str(Val(vSize_ID)), Cn1)
        Da.Fill(Dt)

        vSize_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSize_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Size_IdNoToName = Trim(vSize_Name)

    End Function

    Public Shared Function get_FieldValue(ByVal Cn1 As SqlClient.SqlConnection, ByVal vTable_name As String, ByVal vField_Name As String, ByVal vCondition As String, Optional ByVal vCompany_ID As Integer = 0, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim RetVal As String
        Dim SqlCondt As String

        SqlCondt = ""
        If Trim(vCondition) <> "" Then
            SqlCondt = "(" & Trim(vCondition) & ")"
        End If

        If Val(vCompany_ID) <> 0 Then
            SqlCondt = Trim(SqlCondt) & IIf(Trim(SqlCondt) <> "", " and ", "") & " Company_IdNo = " & Str(Val(vCompany_ID))
        End If

        Da = New SqlClient.SqlDataAdapter("select " & vField_Name & " from " & vTable_name & IIf(Trim(SqlCondt) <> "", " Where ", "") & SqlCondt, Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        RetVal = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                RetVal = Dt.Rows(0)(0).ToString
            End If
        End If

        Dt.Clear()
        Dt.Dispose()
        Da.Dispose()

        get_FieldValue = RetVal

    End Function


    Public Shared Function get_MaxIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vTable_name As String, ByVal vField_name As String, ByVal vCondition As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim MxId As Integer

        Da = New SqlClient.SqlDataAdapter("select max(" & vField_name & ") from " & vTable_name & IIf(Trim(vCondition) <> "", " Where ", "") & vCondition, Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        MxId = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                MxId = Val(Dt.Rows(0)(0).ToString)
            End If
        End If
        MxId = MxId + 1

        Dt.Dispose()
        Da.Dispose()

        get_MaxIdNo = Val(MxId)

    End Function

    Public Shared Function get_MaxCode(ByVal Cn1 As SqlClient.SqlConnection, ByVal vTable_name As String, ByVal vPK_Fieldname As String, ByVal vOrderBy_Fieldname As String, ByVal vCondition As String, ByVal vCompany_ID As Integer, ByVal vFinYr As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim MxId As Double
        Dim SqlCondt As String

        SqlCondt = ""
        If Trim(vCondition) <> "" Then
            SqlCondt = "(" & Trim(vCondition) & ")"
        End If
        If Val(vCompany_ID) <> 0 Then
            SqlCondt = Trim(SqlCondt) & IIf(Trim(SqlCondt) <> "", " and ", "") & " Company_IdNo = " & Str(Val(vCompany_ID))
        End If

        If Trim(vFinYr) <> "" Then
            SqlCondt = Trim(SqlCondt) & IIf(Trim(SqlCondt) <> "", " and ", "") & " " & vPK_Fieldname & " like '%" & Trim(vFinYr) & "'"
        End If

        Da = New SqlClient.SqlDataAdapter("select max(" & vOrderBy_Fieldname & ") from " & vTable_name & IIf(Trim(SqlCondt) <> "", " Where ", "") & SqlCondt, Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        MxId = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                MxId = Int(Val(Dt.Rows(0)(0).ToString))
            End If
        End If
        MxId = MxId + 1

        Dt.Clear()
        Dt.Dispose()
        Da.Dispose()

        get_MaxCode = Trim(Val(MxId))

    End Function

    Public Shared Function get_Item_CurrentStock(ByVal Cn1 As SqlClient.SqlConnection, ByVal vComp_IdNo As Integer, ByVal vItem_IdNo As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal ExcludeCode As String = "") As Decimal
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim CurStk As Decimal = 0

        If Val(Common_Procedures.settings.Show_Opening_Stock_in_all_accounts_common) = 1 Then
            Da = New SqlClient.SqlDataAdapter("select sum(Quantity) from Item_Processing_Details where  Item_IdNo = " & Str(Val(vItem_IdNo)) & IIf(Len(Trim(ExcludeCode)) > 0, " and not Reference_Code in (" & ExcludeCode & ")", ""), Cn1)
        Else
            Da = New SqlClient.SqlDataAdapter("select sum(Quantity) from Item_Processing_Details where Company_IdNo = " & Str(Val(vComp_IdNo)) & " and Item_IdNo = " & Str(Val(vItem_IdNo)) & IIf(Len(Trim(ExcludeCode)) > 0, " and not Reference_Code in (" & ExcludeCode & ")", ""), Cn1)
        End If


        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Dt = New DataTable
        Da.Fill(Dt)

        CurStk = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                CurStk = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        get_Item_CurrentStock = Val(CurStk)

    End Function

    Public Shared Sub Default_GroupHead_Updation(ByVal Cn1 As SqlClient.SqlConnection)
        Dim cmd As New SqlClient.SqlCommand

        cmd.Connection = Cn1

        cmd.CommandText = "delete from AccountsGroup_Head where AccountsGroup_IdNo <= 30"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (0, '',          '',       '',    0, '',       0,  0,      '',                         ''      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (1, 'BRANCH / DIVISION',          'BRANCHDIVISION',       'BRANCH / DIVISION',    1, '~1~',       0,  7,      '',                         'SUBSIDIARY FIRMS'      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (2, 'CAPITAL ACCOUNT',            'CAPITALACCOUNT',       'CAPITAL ACCOUNT',      1,  '~2~',      0,  1,      '',                         'EQUITY'                )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (3, 'RESERVESE & SURPLUS',        'RESERVESESURPLUS',     'CAPITAL ACCOUNT',      1,  '~3~2~',    0,  1.1,    '',                         'RETAINED EARNINGS'     )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (4, 'CURRENT ASSETS',             'CURRENTASSETS',        'CURRENT ASSETS',       1,  '~4~',      0,  6,      'CURRENT ASSETS',           ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (5, 'BANK ACCOUNTS',              'BANKACCOUNTS',         'CURRENT ASSETS',       1,  '~5~4~',    0,  6.7,    'CURRENT ASSETS',           ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (6, 'CASH-IN-HAND',               'CASHINHAND',           'CURRENT ASSETS',       1,  '~6~4~',    0,  6.6,    'CURRENT ASSETS',           ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (7, 'DEPOSITS (ASSET)',           'DEPOSITSASSET',        'CURRENT ASSETS',       1,  '~7~4~',    0,  6.2,    'CURRENT ASSETS',           ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (8, 'LOANS & ADVANCES (ASSET)',   'LOANSADVANCESASSET',   'CURRENT ASSETS',       1,  '~8~4~',    0,  6.3,    'CURRENT ASSETS',           ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (9, 'STOCK-IN-HAND',              'STOCKINHAND',          'CURRENT ASSETS',       1,  '~9~4~',    0,  6.1,    'CURRENT ASSETS',           ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (10, 'SUNDRY DEBTORS',            'SUNDRYDEBTORS',        'CURRENT ASSETS',       1,  '~10~4~',   0,  6.5,    'CURRENT ASSETS',           'ACCOUNTS RECEIVABLE'   )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (11, 'CURRENT LIABILITIES',       'CURRENTLIABILITIES',   'CURRENT LIABILITIES',  1,  '~11~',     0,  3,      'CURRENT LIABILITIES',      ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (12, 'DUTIES & TAXES',            'DUTIESTAXES',          'CURRENT LIABILITIES',  1,  '~12~11~',  0,  3.2,    'CURRENT LIABILITIES',      ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (13, 'PROVISIONS',                'PROVISIONS',           'CURRENT LIABILITIES',  1,  '~13~11~',  0,  3.3,    'CURRENT LIABILITIES',      ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (14, 'SUNDRY CREDITORS',          'SUNDRYCREDITORS',      'CURRENT LIABILITIES',  1,  '~14~11~',  0,  3.4,    'CURRENT LIABILITIES',      'ACCOUNTS PAYABLE'      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (15, 'EXPENSES (DIRECT)',         'EXPENSESDIRECT',       'EXPENSES (DIRECT)',    1,  '~15~18~',  1,  13,     'EXPENDITURE ACCOUNT',      'MFG./TRDG. EXPENSES'   )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (16, 'EXPENSES (INDIRECT)',       'EXPENSESINDIRECT',     'EXPENSES (INDIRECT)',  1,  '~16~18~',  1,  15,     'EXPENDITURE ACCOUNT',      'ADMIN. EXPENSES'       )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (17, 'FIXED ASSETS',              'FIXEDASSETS',          'FIXED ASSETS',         1,  '~17~',     0,  4,      '',                         'IMMOVABLE PROPERTIES'  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (18, 'REVENUE ACCOUNTS',          'REVENUEACCOUNTS',      'REVENUE ACCOUNTS',     1,  '~18~',     0,  18,     '',                         ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (19, 'INCOME (REVENUE)',          'INCOMEREVENUE',        'INCOME (REVENUE)',     1,  '~19~18~',  1,  12,     'REVENUE ACCOUNTS',         ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (20, 'INVESTMENTS',               'INVESTMENTS',          'INVESTMENTS',          1,  '~20~',     0,  5,      '',                         ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (21, 'LOANS (LIABILITY)',         'LOANSLIABILITY',       'LOANS (LIABILITY)',    1,  '~21~',     0,  2,      '',                         ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (23, 'BANK OCC A/C',              'BANKOCCAC',            'LOANS (LIABILITY)',    1,  '~23~21~',  0,  2.1,    'LOANS (LIABILITY)',        ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (24, 'SECURED LOANS',             'SECUREDLOANS',         'LOANS (LIABILITY)',    1,  '~24~21~',  0,  2.2,    'LOANS (LIABILITY)',        ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (25, 'UNSECURED LOANS',           'UNSECUREDLOANS',       'LOANS (LIABILITY)',    1,  '~25~21~',  0,  2.3,    'LOANS (LIABILITY)',        ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (26, 'MISC.EXPENSES (ASSET)',     'MISCEXPENSESASSET',    'MISC.EXPENSES (ASSET)',1,  '~26~',     0,  8,      'Misc Expenses (ASSET)',    'Misc Expenses (ASSET)' )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (27, 'PURCHASE ACCOUNT',          'PURCHASEACCOUNT',      'PURCHASE ACCOUNT',     1,  '~27~18~',  1,  11,     'REVENUE ACCOUNTS',         ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (28, 'SALES ACCOUNT',             'SALESACCOUNT',         'SALES ACCOUNT',        1,  '~28~18~',  1,  10,     'REVENUE ACCOUNTS',         ''                      )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (29, 'SUSPENSE ACCOUNT',          'SUSPENSEACCOUNT',      'SUSPENSE ACCOUNT',     1,  '~29~',     0,  9,      '',                         'TEMPORARY A/CS'        )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (30, 'PROFIT & LOSS A/C',         'PROFITLOSSAC',         'PROFIT & LOSS A/C',    1,  '~30~',     0,  16,     'Profit & Loss A/c',        'Profit & Loss Account' )"
        cmd.ExecuteNonQuery()

    End Sub

    Public Shared Sub Default_LedgerHead_Updation(ByVal Cn1 As SqlClient.SqlConnection)
        Dim cmd As New SqlClient.SqlCommand
        Dim Da1 As New SqlClient.SqlDataAdapter
        Dim Dt1 As New DataTable
        Dim i As Integer = 0
        Dim vBankAcName As String = ""

        cmd.Connection = Cn1

        vBankAcName = ""
        If Trim(UCase(Common_Procedures.settings.CustomerCode)) = "1167" Then  '--- F Fashion
            vBankAcName = Common_Procedures.Ledger_IdNoToName(Cn1, 51)
        End If

        cmd.CommandText = "delete from Ledger_Head where Ledger_IdNo <= 100"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (0,     '',                       '',                 '',                     '',         0,      0,      '',         '',                 '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (1,     'CASH A/C',               'CASHAC',           'CASH A/C',             '',         0,      6,      '~6~4~',    'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (4,     'GODOWN',                 'GODOWN',           'GODOWN',               '',         0,      9,      '~9~4~',    'BALANCE ONLY',     'GODOWN', '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (8,     'TDS CHARGES',            'TDSCHARGES',       'TDS CHARGES',          '',         0,      12,     '~12~11~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()

        If Trim(Common_Procedures.settings.CustomerCode) = "1011" Then
            cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (9,    'BATTERY CHARGING A/C',    'BATTERYCHARGINGAC', 'BATTERY CHARGING A/C', '',         0,      19,     '~19~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
            cmd.ExecuteNonQuery()

        Else
            cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (9,     'FREIGHT CHARGES',        'FREIGHTCHARGES',   'FREIGHT CHARGES',       '',         0,      16,     '~16~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
            cmd.ExecuteNonQuery()

        End If

        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (12,    'STOCK-IN-HAND',          'STOCKINHAND',      'STOCK-IN-HAND',        '',         0,      9,      '~9~4~',    'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (13,    'PROFIT & LOSS A/C',      'PROFITLOSSAC',     'PROFIT & LOSS A/C',    '',         0,      30,     '~30~',     'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (17,    'DISCOUNT A/C',           'DISCOUNTAC',       'DISCOUNT A/C',         '',         0,      16,     '~16~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (20,    'VAT A/C',                'VATAC',            'VAT A/C',              '',         0,      12,     '~12~11~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (21,    'PURCHASE A/C',           'PURCHASEAC',       'PURCHASE A/C',         '',         0,      27,     '~27~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        If Trim(UCase(Common_Procedures.settings.CustomerCode)) = "1201" Then  '--- SWASTHICK KNITT (Tirupur ) Embroidery
            'cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (22,    'EMBROIDERY CHARGES A/C',              'EMBROIDERYCHARGESAC',          'EMBROIDERYCHARGESA/C',            '',         0,      28,     '~28~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
            'cmd.ExecuteNonQuery()
        Else
            cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (22,    'SALES A/C',              'SALESAC',          'SALES A/C',            '',         0,      28,     '~28~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
            cmd.ExecuteNonQuery()
        End If

        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (23,    'CESS A/C',                'CESSAC',          'CESS A/C',             '',         0,      12,     '~12~11~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (24,    'ROUNDOFF A/C',           'ROUNDOFFAC',       'ROUNDOFF A/C',         '',         0,      16,     '~16~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (25,    'CGST A/C',                'CGSTAC',            'CGST A/C',              '',         0,      12,     '~12~11~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (26,    'SGST A/C',                'SGSTAC',            'SGST A/C',              '',         0,      12,     '~12~11~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (27,    'IGST A/C',                'IGSTAC',            'IGST A/C',              '',         0,      12,     '~12~11~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()

        ' If Trim(UCase(Common_Procedures.settings.CustomerCode)) = "1193" Then  '--- BIKE STAND
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (28,    'SERVICE REVENUE A/C',                'SERVICEREVENUEAC',            'SERVICE REVENUE A/C',              '',         0,      19,     '~19~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()
        '  End If

        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (30,    'RECEIVED IN ADVANCE A/C',                'RECEIVEDINADVANCEAC',            'RECEIVED IN ADVANCE A/C',              '',         0,      19,     '~19~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()

        If Trim(UCase(Common_Procedures.settings.CustomerCode)) = "1167" Then  '--- F Fashion
            If Trim(vBankAcName) = "" Then vBankAcName = "BANK A/C"
            cmd.CommandText = "Insert into Ledger_Head ( Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (51,    '" & Trim(vBankAcName) & "',    '" & Common_Procedures.OrderBy_CodeToValue(Trim(vBankAcName)) & "',     '" & Trim(vBankAcName) & "',        '',     0,      5,     '~5~4~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
            cmd.ExecuteNonQuery()
        End If

        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (52,     'SERVIE REPLACEMENT A/C',        'SERVIEREPLACEMENTAC',   'SERVIE REPLACEMENT A/C',       '',         0,      16,     '~16~18~',  'BALANCE ONLY',     '',       '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Ledger_AlaisHead Where Ledger_IdNo <= 100"
        cmd.ExecuteNonQuery()

        Da1 = New SqlClient.SqlDataAdapter("select * from Ledger_Head where Ledger_IdNo <= 100", Cn1)
        Dt1 = New DataTable
        Da1.Fill(Dt1)

        If Dt1.Rows.Count > 0 Then

            For i = 0 To Dt1.Rows.Count - 1
                cmd.CommandText = "Insert into Ledger_AlaisHead(Ledger_IdNo, Sl_No, Ledger_DisplayName, Ledger_Type, AccountsGroup_IdNo ) Values (" & Str(Val(Dt1.Rows(i).Item("Ledger_IdNo").ToString)) & ",    1,      '" & Trim(Dt1.Rows(i).Item("Ledger_Name").ToString) & "',   '" & Trim(Dt1.Rows(i).Item("Ledger_Type").ToString) & "',    " & Str(Val(Dt1.Rows(i).Item("AccountsGroup_IdNo").ToString)) & ")"
                cmd.ExecuteNonQuery()
            Next

        End If

        cmd.Dispose()
        Dt1.Dispose()
        Da1.Dispose()

    End Sub

    Public Shared Sub Default_MonthHead_Updation(ByVal Cn1 As SqlClient.SqlConnection)
        Dim cmd As New SqlClient.SqlCommand

        cmd.Connection = Cn1

        cmd.CommandText = "delete from Month_Head"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (0,     '',                '',         0)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (4,     'APRIL',           'APR',      1)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (5,     'MAY',             'MAY',      2)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (6,     'JUNE',            'JUN',      3)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (7,     'JULY',            'JUL',      4)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (8,     'AUGUST',          'AUG',      5)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (9,     'SEPTEMBER',       'SEP',      6)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (10,     'OCTOBER',        'OCT',      7)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (11,     'NOVEMBER',       'NOV',      8)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (12,     'DECEMBER',       'DEC',      9)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (1,     'JANUARY',         'JAN',      10)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (2,     'FEBRUARY',        'FEB',      11)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Month_Head(Month_IdNo, Month_Name, Month_ShortName, Idno ) Values (3,     'MARCH',           'MAR',      12)"
        cmd.ExecuteNonQuery()

    End Sub

    Public Shared Sub Default_StateHead_Updation(ByVal Cn1 As SqlClient.SqlConnection)
        Dim cmd As New SqlClient.SqlCommand

        cmd.Connection = Cn1

        cmd.CommandText = "delete from State_Head"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (0,     '',                '',         0, '')"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (1,     'TAMIL NADU'         ,    'TAMILNADU'      ,      0,  '33')"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (2,     'ANDHRA PRADESH'     ,    'ANDHRAPRADESH'  ,      1,  28)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (3,     'ARUNACHAL PRADESH'  ,    'ARUNACHALPRADESH',      1,  12)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (4,     'ASSAM'              ,    'ASSAM'           ,      1,  18)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (5,     'BIHAR'              ,    'BIHAR'           ,      1,  10)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (6,     'CHHATTISGARH '      ,    'CHHATTISGARH'    ,      1,  04)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (7,     'GOA'                ,    'GOA'             ,      1,  30)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (8,     'GUJARAT'            ,    'GUJARAT'         ,      1,  24)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (9,     'HARYANA '           ,    'HARYANA'         ,      1,  06)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (10,     'HIMACHAL PRADESH'  ,    'HIMACHALPRADESH',      1,  02)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (11,     'JAMMU AND KASHMIR' ,    'JAMMUANDKASHMIR' ,      1, 01)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (12,     'JHARKHAND'         ,    'JHARKHAND'       ,      1, 20)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (13,     'KARNATAKA'         ,    'KARNATAKA'       ,     1, 29)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (14,     'KERALA'            ,    'KERALA'          ,     1, 32)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (15,     'MADHYA PRADESH'    ,    'MADHYAPRADESH'   ,      1, 23)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (16,     'MAHARASHTRA'       ,    'MAHARASHTRA'     ,      1, 27)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (17,     'MANIPUR'           ,    'MANIPUR'         ,      1, 14)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (18,     'MEGHALAYA'         ,    'MEGHALAYA'       ,      1, 17)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (19,     'MIZORAM'           ,    'MIZORAM'         ,     1, 15)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (20,     'NAGALAND'          ,    'NAGALAND'        ,      1, 13)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (21,     'ODISHA'            ,    'ODISHA'          ,      1, 21)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (22,     'PUNJAB'            ,    'PUNJAB'          ,      1, 03)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (23,     'RAJASTHAN'         ,    'RAJASTHAN'       ,      1, 08)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (24,     'SIKKIM'            ,    'SIKKIM'          ,      1, 11)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (25,     'TELANGANA'         ,    'TELANGANA'       ,      1, 36)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (26,     'TRIPURA'           ,    'TRIPURA'         ,      1, 16)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (27,     'UTTAR PRADESH'     ,    'UTTARPRADESH'   ,      1, 09)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (28,     'UTTARAKHAND'       ,    'UTTARAKHAND'     ,      1, 05)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (29,     'WEST BENGAL'       ,    'WESTBENGAL'     ,      1, 19)"
        cmd.ExecuteNonQuery()


        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (30,     'PUDUCHERRY'         ,    'PUDUCHERRY'       ,      1, 34)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (31,     'NEW DELHI'        ,    'NEWDELHI'         ,      1,  07)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (32,     'LAKSHADWEEPH'     ,    'LAKSHADWEEPH'   ,      1, 31)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (33,     'ANDAMAN AND NICOBAR ISLANDS'       ,   'ANDAMANANDNICOBARISLANDS'    ,      1, 35)"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code  ) Values (34,     'CHANDIGARH'                   ,   'CHANDIGARH'                  ,      1,  04)"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code ) Values (35,     'DADRA AND NAGAR HAVELI'       ,    'DADRAANDNAGARHAVELI'        ,     1,  26)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into State_Head(State_IdNo, State_Name, Sur_Name, Cst_Value, State_Code ) Values (36,    'DAMAN AND DIU'                ,    'DAMANANDDIU'                ,  1,    25)"
        cmd.ExecuteNonQuery()


    End Sub

    Public Shared Sub Default_Shift_Updation(ByVal Cn1 As SqlClient.SqlConnection)
        Dim cmd As New SqlClient.SqlCommand

        cmd.Connection = Cn1

        cmd.CommandText = "delete from Shift_Head"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "Insert into Shift_Head(Shift_IdNo, Shift_Name) Values (0, '')"
        cmd.ExecuteNonQuery()

        If Trim(UCase(Common_Procedures.settings.CustomerCode)) = "1016" Then
            cmd.CommandText = "Insert into Shift_Head(Shift_IdNo, Shift_Name) Values (1, 'DAY')"
            cmd.ExecuteNonQuery()

            cmd.CommandText = "Insert into Shift_Head(Shift_IdNo, Shift_Name) Values (2, 'NIGHT')"
            cmd.ExecuteNonQuery()

        Else

            cmd.CommandText = "Insert into Shift_Head(Shift_IdNo, Shift_Name) Values (1, '1ST SHIFT')"
            cmd.ExecuteNonQuery()

            cmd.CommandText = "Insert into Shift_Head(Shift_IdNo, Shift_Name) Values (2, '2ND SHIFT')"
            cmd.ExecuteNonQuery()

            cmd.CommandText = "Insert into Shift_Head(Shift_IdNo, Shift_Name) Values (3, '3RD SHIFT')"
            cmd.ExecuteNonQuery()

        End If



    End Sub


    Public Shared Sub Default_Master_Updation(ByVal Cn1 As SqlClient.SqlConnection)
        Dim Cn2 As New SqlClient.SqlConnection
        Dim cmd As New SqlClient.SqlCommand
        Dim Dat As Date

        On Error Resume Next


        If Trim(Common_Procedures.ConnectionString_CompanyGroupdetails) <> "" Then

            Cn2 = New SqlClient.SqlConnection(Common_Procedures.ConnectionString_CompanyGroupdetails)

            Cn2.Open()

            cmd.Connection = Cn2

            Dat = #1/1/1900#
            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@EntryDate", Dat.ToShortDateString)

            cmd.CommandText = "Delete from CompanyGroup_Head where CompanyGroup_IdNo = 0"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "Insert into CompanyGroup_Head(CompanyGroup_IdNo, CompanyGroup_Name, From_Date, To_Date, Financial_Range) values (0, '', @EntryDate, @EntryDate, '')"
            cmd.ExecuteNonQuery()

            cmd.CommandText = "Delete from User_Head where user_idno = 0"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "Insert into User_Head(User_IdNo, User_Name, Sur_Name, Account_Password, UnAccount_Password) values (0, '', '', '', '') "
            cmd.ExecuteNonQuery()

            Cn2.Close()
            Cn2.Dispose()

        End If

        cmd.Connection = Cn1

        Dat = #1/1/2000#
        cmd.Parameters.Clear()
        cmd.Parameters.AddWithValue("@EntryDate", Dat.ToShortDateString)


        cmd.CommandText = "delete from Ledger_Head where Ledger_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_Head(Ledger_IdNo, Ledger_Name, Sur_Name, Ledger_MainName, Ledger_AlaisName, Area_IdNo, AccountsGroup_IdNo, Parent_Code, Bill_Type, Ledger_Type, Ledger_Address1, Ledger_Address2, Ledger_Address3, Ledger_Address4, Ledger_TinNo, Ledger_CstNo ) Values (0,   '',   '',   '',     '',     0,      0,      '',     '',     '',     '',     '',     '',     '',     '',     ''  )"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Ledger_AlaisHead where Ledger_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_AlaisHead(Ledger_IdNo, Sl_No, Ledger_DisplayName) Values (0,      0,     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Ledger_PhoneNo_Head where Ledger_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Ledger_PhoneNo_Head(Ledger_IdNo, Sl_No, Ledger_PhoneNo) Values (0,      0,     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from AccountsGroup_Head where AccountsGroup_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into AccountsGroup_Head(AccountsGroup_IdNo, AccountsGroup_Name, Sur_Name, Parent_Name, Indicate, Parent_Idno, Carried_Balance, Order_Position, TallyName, TallySubName ) Values (0, '',          '',       '',    0, '',       0,  0,      '',                         ''      )"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Company_Head where Company_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Company_Head(Company_IdNo, Company_Name, Company_SurName, Company_ShortName, Company_Address1, Company_Address2, Company_Address3, Company_Address4, Company_City, Company_PinCode, Company_PhoneNo, Company_TinNo, Company_CstNo, Company_FaxNo, Company_EMail, Company_ContactPerson, Company_Description) Values (0,      '',     '',     '',     '',     '',     '',     '',     '',     '',     '',     '',     '',     '',     '',     '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Item_Head where Item_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Item_Head(Item_IdNo, Item_Name, Sur_Name, Item_Code, ItemGroup_IdNo, Unit_IdNo, Tax_Percentage, Sale_TaxRate, Sales_Rate, Cost_Rate, Minimum_Stock) Values (0,   '',     '',     '',     0,     0,     0,     0,     0,     0,     0)"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from ItemGroup_Head where ItemGroup_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into ItemGroup_Head(ItemGroup_IdNo, ItemGroup_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Unit_Head where Unit_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Unit_Head(Unit_IdNo, Unit_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Cetegory_Head where Cetegory_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Cetegory_Head(Cetegory_IdNo, Cetegory_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()


        cmd.CommandText = "delete from Variety_Head where Variety_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Variety_Head(Variety_IdNo, Variety_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Area_Head where Area_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Area_Head(Area_IdNo, Area_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Waste_Head where Waste_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Waste_Head(Waste_IdNo, Waste_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Size_Head where Size_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Size_Head(Size_IdNo, Size_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Transport_Head where Transport_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Transport_Head(Transport_IdNo, Transport_Name, Sur_Name) Values (0,   '',     '')"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Shift_Head where Shift_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Shift_Head ( Shift_IdNo , Shift_Name ) Values (0,   '' )"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Price_List_Head where Price_List_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Price_List_Head ( Price_List_IdNo , Price_List_Name, sur_name) Values (0,   '', '' )"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Machine_Head where Machine_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Machine_Head ( Machine_IdNo , Machine_Name, sur_name) Values (0,   '', '' )"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Tax_Head where Tax_IdNo = 0"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Tax_Head(Tax_IdNo, Tax_Name, Sur_Name, Tax_Ledger_Ac_IdNo) Values (0,   '',     '', 0)"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "delete from Sales_Head where Sales_Code = ''"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Insert into Sales_Head ( Sales_Code, Company_IdNo, Sales_No, for_OrderBy, Sales_Date) Values ('-99999/00-00',   0,     '',     0,   @EntryDate  )"
        cmd.ExecuteNonQuery()

    End Sub

    Public Shared Sub Sql_AutoBackUP(ByVal Db_Name As String)
        Dim cn1 As SqlClient.SqlConnection
        Dim cmd As New SqlClient.SqlCommand
        Dim Fl_Name As String, Fl_Name2 As String
        Dim Fl_Name3 As String
        Dim Fl_Name_OneDrive As String
        Dim ServrNm As String = ""
        Dim ServrPath As String = ""

        cn1 = New SqlClient.SqlConnection(Common_Procedures.ConnectionString_CompanyGroupdetails)
        cn1.Open()

        ServrNm = Common_Procedures.get_Server_SystemName()
        If ServrNm = Trim(UCase(SystemInformation.ComputerName)) Then

            Fl_Name = Common_Procedures.AppPath & "\Auto_BackUP"
            Fl_Name3 = Common_Procedures.AppPath & "\Auto_BackUP"
            ' Fl_Name_OneDrive = Common_Procedures.AppPath & "\OneDrive"
        Else
            ServrPath = Trim(Common_Procedures.get_FieldValue(cn1, "Settings_Head", "Autobackup_Path_Server", ""))

            If ServrPath = "" Then Exit Sub

            Fl_Name = Trim(ServrPath) & "\Auto_BackUP"
            Fl_Name3 = Trim(ServrPath) & "\Auto_BackUP"
            ' Fl_Name_OneDrive = Trim(ServrPath) & "\OneDrive"
        End If


        If System.IO.Directory.Exists(Fl_Name) = False Then
            System.IO.Directory.CreateDirectory(Fl_Name)
        End If

        Fl_Name3 = Trim(Fl_Name) & "\" & Trim(Db_Name) & "_" & Trim(Format(Date.Today, "ddMMMyy")) & ".plit"
        Fl_Name = Trim(Fl_Name) & "\" & Trim(Db_Name) & "_" & Trim(Format(Date.Today, "ddMMMyy"))  ' & ".bak"


        cmd.Connection = cn1

        If Common_Procedures.OneDrive_Status = True Then
            'onedrive backup

            Fl_Name_OneDrive = Common_Procedures.OneDrive_Path

            If System.IO.Directory.Exists(Fl_Name_OneDrive) = False Then
                System.IO.Directory.CreateDirectory(Fl_Name_OneDrive)
            End If

            Fl_Name_OneDrive = Trim(Fl_Name_OneDrive) & "\" & Trim(Db_Name) & "_" & Trim(Format(Date.Today, "ddMMMyy")) & ".plit"

            cmd.CommandText = "BACKUP DATABASE " & Trim(Db_Name) & " TO DISK = '" & Trim(Fl_Name_OneDrive) & "' WITH INIT"
            cmd.ExecuteNonQuery()

        End If

        cmd.CommandText = "BACKUP DATABASE " & Trim(Db_Name) & " TO DISK = '" & Trim(Fl_Name) & "' WITH INIT"
        cmd.ExecuteNonQuery()


        cmd.CommandText = "BACKUP DATABASE " & Trim(Db_Name) & " TO DISK = '" & Trim(Fl_Name3) & "' WITH INIT"
        cmd.ExecuteNonQuery()

        cmd.Dispose()

        cn1.Close()
        cn1.Dispose()

        Sql_AutoBackUP_File_To_Client_Sysytem(Fl_Name, Trim(Db_Name), ServrNm)


        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim d As DriveInfo

        For Each d In allDrives

            If d.IsReady = True Then

                If d.DriveType = DriveType.Removable Then

                    Fl_Name2 = Trim(d.Name) & "PLit\Auto_BackUP"

                    If System.IO.Directory.Exists(Fl_Name2) = True Then

                        Fl_Name2 = Trim(Fl_Name2) & "\" & Trim(Db_Name) & "_" & Trim(Format(Date.Today, "ddMMMyy")) & ".plit"

                        System.IO.File.Copy(Fl_Name, Fl_Name2, True)

                        Exit For

                    End If

                End If

            End If

        Next

        If Trim(UCase(Common_Procedures.settings.CustomerCode)) = "1130" Then
            For Each d In allDrives

                If d.IsReady = True Then

                    If d.DriveType = DriveType.Fixed Then

                        Fl_Name2 = "D:\PLit\Auto_BackUP"

                        If System.IO.Directory.Exists(Fl_Name2) = True Then

                            Fl_Name2 = Trim(Fl_Name2) & "\" & Trim(Db_Name) & "_" & Trim(Format(Date.Today, "ddMMMyy")) & ".plit"

                            System.IO.File.Copy(Fl_Name, Fl_Name2, True)

                            Exit For

                        End If

                    End If

                End If

            Next
        End If





    End Sub

    Public Shared Sub Sql_AutoBackUP_File_To_Client_Sysytem(ByVal File_Name As String, ByVal Db_Name As String, ByVal Servnam As String)
        Dim cn1 As SqlClient.SqlConnection
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim Path1 As String = ""
        Dim Path2 As String = ""

        Try
            Common_Procedures.Sql_AutoBackUP_Client_Path()

            cn1 = New SqlClient.SqlConnection(Common_Procedures.ConnectionString_CompanyGroupdetails)
            cn1.Open()

            Da = New SqlClient.SqlDataAdapter("select * from AutoBackup_Path_Head  order by Auto_SlNo asc ", cn1)
            Da.Fill(Dt)
            If Dt.Rows.Count > 0 Then
                For i = 0 To Dt.Rows.Count - 1
                    If IsDBNull(Dt.Rows(0).Item("App_Path").ToString) = False Then
                        If Trim(Dt.Rows(0).Item("App_Path").ToString) <> "" Then

                            If Directory.Exists(Trim(Dt.Rows(0).Item("App_Path").ToString)) Then
                                'System.IO.File.Copy(File_Name, Trim(Dt.Rows(0).Item("App_Path").ToString) & "\" & Trim(Db_Name) & "_" & Trim(Format(Date.Today, "ddMMMyy")))
                                System.IO.File.Copy(File_Name, Trim(Dt.Rows(0).Item("App_Path").ToString) & "\" & Trim(Db_Name) & "_" & Trim(Format(Date.Today, "ddMMMyy")) & ".tssql")

                                'Directory.CreateDirectory(Path1)
                            End If

                        End If

                    End If
                Next

            End If

            Dt.Dispose()
            Da.Dispose()

            cn1.Close()
            cn1.Dispose()


        Catch ex As Exception
            '----
        End Try


    End Sub
    Public Shared Sub Sql_AutoBackUP_Client_Path()
        Dim cmd As New SqlClient.SqlCommand
        Dim cn1 As SqlClient.SqlConnection
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim ServrNm As String = ""
        Dim Path_Sts As Boolean = False
        Dim Client_Count As Integer = 0
        Dim Nr As Integer = 0
        Dim Path1 As String = ""
        Try
            ServrNm = Common_Procedures.get_Server_SystemName()
            If ServrNm = Trim(UCase(SystemInformation.ComputerName)) Then
                Exit Sub
            End If

            cn1 = New SqlClient.SqlConnection(Common_Procedures.ConnectionString_CompanyGroupdetails)
            cn1.Open()
            cmd.Connection = cn1

            Path1 = Replace(Trim(Common_Procedures.AppPath & "\Auto_BackUP"), ":", "")
            Path1 = "\\" & Trim(UCase(SystemInformation.ComputerName)) & "\" & Trim(Path1)

            If Not Directory.Exists(Path1) Then
                Directory.CreateDirectory(Path1)
            End If

            Da = New SqlClient.SqlDataAdapter("select top 1 * from AutoBackup_Path_Head where Computer_Name = '" & Trim(Trim(UCase(SystemInformation.ComputerName))) & "' order by Auto_SlNo asc ", cn1)
            Da.Fill(Dt)
            If Dt.Rows.Count > 0 Then

                Nr = 0
                cmd.CommandText = "update AutoBackup_Path_Head set App_Path = '" & Trim(Path1) & "' where Computer_Name = '" & Trim(Trim(UCase(SystemInformation.ComputerName))) & "'"
                Nr = cmd.ExecuteNonQuery()

            Else

                cmd.CommandText = "Insert into AutoBackup_Path_Head(Computer_Name,App_Path) Values ('" & Trim(Trim(UCase(SystemInformation.ComputerName))) & "' ,'" & Trim(Path1) & "' )"
                cmd.ExecuteNonQuery()

            End If
            Dt.Dispose()
            Da.Dispose()


            cn1.Close()
            cn1.Dispose()

        Catch ex As Exception
            '----
        End Try


    End Sub
   


    

    Public Shared Function UserRight_Check(ByVal User_Access_Type As String, ByVal NewEntry_Status As Boolean) As Boolean

        UserRight_Check = True

        If Val(Common_Procedures.User.IdNo) <> 1 Then
            If InStr(Trim(UCase(User_Access_Type)), "~L~") = 0 Then
                If NewEntry_Status = True Then
                    If InStr(Trim(UCase(User_Access_Type)), "~A~") = 0 Then
                        MessageBox.Show("You have No Rights to Add", "DOES NOT SAVE...", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        UserRight_Check = False
                    End If

                Else
                    If InStr(Trim(UCase(User_Access_Type)), "~E~") = 0 Then
                        MessageBox.Show("You have No Rights to Change", "DOES NOT SAVE...", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        UserRight_Check = False
                    End If

                End If
            End If
        End If

    End Function
    Public Shared Function UserRight_Check_FOR_EDIT_SOME_FIELD(ByVal User_Access_Type As String) As Boolean

        UserRight_Check_FOR_EDIT_SOME_FIELD = True

        If Val(Common_Procedures.User.IdNo) <> 1 Then

            If InStr(Trim(UCase(User_Access_Type)), "~L~") = 0 Then
                UserRight_Check_FOR_EDIT_SOME_FIELD = False
            End If
            If InStr(Trim(UCase(User_Access_Type)), "~A~") = 0 Then
                UserRight_Check_FOR_EDIT_SOME_FIELD = False
            End If
            If InStr(Trim(UCase(User_Access_Type)), "~E~") = 0 Then
                UserRight_Check_FOR_EDIT_SOME_FIELD = False
            End If


        End If

    End Function

    Public Shared Sub ComboBox_ItemSelection_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal Cn1 As SqlClient.SqlConnection, ByVal CboName As ComboBox, ByVal NextCtrlName As Object, ByVal vTableName As String, ByVal vSelectionFieldName As String, ByVal vSqlCondition As String, ByVal vBlankFieldCondition As String, Optional ByVal vBlock_Typing_Status As Boolean = True, Optional ByVal UPPERCASE As Boolean = True)
        Dim da As New SqlClient.SqlDataAdapter
        Dim Cmd As New SqlClient.SqlCommand
        Dim dt As New DataTable
        Dim SqlCondt As String, Condt2 As String
        Dim FindStr As String = ""
        Dim indx As Integer = -1
        Dim SelStrt As Integer = 0
        Dim Mtch_STS As Boolean = False


        Try

            With CboName

                If Asc(e.KeyChar) <> 27 Then

                    SelStrt = .SelectionStart

                    If Asc(e.KeyChar) = 13 Then

                        Try

                            If Trim(.Text) <> "" Then

                                If .DroppedDown = True Then

                                    If .Items.Count > 0 Then

                                        indx = .FindString(FindStr)

                                        If indx <> -1 Then

                                            If .SelectedIndex >= 0 Then
                                                .SelectedItem = .Items(.SelectedIndex)
                                                If UPPERCASE = True Then
                                                    .Text = UCase(.GetItemText(.SelectedItem))
                                                Else
                                                    .Text = .GetItemText(.SelectedItem)
                                                End If

                                            Else

                                                If Trim(vTableName) <> "" And Trim(vSelectionFieldName) <> "" Then
                                                    .SelectedIndex = 0
                                                    .SelectedItem = .Items(0)
                                                    If UPPERCASE = True Then
                                                        .Text = UCase(.GetItemText(.SelectedItem))
                                                    Else
                                                        .Text = .GetItemText(.SelectedItem)
                                                    End If

                                                End If

                                            End If

                                        End If

                                    End If

                                End If

                            End If

                        Catch ex As Exception
                            '---

                        End Try


                        If IsNothing(NextCtrlName) = False Then
                            If NextCtrlName.Enabled Then
                                NextCtrlName.Focus()

                            Else
                                SendKeys.Send("{TAB}")

                            End If
                        End If

                    Else

                        SqlCondt = ""
                        Condt2 = ""
                        FindStr = ""
                        indx = -1

                        If Asc(e.KeyChar) = 8 Then

                            If Trim(.Text) <> "" Then

                                If .SelectionLength = 0 Then
                                    If .SelectionStart > 1 Then
                                        FindStr = .Text.Substring(0, .SelectionStart - 1)
                                    End If
                                    FindStr = FindStr & Mid(CboName.Text, CboName.SelectionStart + 1, Len(CboName.Text))

                                Else

                                    If .SelectionStart <= 1 Then
                                        .Text = ""
                                    Else
                                        FindStr = .Text.Substring(0, .SelectionStart - 1)
                                    End If

                                End If

                            End If

                        Else

                            If .SelectionLength = 0 Then
                                If .SelectionStart > 0 Then FindStr = .Text.Substring(0, .SelectionStart)
                                FindStr = FindStr & e.KeyChar & Mid(CboName.Text, CboName.SelectionStart + 1, Len(CboName.Text))

                            Else
                                FindStr = .Text.Substring(0, .SelectionStart) & e.KeyChar

                            End If

                        End If

                        FindStr = LTrim(FindStr)
                        If Trim(vTableName) <> "" Then

                            indx = .FindString(FindStr)

                            SqlCondt = ""
                            If Trim(FindStr) <> "" Then
                                SqlCondt = " Where " & vSqlCondition & IIf(Trim(vSqlCondition) <> "", " and ", "") & " (" & vSelectionFieldName & " like '" & FindStr & "%' or " & vSelectionFieldName & " like '% " & FindStr & "%' or " & vSelectionFieldName & " like '% (" & FindStr & "%' or " & vSelectionFieldName & " like '(" & FindStr & "%' or " & vSelectionFieldName & " like '% {" & FindStr & "%' or " & vSelectionFieldName & " like '{" & FindStr & "%'   or " & vSelectionFieldName & " like '% [" & FindStr & "%' or " & vSelectionFieldName & " like '[" & FindStr & "%')"

                            Else

                                Condt2 = ""
                                If Trim(vSqlCondition) <> "" Then
                                    Condt2 = Trim(vSqlCondition)
                                    If Trim(vBlankFieldCondition) <> "" Then Condt2 = Condt2 & IIf(Trim(Condt2) <> "", " or ", "") & vBlankFieldCondition
                                End If

                                If Trim(Condt2) <> "" Then
                                    SqlCondt = " Where " & Trim(Condt2)
                                End If

                            End If

                            Mtch_STS = False
                            da = New SqlClient.SqlDataAdapter("select distinct(" & vSelectionFieldName & ") from " & vTableName & " " & SqlCondt & " order by " & vSelectionFieldName, Cn1)
                            dt = New DataTable
                            da.Fill(dt)
                            If dt.Rows.Count > 0 Then
                                Mtch_STS = True
                            End If

                            If Mtch_STS = True Then

                                da = New SqlClient.SqlDataAdapter("Select distinct(" & vSelectionFieldName & ") from " & vTableName & " " & SqlCondt & " order by " & vSelectionFieldName, Cn1)
                                dt = New DataTable
                                da.Fill(dt)
                                .DataSource = dt
                                .DisplayMember = Trim(vSelectionFieldName)

                                If .Items.Count > 0 Then
                                    If Asc(e.KeyChar) = 32 And Len(FindStr) = 0 Then .DroppedDown = False
                                    .DroppedDown = True
                                End If

                                If UPPERCASE = True Then
                                    .Text = UCase(FindStr)
                                Else
                                    .Text = FindStr
                                End If

                                If Asc(e.KeyChar) = 8 Then
                                    If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                Else
                                    .SelectionStart = SelStrt + 1
                                End If

                            Else

                                If vBlock_Typing_Status = True Then
                                    If Trim(FindStr) <> "" Then
                                        If UPPERCASE = True Then
                                            .Text = UCase(Microsoft.VisualBasic.Left(FindStr, Len(FindStr) - 1))
                                        Else
                                            .Text = Microsoft.VisualBasic.Left(FindStr, Len(FindStr) - 1)
                                        End If

                                        .SelectionStart = .Text.Length
                                    End If
                                Else
                                    .DataSource = Nothing
                                    .DisplayMember = ""

                                    If UPPERCASE = True Then
                                        .Text = UCase(FindStr)
                                    Else
                                        .Text = FindStr
                                    End If
                                    If Asc(e.KeyChar) = 8 Then
                                        If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                    Else
                                        .SelectionStart = SelStrt + 1
                                    End If
                                End If

                            End If

                            e.Handled = True

                            If Mtch_STS = False And vBlock_Typing_Status = False Then
                                If .DroppedDown = True Then

                                    Cmd.Connection = Cn1

                                    Cmd.CommandText = "truncate table Combo_Temp"
                                    Cmd.ExecuteNonQuery()

                                    Cmd.CommandText = "insert into Combo_Temp(name1) values ('" & Trim(UCase(FindStr)) & "')"
                                    Cmd.ExecuteNonQuery()

                                    da = New SqlClient.SqlDataAdapter("Select distinct(Name1) from Combo_Temp", Cn1)
                                    dt = New DataTable
                                    da.Fill(dt)
                                    .DataSource = dt
                                    .DisplayMember = "Name1"

                                    indx = .FindString(FindStr)

                                    If indx <> -1 Then
                                        If .SelectedIndex >= 0 Then
                                            .SelectedIndex = 0
                                            .SelectedItem = .Items(.SelectedIndex)
                                            If UPPERCASE = True Then
                                                .Text = UCase(.GetItemText(.SelectedItem))
                                            Else
                                                .Text = .GetItemText(.SelectedItem)
                                            End If
                                        End If
                                    End If


                                    Try
                                        .DroppedDown = False
                                    Catch ex As Exception
                                        '----
                                    End Try

                                    '.DataSource = Nothing
                                    '.DisplayMember = ""

                                    Try
                                        If UPPERCASE = True Then
                                            .Text = UCase(FindStr)
                                        Else
                                            .Text = FindStr
                                        End If

                                    Catch ex As Exception
                                        ''---
                                        ''Try
                                        ''    .Text = FindStr
                                        ''Catch ex1 As Exception
                                        ''    ----
                                        ''End Try

                                    End Try

                                    If Asc(e.KeyChar) = 8 Then
                                        If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                    Else
                                        .SelectionStart = SelStrt + 1 'FindStr.Length
                                    End If

                                    .SelectionLength = .Text.Length

                                    '.SelectedIndex = -1

                                End If
                            End If

                        Else

                            indx = .FindString(FindStr)
                            If indx <> -1 Then

                                If .Items.Count > 0 Then
                                    If Asc(e.KeyChar) = 32 And Len(FindStr) = 0 Then .DroppedDown = False
                                    .DroppedDown = True
                                End If

                                '.SelectedText = ""
                                .SelectedIndex = indx
                                .SelectedItem = .Items(.SelectedIndex)
                                'If UPPERCASE = True Then
                                '    .Text = UCase(.GetItemText(.SelectedItem))
                                'Else
                                .Text = .GetItemText(.SelectedItem)
                                'End If

                                If Asc(e.KeyChar) = 8 Then
                                    If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                Else
                                    .SelectionStart = SelStrt + 1 'FindStr.Length
                                End If

                                .SelectionLength = .Text.Length
                                e.Handled = True

                            Else

                                If vBlock_Typing_Status = True Then
                                    If Trim(FindStr) <> "" Then
                                        'If UPPERCASE = True Then
                                        '    .Text = UCase(Microsoft.VisualBasic.Left(FindStr, Len(FindStr) - 1))
                                        'Else
                                        .Text = Microsoft.VisualBasic.Left(FindStr, Len(FindStr) - 1)
                                        'End If
                                        .SelectionStart = .Text.Length
                                    End If

                                Else
                                    .DataSource = Nothing
                                    .DisplayMember = ""
                                    .SelectedText = ""
                                    .SelectedIndex = -1

                                    'If UPPERCASE = True Then
                                    '    .Text = UCase(FindStr)
                                    'Else
                                    .Text = FindStr
                                    'End If

                                    If Asc(e.KeyChar) = 8 Then
                                        If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                    Else
                                        .SelectionStart = SelStrt + 1 'FindStr.Length
                                    End If
                                End If

                                e.Handled = True
                            End If

                        End If

                    End If

                End If

            End With


        Catch ex As NullReferenceException
            'MessageBox.Show(ex.Message, "ERROR WHILE KEYPRESS IN COMBOBOX " & sender.ToString & "....", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)

        Catch ex As ObjectDisposedException
            'MessageBox.Show(ex.Message, "ERROR WHILE KEYPRESS IN COMBOBOX " & sender.ToString & "....", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)

        Catch ex As ArgumentException
            'MessageBox.Show(ex.Message, "ERROR WHILE KEYPRESS IN COMBOBOX " & sender.ToString & "....", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "ERROR IN WHILE KEYPRESS IN COMBOBOX " & sender.ToString & "....", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)

        Finally
            '---

        End Try

    End Sub


    Public Shared Sub ComboBox_ItemSelection_KeyPress_111(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal Cn1 As SqlClient.SqlConnection, ByVal CboName As ComboBox, ByVal NextCtrlName As Object, ByVal vTableName As String, ByVal vSelectionFieldName As String, ByVal vSqlCondition As String, ByVal vBlankFieldCondition As String, Optional ByVal vBlock_Typing_Status As Boolean = True)
        Dim da As New SqlClient.SqlDataAdapter
        Dim dt As New DataTable
        Dim SqlCondt As String, Condt2 As String
        Dim FindStr As String
        Dim indx As Integer = -1
        Dim SelStrt As Integer = 0

        Try

            With CboName

                If Asc(e.KeyChar) <> 27 Then

                    SelStrt = .SelectionStart

                    If Asc(e.KeyChar) = 13 Then

                        If Trim(.Text) <> "" Then

                            If .DroppedDown = True Then

                                If .Items.Count > 0 Then

                                    If .SelectedIndex >= 0 Then
                                        .SelectedItem = .Items(.SelectedIndex)
                                        .Text = .GetItemText(.SelectedItem)

                                    Else
                                        If Trim(vTableName) <> "" And Trim(vSelectionFieldName) <> "" Then
                                            .SelectedIndex = 0
                                            .SelectedItem = .Items(0)
                                            .Text = .GetItemText(.SelectedItem)

                                        End If

                                    End If

                                End If

                            End If

                        End If

                        If IsNothing(NextCtrlName) = False Then
                            If NextCtrlName.Enabled Then
                                NextCtrlName.Focus()

                            Else
                                SendKeys.Send("{TAB}")

                            End If
                        End If

                    Else

                        SqlCondt = ""
                        Condt2 = ""
                        FindStr = ""
                        indx = -1

                        If Asc(e.KeyChar) = 8 Then

                            If Trim(.Text) <> "" Then

                                If .SelectionLength = 0 Then
                                    If .SelectionStart > 1 Then
                                        FindStr = .Text.Substring(0, .SelectionStart - 1)
                                    End If
                                    FindStr = FindStr & Mid(CboName.Text, CboName.SelectionStart + 1, Len(CboName.Text))

                                Else


                                    If .SelectionStart <= 1 Then
                                        .Text = ""
                                    Else
                                        FindStr = .Text.Substring(0, .SelectionStart - 1)
                                    End If

                                End If

                            End If

                        Else

                            If .SelectionLength = 0 Then

                                If .SelectionStart > 0 Then FindStr = .Text.Substring(0, .SelectionStart)

                                FindStr = FindStr & e.KeyChar & Mid(CboName.Text, CboName.SelectionStart + 1, Len(CboName.Text))

                            Else
                                FindStr = .Text.Substring(0, .SelectionStart) & e.KeyChar

                            End If

                        End If

                        FindStr = LTrim(FindStr)

                        If Trim(vTableName) <> "" Then

                            indx = .FindString(FindStr)

                            If indx <> -1 Or vBlock_Typing_Status = False Then

                                SqlCondt = ""

                                If Trim(FindStr) <> "" Then
                                    If Trim(UCase(Common_Procedures.settings.CustomerCode)) = "1005" Then '---- Jeno Textiles (Somanur)
                                        SqlCondt = " Where " & vSqlCondition & IIf(Trim(vSqlCondition) <> "", " and ", "") & " (" & vSelectionFieldName & " like '" & FindStr & "%') "
                                    Else
                                        SqlCondt = " Where " & vSqlCondition & IIf(Trim(vSqlCondition) <> "", " and ", "") & " (" & vSelectionFieldName & " like '" & FindStr & "%' or " & vSelectionFieldName & " like '% " & FindStr & "%') "
                                    End If

                                Else

                                    Condt2 = ""
                                    If Trim(vSqlCondition) <> "" Then
                                        Condt2 = Trim(vSqlCondition)
                                        If Trim(vBlankFieldCondition) <> "" Then Condt2 = Condt2 & IIf(Trim(Condt2) <> "", " or ", "") & vBlankFieldCondition
                                    End If

                                    If Trim(Condt2) <> "" Then
                                        SqlCondt = " Where " & Trim(Condt2)
                                    End If

                                End If

                                da = New SqlClient.SqlDataAdapter("select distinct(" & vSelectionFieldName & ") from " & vTableName & " " & SqlCondt & " order by " & vSelectionFieldName, Cn1)
                                da.Fill(dt)
                                .DataSource = dt
                                .DisplayMember = Trim(vSelectionFieldName)

                                .Text = FindStr

                                If Asc(e.KeyChar) = 8 Then
                                    If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                Else
                                    .SelectionStart = SelStrt + 1
                                End If

                            End If

                            e.Handled = True

                        Else

                            indx = .FindString(FindStr)
                            If indx <> -1 Then
                                .SelectedText = ""
                                .SelectedIndex = indx

                                If Asc(e.KeyChar) = 8 Then
                                    If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                Else
                                    .SelectionStart = SelStrt + 1 'FindStr.Length
                                End If
                                .SelectionLength = .Text.Length
                                e.Handled = True

                            Else

                                .SelectedText = ""
                                .SelectedIndex = -1

                                .Text = FindStr

                                If Asc(e.KeyChar) = 8 Then
                                    If SelStrt > 0 Then .SelectionStart = SelStrt - 1
                                Else
                                    .SelectionStart = SelStrt + 1 'FindStr.Length
                                End If

                                e.Handled = True

                            End If

                        End If

                    End If

                End If

            End With

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "ERROR IN WHILE KEYPRESS IN COMBOBOX " & sender.ToString & "....", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub


    Public Shared Sub ComboBox_ItemSelection_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs, ByVal Cn1 As SqlClient.SqlConnection, ByVal CboName As ComboBox, ByVal PreviousCtrlName As Object, ByVal NextCtrlName As Object, ByVal vTableName As String, ByVal vSelectionFieldName As String, ByVal vSqlCondition As String, ByVal vBlankFieldCondition As String)
        Dim da As New SqlClient.SqlDataAdapter
        Dim dt As New DataTable
        Dim SqlCondt As String, Condt2 As String
        Dim FindStr As String
        Dim indx As Integer
        Dim SelStrt As Integer

        Try

            With CboName

                If (e.KeyValue = 38 And .DroppedDown = False) Or (e.Control = True And e.KeyValue = 38) Then
                    e.Handled = True
                    If IsNothing(PreviousCtrlName) = False Then
                        PreviousCtrlName.Focus()
                    End If

                ElseIf (e.KeyValue = 40 And .DroppedDown = False) Or (e.Control = True And e.KeyValue = 40) Then
                    e.Handled = True
                    If IsNothing(NextCtrlName) = False Then
                        NextCtrlName.Focus()
                    End If

                ElseIf e.KeyValue = 46 Then

                    SqlCondt = ""
                    Condt2 = ""
                    FindStr = ""
                    indx = -1

                    SelStrt = .SelectionStart

                    If .SelectionStart <= 1 And .SelectionLength > 0 Then
                        .Text = ""
                    End If

                    If Trim(.Text) <> "" Then

                        If .SelectionLength = 0 Then

                            If .SelectionStart > 0 Then
                                FindStr = .Text.Substring(0, .SelectionStart)
                            End If
                            FindStr = FindStr & Mid(CboName.Text, CboName.SelectionStart + 2, Len(CboName.Text))

                        Else

                            FindStr = .Text.Substring(0, .SelectionStart - 1)

                        End If

                        'If .SelectionLength = 0 Then
                        '    FindStr = .Text.Substring(0, .Text.Length - 1)
                        'Else
                        '    FindStr = .Text.Substring(0, .SelectionStart - 1)
                        'End If
                    End If

                    FindStr = LTrim(FindStr)

                    If Trim(vTableName) <> "" Then

                        SqlCondt = ""

                        If Trim(FindStr) <> "" Then
                            SqlCondt = " Where " & vSqlCondition & IIf(Trim(vSqlCondition) <> "", " and ", "") & " (" & vSelectionFieldName & " like '" & FindStr & "%' or " & vSelectionFieldName & " like '% " & FindStr & "%') "

                        Else

                            Condt2 = ""
                            If Trim(vSqlCondition) <> "" Then
                                Condt2 = Trim(vSqlCondition)
                                If Trim(vBlankFieldCondition) <> "" Then Condt2 = Condt2 & IIf(Trim(Condt2) <> "", " or ", "") & vBlankFieldCondition
                            End If

                            If Trim(Condt2) <> "" Then
                                SqlCondt = " Where " & Trim(Condt2)
                            End If

                        End If

                        da = New SqlClient.SqlDataAdapter("select " & vSelectionFieldName & " from " & vTableName & " " & SqlCondt & " order by " & vSelectionFieldName, Cn1)
                        da.Fill(dt)
                        .DataSource = dt
                        .DisplayMember = Trim(vSelectionFieldName)

                        .Text = FindStr

                        .SelectionStart = SelStrt  ' FindStr.Length

                        e.Handled = True

                    Else

                        indx = .FindString(FindStr)

                        If indx <> -1 Then
                            .SelectedText = ""
                            .SelectedIndex = indx

                            .SelectionStart = SelStrt  ' FindStr.Length
                            '.SelectionStart = FindStr.Length
                            .SelectionLength = .Text.Length
                            e.Handled = True

                        Else
                            .Text = FindStr

                            .SelectionStart = SelStrt  ' FindStr.Length

                            e.Handled = True

                        End If

                    End If

                ElseIf e.KeyValue <> 13 And e.KeyValue <> 17 And e.KeyValue <> 27 Then
                    If .DroppedDown = False Then
                        .DroppedDown = True
                    End If

                End If

            End With

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "ERROR IN WHILE KEYDOWN IN COMBOBOX " & sender.ToString & "....", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Shared Sub ComboBox_ItemSelection_SetDataSource(ByVal sender As Object, ByVal Cn1 As SqlClient.SqlConnection, ByVal vTableName As String, ByVal vSelectionFieldName As String, ByVal vSqlCondition As String, ByVal vBlankFieldCondition As String)
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt1 As New DataTable
        Dim vCboTxt As String = ""
        Dim SqlCondt As String = ""

        Try
            With sender
                If Trim(vTableName) <> "" And Trim(vSelectionFieldName) <> "" Then

                    vCboTxt = .Text

                    SqlCondt = ""
                    If Trim(vSqlCondition) <> "" Then
                        SqlCondt = " Where " & Trim(vBlankFieldCondition) & IIf(Trim(vBlankFieldCondition) <> "", " or ", "") & Trim(vSqlCondition)
                    End If

                    Da = New SqlClient.SqlDataAdapter("select distinct(" & Trim(vSelectionFieldName) & ") from " & vTableName & " " & SqlCondt & " order by " & vSelectionFieldName, Cn1)
                    Dt1 = New DataTable
                    Da.Fill(Dt1)
                    .DataSource = Dt1
                    .DisplayMember = Trim(vSelectionFieldName)

                    .Text = Trim(vCboTxt)

                    .BackColor = Color.FromArgb(255, 192, 255)
                    .ForeColor = Color.Black
                    .SelectAll()

                End If

            End With
        Catch ex As Exception
            'MessageBox.Show(ex.Message, "ERROR IN SETTING DATASOURCE " & sender.ToString & "....", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Da.Dispose()

        End Try

    End Sub

    Public Shared Function Check_Negative_Stock_Status(ByVal Cn1 As SqlClient.SqlConnection, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Boolean
        Dim Da1 As New SqlClient.SqlDataAdapter
        Dim Dt1 As New DataTable
        Dim Led_Idno As Integer = 0
        Dim I As Integer
        Dim Stk As Single = 0
        Dim ForStk_Weight As String = ""
        Dim Descp As String = ""
        Dim CurStk As Decimal = 0

        Check_Negative_Stock_Status = False

        Da1 = New SqlClient.SqlDataAdapter("Select Company_Idno, Item_IdNo, sum(Quantity) from TempTable_For_NegativeStock group by Company_Idno, Item_IdNo Order by Company_Idno, Item_IdNo", Cn1)
        If IsNothing(sqltr) = False Then
            Da1.SelectCommand.Transaction = sqltr
        End If
        Dt1 = New DataTable
        Da1.Fill(Dt1)

        If Dt1.Rows.Count > 0 Then

            For I = 0 To Dt1.Rows.Count - 1

                Descp = "NEGATIVE STOCK : " & Chr(13)

                If Val(Dt1.Rows(I).Item("Item_IdNo").ToString) <> 0 Then

                    CurStk = get_Item_CurrentStock(Cn1, Val(Dt1.Rows(I).Item("Company_Idno").ToString), Val(Dt1.Rows(I).Item("Item_IdNo").ToString), sqltr)

                    If CurStk < 0 Then

                        Descp = Descp & "Item : " & Common_Procedures.Item_IdNoToName(Cn1, Val(Dt1.Rows(I).Item("Item_IdNo").ToString), sqltr)
                        Descp = Descp & Chr(13) & " Stock : " & Val(CurStk)

                        Check_Negative_Stock_Status = True
                        Throw New ApplicationException(Descp)
                        Exit Function

                    End If

                End If

            Next I

        End If

        Dt1.Clear()

        Dt1.Dispose()
        Da1.Dispose()

    End Function

    Public Shared Function VoucherBill_Deletion(ByVal Cn1 As SqlClient.SqlConnection, ByVal ent_idn As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Boolean
        Dim Da1 As New SqlClient.SqlDataAdapter
        Dim Dt1 As New DataTable
        Dim Cmd As New SqlClient.SqlCommand
        Dim vou_bil_code As String = ""
        Dim Amt As Double = 0

        vou_bil_code = get_FieldValue(Cn1, "Voucher_Bill_Head", "VoUcher_Bill_Code", "(Entry_Identification = '" & Trim(ent_idn) & "')", , sqltr)

        Amt = 0
        Da1 = New SqlClient.SqlDataAdapter("Select sum(amount) from voucher_bill_details where voucher_bill_code = '" & Trim(vou_bil_code) & "' and entry_identification <> '" & Trim(ent_idn) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da1.SelectCommand.Transaction = sqltr
        End If
        Dt1 = New DataTable
        Da1.Fill(Dt1)
        If Dt1.Rows.Count > 0 Then
            If IsDBNull(Dt1.Rows(0)(0).ToString) = False Then Amt = Val(Dt1.Rows(0)(0).ToString)
        End If
        Dt1.Clear()

        If Val(Amt) <> 0 Then
            VoucherBill_Deletion = False
            Throw New ApplicationException("Already Received/Paid Amount is  Rs." & Trim(Format(Amt, "#########0.00")))
            'Err.Description = "Already Received/Paid Amount is  Rs." & Trim(Format(Amt, "#########0.00"))

        Else

            Cmd.Connection = Cn1

            If IsNothing(sqltr) = False Then
                Cmd.Transaction = sqltr
            End If

            Cmd.CommandText = "update voucher_bill_head set credit_amount = a.credit_amount - b.amount from voucher_bill_head a, voucher_bill_details b where b.entry_identification = '" & Trim(ent_idn) & "' and b.crdr_type = 'CR' and a.voucher_bill_code = b.voucher_bill_code"
            Cmd.ExecuteNonQuery()

            Cmd.CommandText = "update voucher_bill_head set debit_amount = a.debit_amount - b.amount from voucher_bill_head a, voucher_bill_details b where b.entry_identification = '" & Trim(ent_idn) & "' and b.crdr_type = 'DR' and a.voucher_bill_code = b.voucher_bill_code"
            Cmd.ExecuteNonQuery()

            Cmd.CommandText = "delete from voucher_bill_details where entry_identification = '" & Trim(ent_idn) & "'"
            Cmd.ExecuteNonQuery()

            Cmd.CommandText = "delete from voucher_bill_head where entry_identification = '" & Trim(ent_idn) & "'"
            Cmd.ExecuteNonQuery()

            VoucherBill_Deletion = True

        End If

        Cmd.Dispose()
        Da1.Dispose()
        Dt1.Dispose()

    End Function

    Public Shared Function Voucher_Deletion(ByVal Cn1 As SqlClient.SqlConnection, ByVal Comp_IdNo As Integer, ByVal Ent_IdnCode As String, Optional ByVal SqlTr As SqlClient.SqlTransaction = Nothing) As Boolean
        Dim cmd As New SqlClient.SqlCommand

        Voucher_Deletion = False

        cmd.Connection = Cn1
        If IsNothing(SqlTr) = False Then
            cmd.Transaction = SqlTr
        End If

        cmd.CommandText = "Delete from Voucher_Head Where Company_IdNo = " & Str(Val(Comp_IdNo)) & " and Voucher_Code = '" & Trim(Ent_IdnCode) & "' and Entry_Identification = '" & Trim(Ent_IdnCode) & "'"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Delete from Voucher_Details Where Company_IdNo = " & Str(Val(Comp_IdNo)) & " and Voucher_Code = '" & Trim(Ent_IdnCode) & "' and Entry_Identification = '" & Trim(Ent_IdnCode) & "'"
        cmd.ExecuteNonQuery()

        Voucher_Deletion = True

    End Function


    Public Shared Function AccountsGroup_CodeToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vAccountsGroup_CD As String) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vAccountsGroup_Nm As String

        Da = New SqlClient.SqlDataAdapter("Select AccountsGroup_Name from AccountsGroup_Head where Parent_Idno = '" & Trim(vAccountsGroup_CD) & "'", Cn1)
        Dt = New DataTable
        Da.Fill(Dt)

        vAccountsGroup_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vAccountsGroup_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        AccountsGroup_CodeToName = Trim(vAccountsGroup_Nm)

    End Function
    Public Shared Function Show_CompanyCondition_for_Report(ByVal Cn1 As SqlClient.SqlConnection) As Boolean
        Dim da As New SqlClient.SqlDataAdapter
        Dim dt1 As New DataTable
        Dim NoofComps As Integer
        Dim CompCondt As String

        Show_CompanyCondition_for_Report = False

        Try

            CompCondt = ""
            If Trim(UCase(Common_Procedures.User.Type)) <> "UNACCOUNT" Then
                CompCondt = "(Company_Type <> 'UNACCOUNT')"
            End If

            da = New SqlClient.SqlDataAdapter("select count(*) from Company_Head Where " & CompCondt & IIf(Trim(CompCondt) <> "", " and ", "") & " Company_IdNo <> 0", Cn1)
            dt1 = New DataTable
            da.Fill(dt1)

            NoofComps = 0
            If dt1.Rows.Count > 0 Then
                If IsDBNull(dt1.Rows(0)(0).ToString) = False Then
                    NoofComps = Val(dt1.Rows(0)(0).ToString)
                End If
            End If
            dt1.Clear()

            If Val(NoofComps) > 1 Then

                Show_CompanyCondition_for_Report = True

            End If


        Catch ex As Exception
            'MessageBox.Show(ex.Message, "DOES NOT SHOW...", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            dt1.Dispose()
            da.Dispose()

        End Try

    End Function

    Public Shared Function Voucher_Updation(ByVal Cn1 As SqlClient.SqlConnection, ByVal Vou_Type As String, ByVal Comp_IdNo As Integer, ByVal Ent_IdnCode As String, ByVal Ref_No As String, ByVal Vou_Date As Date, ByVal Par_BilNo As String, ByVal Led_IDNos As String, ByVal Vou_Amts As String, ByRef ErrMsg As String, Optional ByVal SqlTr As SqlClient.SqlTransaction = Nothing) As Boolean
        Dim cmd As New SqlClient.SqlCommand
        Dim vforOrdBy As Double = 0
        Dim LedAr() As String, AmtAr() As String
        Dim db_idno As Integer = 0
        Dim cr_idno As Integer = 0
        Dim vTotCrAmt As Double = 0
        Dim vTotDrAmt As Double = 0
        Dim Mx_DrAmt As Double = 0
        Dim Mx_CrAmt As Double = 0
        Dim i As Integer = 0
        Dim Sno As Integer = 0
        Dim EntID As String = ""
        Dim Nr As Integer = 0

        Voucher_Updation = False
        ErrMsg = ""

        vforOrdBy = Val(Common_Procedures.OrderBy_CodeToValue(Ref_No))

        LedAr = Split(Led_IDNos, "|")
        AmtAr = Split(Vou_Amts, "|")

        If UBound(LedAr) <> UBound(AmtAr) Then
            ErrMsg = "Invalid Voucher Posting, mismatch of ledger and Amount details"
            Exit Function
        End If

        db_idno = 0 : cr_idno = 0
        Mx_DrAmt = 0 : Mx_CrAmt = 0
        vTotDrAmt = 0 : vTotCrAmt = 0

        For i = 0 To UBound(LedAr)

            If Val(LedAr(i)) <> 0 And Val(AmtAr(i)) <> 0 Then

                If Val(AmtAr(i)) < 0 Then
                    If (db_idno = 0 Or Math.Abs(Val(AmtAr(i))) > Mx_DrAmt) Then
                        db_idno = Val(LedAr(i))
                        Mx_DrAmt = Math.Abs(Val(AmtAr(i)))
                    End If
                    vTotDrAmt = vTotDrAmt + Format(Val(AmtAr(i)), "###########0.00")
                End If

                If Val(AmtAr(i)) > 0 Then
                    If (cr_idno = 0 Or Math.Abs(Val(AmtAr(i))) > Mx_CrAmt) Then
                        cr_idno = Val(LedAr(i))
                        Mx_CrAmt = Math.Abs(Val(AmtAr(i)))
                    End If
                    vTotCrAmt = vTotCrAmt + Format(Val(AmtAr(i)), "###########0.00")
                End If

            End If

        Next

        vTotDrAmt = Format(Val(vTotDrAmt), "#########0.00")
        vTotCrAmt = Format(Val(vTotCrAmt), "#########0.00")

        If Math.Abs(vTotDrAmt) <> Math.Abs(vTotCrAmt) Then
            ErrMsg = "Invalid Voucher Amount - Debit and Credit amount not equal"
            Exit Function
        End If

        EntID = Left(Trim(Ent_IdnCode), 6) & Trim(Ref_No)

        cmd.Connection = Cn1
        If IsNothing(SqlTr) = False Then
            cmd.Transaction = SqlTr
        End If

        cmd.CommandText = "Delete from Voucher_Head Where Company_IdNo = " & Str(Val(Comp_IdNo)) & " and Voucher_Code = '" & Trim(Ent_IdnCode) & "' and Entry_Identification = '" & Trim(Ent_IdnCode) & "'"
        Nr = cmd.ExecuteNonQuery()
        cmd.CommandText = "Delete from Voucher_Details Where Company_IdNo = " & Str(Val(Comp_IdNo)) & " and Voucher_Code = '" & Trim(Ent_IdnCode) & "' and Entry_Identification = '" & Trim(Ent_IdnCode) & "'"
        cmd.ExecuteNonQuery()

        If Val(vTotDrAmt) <> 0 And Val(vTotCrAmt) <> 0 Then

            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@VouchDate", Vou_Date.Date)

            cmd.CommandText = "Insert into Voucher_Head ( Voucher_Code         ,                 For_OrderByCode                       ,         Company_IdNo       ,          Voucher_No   ,                   For_OrderBy                         ,         Voucher_Type    , Voucher_Date,          Creditor_Idno   ,          Debtor_Idno     ,              Total_VoucherAmount                       ,          Narration       , Indicate,             Year_For_Report                               ,       Entry_Identification ,           Entry_ID   , Voucher_Receipt_Code ,Salesman_IdNo) " & _
                                        "   Values ('" & Trim(Ent_IdnCode) & "', " & Str(Format(Val(vforOrdBy), "###########0.00")) & ", " & Str(Val(Comp_IdNo)) & ", '" & Trim(Ref_No) & "', " & Str(Format(Val(vforOrdBy), "###########0.00")) & ", '" & Trim(Vou_Type) & "',   @VouchDate, " & Str(Val(cr_idno)) & ", " & Str(Val(db_idno)) & ", " & Str(Format(Val(vTotDrAmt), "###########0.00")) & " , '" & Trim(Par_BilNo) & "',     1   , " & Str(Val(Year(Common_Procedures.Company_FromDate))) & ", '" & Trim(Ent_IdnCode) & "', '" & Trim(EntID) & "',             ''       ," & Common_Procedures.User.IdNo & ") "
            cmd.ExecuteNonQuery()

            Sno = 0

            For i = 0 To UBound(LedAr)

                If Val(LedAr(i)) <> 0 And Val(AmtAr(i)) <> 0 Then

                    Sno = Sno + 1
                    cmd.CommandText = "Insert into Voucher_Details (         Voucher_Code      ,                 For_OrderByCode                       ,          Company_IdNo      ,        Voucher_No     ,                For_OrderBy                            ,       Voucher_Type      , Voucher_Date,           SL_No      ,          Ledger_IdNo      ,              Voucher_Amount                         ,         Narration        ,             Year_For_Report                               ,   Entry_Identification     ,           Entry_ID   ) " & _
                                      "            Values          ('" & Trim(Ent_IdnCode) & "', " & Str(Format(Val(vforOrdBy), "###########0.00")) & ", " & Str(Val(Comp_IdNo)) & ", '" & Trim(Ref_No) & "', " & Str(Format(Val(vforOrdBy), "###########0.00")) & ", '" & Trim(Vou_Type) & "',  @VouchDate , " & Str(Val(Sno)) & ", " & Str(Val(LedAr(i))) & ", " & Str(Format(Val(AmtAr(i)), "##########0.00")) & ", '" & Trim(Par_BilNo) & "', " & Str(Val(Year(Common_Procedures.Company_FromDate))) & ", '" & Trim(Ent_IdnCode) & "', '" & Trim(EntID) & "')"
                    cmd.ExecuteNonQuery()

                End If

            Next i

        End If

        cmd.Dispose()

        Voucher_Updation = True

    End Function

 

    Public Shared Sub maskEdit_Date_ON_DelBackSpace(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs, ByVal mskOldText As String, ByVal mskSelStrt As Integer)
        Dim vmRetTxt As String = ""
        Dim vmRetSelStrt As Integer = -1

        If e.KeyCode = 46 Or e.KeyCode = 8 Then

            If e.KeyCode = 46 Then
                If mskSelStrt <= 2 Then
                    vmRetTxt = "  " & Microsoft.VisualBasic.Mid(mskOldText, 3, Len(mskOldText))
                    vmRetSelStrt = 0
                ElseIf mskSelStrt >= 3 And mskSelStrt <= 5 Then
                    vmRetTxt = Microsoft.VisualBasic.Left(mskOldText, 3) & "  " & Microsoft.VisualBasic.Mid(mskOldText, 6, Len(mskOldText))
                    vmRetSelStrt = 3
                Else
                    vmRetTxt = Microsoft.VisualBasic.Left(mskOldText, 6)
                    vmRetSelStrt = 6
                End If

                sender.Text = vmRetTxt
                sender.SelectionStart = vmRetSelStrt

            ElseIf e.KeyCode = 8 Then
                If mskSelStrt > 0 Then
                    vmRetTxt = Microsoft.VisualBasic.Left(mskOldText, mskSelStrt - 1) & " " & Microsoft.VisualBasic.Mid(mskOldText, mskSelStrt + 1, Len(mskOldText))
                Else
                    vmRetTxt = mskOldText
                End If

                sender.Text = vmRetTxt

                If mskSelStrt > 0 Then
                    sender.SelectionStart = mskSelStrt - 1
                End If

            End If

        End If

    End Sub
    Public Shared Function Month_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vMnth_IdNo As Integer, Optional ByVal SqlTr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vMnth_Name As String

        Da = New SqlClient.SqlDataAdapter("select Month_Name from Month_Head where Month_IdNo = " & Str(Val(vMnth_IdNo)), Cn1)
        If IsNothing(SqlTr) = False Then
            Da.SelectCommand.Transaction = SqlTr
        End If
        Da.Fill(Dt)

        vMnth_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vMnth_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Month_IdNoToName = Trim(vMnth_Name)

    End Function

    Public Shared Function Month_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vMnth_Name As String, Optional ByVal SqlTr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vMnth_IdNo As Integer

        Da = New SqlClient.SqlDataAdapter("select Month_IdNo from Month_Head where Month_Name = '" & Trim(vMnth_Name) & "'", Cn1)
        If IsNothing(SqlTr) = False Then
            Da.SelectCommand.Transaction = SqlTr
        End If
        Da.Fill(Dt)

        vMnth_IdNo = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vMnth_IdNo = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Month_NameToIdNo = Val(vMnth_IdNo)

    End Function
    Public Shared Function Month_IdNoToShortName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vMnth_IdNo As Integer, Optional ByVal SqlTr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vMnth_Name As String

        Da = New SqlClient.SqlDataAdapter("select Month_ShortName from Month_Head where Month_IdNo = " & Str(Val(vMnth_IdNo)), Cn1)
        If IsNothing(SqlTr) = False Then
            Da.SelectCommand.Transaction = SqlTr
        End If
        Da.Fill(Dt)

        vMnth_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vMnth_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Month_IdNoToShortName = Trim(vMnth_Name)

    End Function
    Public Shared Function Month_ShortNameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vMnth_Name As String, Optional ByVal SqlTr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vMnth_IdNo As Integer

        Da = New SqlClient.SqlDataAdapter("select Month_IdNo from Month_Head where Month_ShortName = '" & Trim(vMnth_Name) & "'", Cn1)
        If IsNothing(SqlTr) = False Then
            Da.SelectCommand.Transaction = SqlTr
        End If
        Da.Fill(Dt)

        vMnth_IdNo = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vMnth_IdNo = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Month_ShortNameToIdNo = Val(vMnth_IdNo)

    End Function
    Public Shared Function Accept_AlphaNumericOnlyWithSlash(ByVal KeyAscii_Value As Integer) As Integer
        Accept_AlphaNumericOnlyWithSlash = 0
        If (KeyAscii_Value = 47 Or (KeyAscii_Value >= 48 And KeyAscii_Value <= 57)) Or (KeyAscii_Value >= 65 And KeyAscii_Value <= 90) Or (KeyAscii_Value >= 97 And KeyAscii_Value <= 122) Or KeyAscii_Value = 13 Or KeyAscii_Value = 8 Or KeyAscii_Value = 9 Then
            Accept_AlphaNumericOnlyWithSlash = KeyAscii_Value
        End If
    End Function

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
            'myConnection_string = "Data Source=" & Trim(Common_Procedures.ServerName) & ":3389;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False"
            'myConnection_string = "Data Source=" & Trim(Common_Procedures.ServerName) & ",1033;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False"


            'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";Integrated Security=True"
            'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=169.254.147.41,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False"
            'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Trusted_Connection=True;MultipleActiveResultSets=true;"
            'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";"
            'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(DBName) & ";Integrated Security=True"

        ElseIf Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "WIN" Then
            myConnection_string = "Data Source=" & Trim(Common_Procedures.ServerName) & ";Initial Catalog=" & Trim(DBName) & ";Integrated Security=True;Connect Timeout=60"

        Else
            myConnection_string = "Data Source=" & Trim(Common_Procedures.ServerName) & ";Initial Catalog=" & Trim(DBName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False;Connect Timeout=60"

        End If


        'If Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "IP" Then

        '    'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";Integrated Security=True"
        '    Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False;Connect Timeout=120"

        '    'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=169.254.147.41,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False"

        '    'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Trusted_Connection=True;MultipleActiveResultSets=true;"
        '    'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";"
        '    'Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";Integrated Security=True"

        'ElseIf Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "WIN" Then
        '    Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=" & Trim(Common_Procedures.ServerName) & ";Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";Integrated Security=True;Connect Timeout=120"

        'Else
        '    Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=" & Trim(Common_Procedures.ServerName) & ";Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False;Connect Timeout=120"

        'End If

        ''If Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "IP" Then
        ''    Common_Procedures.ConnectionString_CompanyGroupdetails = "Data Source=122.165.215.63,1433;Network Library=DBMSSOCN;Initial Catalog=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Integrated Security=False"
        ''ElseIf Trim(UCase(Common_Procedures.ServerWindowsLogin)) = "WIN" Then
        ''    Common_Procedures.ConnectionString_CompanyGroupdetails = "Persist Security Info=False;Integrated Security=SSPI;database=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";server=" & Trim(Common_Procedures.ServerName) & ";Connect Timeout=120"
        ''Else
        ''    Common_Procedures.ConnectionString_CompanyGroupdetails = "Persist Security Info=False;Integrated Security=SSPI;database=" & Trim(Common_Procedures.CompanyDetailsDataBaseName) & ";server=" & Trim(Common_Procedures.ServerName) & ";User ID=sa;Password=" & Trim(Common_Procedures.ServerPassword) & ";Connect Timeout=120"
        ''End If

        'Dim myConnection As New SqlClient.SqlConnection()
        'myConnection.ConnectionString = "Persist Security Info=False;Integrated Security=SSPI;database=northwind;server=mySQLServer;Connect Timeout=30"
        'myConnection.Open()

        Create_Sql_ConnectionString = Trim(myConnection_string)

    End Function
    Public Shared Sub Drop_Column_Default_Constraint(ByVal Cn1 As SqlClient.SqlConnection, ByVal TblName As String, ByVal FldName As String)
        Dim Cmd As New SqlClient.SqlCommand
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt1 As New DataTable
        Dim DF_ConsName As String

        Try

            DF_ConsName = ""

            Da = New SqlClient.SqlDataAdapter("select d.name as Default_ConstraintName from sysobjects a inner join dbo.syscolumns c on a.id = c.id inner join dbo.sysobjects d on c.cdefault = d.id Where a.name = '" & Trim(TblName) & "' and c.name = '" & Trim(FldName) & "'", Cn1)
            Dt1 = New DataTable
            Da.Fill(Dt1)
            If Dt1.Rows.Count > 0 Then
                If IsDBNull(Dt1.Rows(0).Item("Default_ConstraintName").ToString) = False Then
                    If Trim(Dt1.Rows(0).Item("Default_ConstraintName").ToString) <> "" Then DF_ConsName = Dt1.Rows(0).Item("Default_ConstraintName").ToString
                End If
            End If
            Dt1.Clear()

            If Trim(DF_ConsName) <> "" Then

                Cmd.Connection = Cn1

                Cmd.CommandText = "ALTER TABLE [dbo].[" & Trim(TblName) & "] DROP CONSTRAINT " & Trim(DF_ConsName)
                Cmd.ExecuteNonQuery()

            End If

        Catch ex As Exception
            '---MessageBox.Show(ex.Message, "ERROR IN DROPPING CONSTRAINT....", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Da.Dispose()
            Dt1.Dispose()
            Cmd.Dispose()

        End Try

    End Sub
    Public Shared Function Salary_PaymentType_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSaPyTy_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSaPyTy_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Salary_Payment_Type_IdNo from PayRoll_Salary_Payment_Type_Head where Salary_Payment_Type_Name = '" & Trim(vSaPyTy_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vSaPyTy_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSaPyTy_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Salary_PaymentType_NameToIdNo = Val(vSaPyTy_ID)

    End Function

    Public Shared Function Salary_PaymentType_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vSaPyTy_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vSaPyTy_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Salary_Payment_Type_Name from PayRoll_Salary_Payment_Type_Head where Salary_Payment_Type_IdNo = " & Str(Val(vSaPyTy_ID)), Cn1)
        Da.Fill(Dt)

        vSaPyTy_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vSaPyTy_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Salary_PaymentType_IdNoToName = Trim(vSaPyTy_Nm)

    End Function

    Public Shared Function Employee_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vEmployee_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vEmployee_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Employee_IdNo from PayRoll_Employee_Head where Employee_Name = '" & Trim(vEmployee_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vEmployee_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vEmployee_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Employee_NameToIdNo = Val(vEmployee_ID)

    End Function

    Public Shared Function Employee_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vEmployee_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vEmployee_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Employee_Name from PayRoll_Employee_Head where Employee_IdNo = " & Str(Val(vEmployee_ID)), Cn1)
        Da.Fill(Dt)

        vEmployee_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vEmployee_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Employee_IdNoToName = Trim(vEmployee_Nm)

    End Function
    Public Shared Function Shift_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vShift_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vShift_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Shift_IdNo from Shift_Head where Shift_Name = '" & Trim(vShift_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vShift_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vShift_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Shift_NameToIdNo = Val(vShift_ID)

    End Function

    Public Shared Function Shift_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vShift_IdNo As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vShift_Name As String = ""

        Da = New SqlClient.SqlDataAdapter("select Shift_Name from Shift_Head where Shift_IdNo = " & Str(Val(vShift_IdNo)), Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vShift_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vShift_Name = Dt.Rows(0)(0).ToString
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Shift_IdNoToName = vShift_Name

    End Function

    Public Shared Function Category_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCategory_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCategory_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Category_IdNo from PayRoll_Category_Head where Category_Name = '" & Trim(vCategory_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vCategory_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCategory_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Category_NameToIdNo = Val(vCategory_ID)

    End Function

    Public Shared Function Category_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vCategory_IdNo As Integer, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vCategory_Name As String = ""

        Da = New SqlClient.SqlDataAdapter("select Category_Name from PayRoll_Category_Head where Category_IdNo = " & Str(Val(vCategory_IdNo)), Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vCategory_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vCategory_Name = Dt.Rows(0)(0).ToString
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Category_IdNoToName = vCategory_Name

    End Function




    Public Shared Function Department_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vDepartment_Name As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vDepartment_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Department_IdNo from Department_Head where Department_Name = '" & Trim(vDepartment_Name) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vDepartment_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vDepartment_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Department_NameToIdNo = Val(vDepartment_ID)

    End Function

    Public Shared Function Department_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vDepartment_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vDepartment_Name As String

        Da = New SqlClient.SqlDataAdapter("select Department_Name from Department_Head where Department_IdNo = " & Str(Val(vDepartment_ID)), Cn1)
        Da.Fill(Dt)

        vDepartment_Name = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vDepartment_Name = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Department_IdNoToName = Trim(vDepartment_Name)

    End Function

    Public Shared Function get_Server_SystemName() As String
        Dim InstNm As String = ""
        Dim ServerNm As String = ""

        If InStr(1, Common_Procedures.ServerName, "\") > 0 Then
            InstNm = Right(Common_Procedures.ServerName, Len(Common_Procedures.ServerName) - InStr(1, Common_Procedures.ServerName, "\"))

            ServerNm = Replace(Trim(UCase(Common_Procedures.ServerName)), Trim(UCase("\" & InstNm)), "")
        Else
            ServerNm = Trim(UCase(Common_Procedures.ServerName))
        End If


        get_Server_SystemName = ServerNm

    End Function

    Public Shared Function is_ServerSystem() As Boolean

        Dim InstNm As String

        InstNm = Right(Common_Procedures.ServerName, Len(Common_Procedures.ServerName) - InStr(1, Common_Procedures.ServerName, "\"))

        is_ServerSystem = False
        If Trim(UCase(Common_Procedures.ServerName)) = Trim(UCase(Trim(SystemInformation.ComputerName))) Or Trim(UCase(Common_Procedures.ServerName)) = Trim(UCase(Trim(SystemInformation.ComputerName) & "\PLit")) Or Trim(UCase(Common_Procedures.ServerName)) = Trim(UCase(Trim(SystemInformation.ComputerName) & "\" & Trim(InstNm))) Then
            is_ServerSystem = True
        End If

    End Function

    Public Shared Function is_Database_File_Exists(ByVal DbName As String) As Boolean
        Dim cn1 As SqlClient.SqlConnection
        Dim da1 As New SqlClient.SqlDataAdapter
        Dim dt1 As New DataTable
        Dim mdf_filname As String = "", ldf_filname As String = "", FlNm As String = ""
        Dim SysNm As String

        is_Database_File_Exists = False
        Err.Description = ""

        Try

            cn1 = New SqlClient.SqlConnection(Common_Procedures.ConnectionString_Master)
            cn1.Open()

            da1 = New SqlClient.SqlDataAdapter("Select * from sysdatabases where name = '" & Trim(DbName) & "'", cn1)
            dt1 = New DataTable
            da1.Fill(dt1)

            If dt1.Rows.Count > 0 Then

                Call get_DataBase_MdfLdf_FileNames(DbName, mdf_filname, ldf_filname)

                If Trim(mdf_filname) = "" Then
                    Err.Description = "database file does not exists"
                    Exit Function
                End If
                If Trim(ldf_filname) = "" Then
                    Err.Description = "database file does not exists"
                    Exit Function
                End If

                FlNm = Trim(mdf_filname)
                If Common_Procedures.is_ServerSystem = False Then

                    SysNm = Common_Procedures.get_Server_SystemName
                    FlNm = "\\" & Trim(SysNm) & "\" & Trim(Replace(mdf_filname, ":\", "\"))

                    'If InStr(1, "\mssql\data\") > 0 Then
                    '    FldrNm()

                    'End If

                    If File.Exists(FlNm) = False Then
                        Err.Description = "database file does not exists"
                        Exit Function
                    End If

                    'Dim sFile As New FileInfo(FlNm)

                    ''FileInfo sFile = new FileInfo(@"\\server\share\file.xml")
                    ''bool fileExist = sFile.Exists;

                    'If sFile.Exists = False Then
                    '    Err.Description = "database file does not exists"
                    '    Exit Function
                    'End If


                Else
                    If File.Exists(FlNm) = False Then
                        Err.Description = "database file does not exists"
                        Exit Function
                    End If

                End If



                FlNm = Trim(ldf_filname)
                If Common_Procedures.is_ServerSystem = False Then
                    SysNm = Common_Procedures.get_Server_SystemName

                    FlNm = "\\" & Trim(SysNm) & "\" & Trim(Replace(ldf_filname, ":\", "\"))
                End If
                If File.Exists(FlNm) = False Then
                    Err.Description = "database file does not exists"
                    Exit Function
                End If

            Else
                Err.Description = Trim(DbName) & " does not exists"
                Exit Function

            End If

            dt1.Dispose()
            da1.Dispose()

            cn1.Close()
            cn1 = Nothing

            is_Database_File_Exists = True

        Catch ex As Exception
            MessageBox.Show("Select Company Group Name", "INVALID COMPANY GROUP SELECTION....", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)

        End Try

    End Function

    Public Shared Sub get_DataBase_MdfLdf_FileNames(ByVal DbName As String, ByRef MDF_FileName As String, ByRef LDF_FileName As String)
        Dim CnMas As SqlClient.SqlConnection
        Dim Da1 As SqlClient.SqlDataAdapter
        Dim Dt1 As DataTable
        Dim Da2 As SqlClient.SqlDataAdapter
        Dim Dt2 As DataTable

        Dim DefPath As String

        CnMas = New SqlClient.SqlConnection(Common_Procedures.ConnectionString_Master)
        CnMas.Open()

        MDF_FileName = ""
        LDF_FileName = ""

        Da1 = New SqlClient.SqlDataAdapter("SELECT * FROM sysdatabases WHERE name = '" & Trim(DbName) & "'", CnMas)
        Dt1 = New DataTable
        Da1.Fill(Dt1)

        If Dt1.Rows.Count > 0 Then

            MDF_FileName = Dt1.Rows(0).Item("FileName").ToString
            If InStr(1, LCase(MDF_FileName), "_data.mdf") > 0 Then
                LDF_FileName = Replace(LCase(MDF_FileName), "_data.mdf", "_log.ldf")
            Else
                LDF_FileName = Replace(LCase(MDF_FileName), ".mdf", "_log.ldf")
            End If


            'If Common_Procedures.is_ServerSystem = True Then
            '    If File.Exists(LDF_FileName) = False Then
            '        LDF_FileName = Replace(LCase(MDF_FileName), "_data.mdf", "_log.ldf")
            '        If File.Exists(LDF_FileName) = False Then
            '            GoTo 100
            '        End If
            '    End If
            'End If


        Else

100:
            Da2 = New SqlClient.SqlDataAdapter("SELECT * FROM sysdatabases WHERE name = 'master'", CnMas)
            Dt2 = New DataTable
            Da2.Fill(Dt2)

            If Dt2.Rows.Count > 0 Then

                DefPath = Replace(LCase(Dt2.Rows(0).Item("FileName").ToString), "\master_data.mdf", "")
                DefPath = Replace(LCase(Dt2.Rows(0).Item("FileName").ToString), "\master.mdf", "")

                MDF_FileName = Trim(DefPath) & "\" & Trim(DbName) & ".mdf"
                LDF_FileName = Trim(DefPath) & "\" & Trim(DbName) & "_log.ldf"

            End If
            Dt2.Dispose()
        End If
        Dt1.Dispose()

        CnMas.Close()
        CnMas = Nothing

    End Sub

    Public Shared Function Is_InterState_Party(ByVal Cn1 As SqlClient.SqlConnection, ByVal CompIdNo As Integer, ByVal LedIdNo As Integer) As Boolean
        Dim CompStateIdNo As Integer = 0
        Dim LedStateIdNo As Integer = 0
        Dim sts As Boolean = False

        CompStateIdNo = Val(Common_Procedures.get_FieldValue(Cn1, "Company_Head", "Company_State_IdNo", "(Company_IdNo = " & Str(Val(CompIdNo)) & ")"))
        LedStateIdNo = Val(Common_Procedures.get_FieldValue(Cn1, "Ledger_Head", "State_Idno", "(Ledger_IdNo = " & Str(Val(LedIdNo)) & ")"))

        If Val(CompStateIdNo) = 0 Or Val(LedStateIdNo) = 0 Then
            sts = False
        ElseIf Val(CompStateIdNo) = Val(LedStateIdNo) Then
            sts = False
        Else
            sts = True
        End If

        Is_InterState_Party = sts

    End Function

    Public Shared Sub FillRegionRectangle(ByVal e As System.Drawing.Printing.PrintPageEventArgs, ByVal X1axis As Decimal, ByVal Y1axis As Decimal, ByVal X2axis As Decimal, ByVal Y2axis As Decimal)
        Dim Hght As Double = 0
        Dim Wdth As Double = 0

        ' Create solid brush.
        Dim blueBrush As New SolidBrush(Color.FromArgb(235, 235, 235))

        Wdth = X2axis - X1axis
        Hght = Y2axis - Y1axis

        ' Create rectangle for region.
        Dim fillRect As New Rectangle(X1axis, Y1axis, Wdth, Hght)

        ' Create region for fill.
        Dim fillRegion As New [Region](fillRect)

        ' Fill region to screen.
        e.Graphics.FillRegion(blueBrush, fillRegion)

    End Sub
    Public Shared Function SimpleEmployee_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vEmployee_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vEmployee_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Employee_IdNo from Employee_Head where Employee_Name = '" & Trim(vEmployee_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vEmployee_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vEmployee_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        SimpleEmployee_NameToIdNo = Val(vEmployee_ID)

    End Function

    Public Shared Function SimpleEmployee_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vEmployee_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vEmployee_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Employee_Name from Employee_Head where Employee_IdNo = " & Str(Val(vEmployee_ID)), Cn1)
        Da.Fill(Dt)

        vEmployee_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vEmployee_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        SimpleEmployee_IdNoToName = Trim(vEmployee_Nm)

    End Function
    Public Shared Function Site_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vGender_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vGender_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Site_IdNo from Site_Head where Site_Name = '" & Trim(vGender_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vGender_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vGender_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Site_NameToIdNo = Val(vGender_ID)

    End Function

    Public Shared Function Site_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vGender_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vGender_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Site_Name from Site_Head where Site_IdNo = " & Str(Val(vGender_ID)), Cn1)
        Da.Fill(Dt)

        vGender_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vGender_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Site_IdNoToName = Trim(vGender_Nm)

    End Function

    Public Shared Function Encrypt(ByVal plainText As String, ByVal passPhrase As String, ByVal saltValue As String) As String
        Dim hashAlgorithm As String = "SHA1"

        Dim passwordIterations As Integer = 2
        Dim initVector As String = "@1B2c3D4e5F6g7H8"
        Dim keySize As Integer = 256

        Dim initVectorBytes As Byte() = Encoding.ASCII.GetBytes(initVector)
        Dim saltValueBytes As Byte() = Encoding.ASCII.GetBytes(saltValue)

        Dim plainTextBytes As Byte() = Encoding.UTF8.GetBytes(plainText)

        Dim mypassword As New Rfc2898DeriveBytes(passPhrase, saltValueBytes, passwordIterations)

        Dim keyBytes As Byte() = mypassword.GetBytes(keySize \ 8)
        Dim symmetricKey As New RijndaelManaged()

        symmetricKey.Mode = CipherMode.CBC

        Dim encryptor As ICryptoTransform = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)

        Dim memoryStream As New MemoryStream()
        Dim cryptoStream As New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)

        cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)
        cryptoStream.FlushFinalBlock()
        Dim cipherTextBytes As Byte() = memoryStream.ToArray()
        memoryStream.Close()
        cryptoStream.Close()
        Dim cipherText As String = Convert.ToBase64String(cipherTextBytes)
        Return cipherText
    End Function

    Public Shared Function Decrypt(ByVal cipherText As String, ByVal passPhrase As String, ByVal saltValue As String) As String
        Dim plainText As String = ""

        Try

            'Dim passPhrase As String = "T.ThanGesWaran"
            'Dim saltValue As String = "N.VaRaLakshmi"
            Dim hashAlgorithm As String = "SHA1"

            Dim passwordIterations As Integer = 2
            Dim initVector As String = "@1B2c3D4e5F6g7H8"
            Dim keySize As Integer = 256
            ' Convert strings defining encryption key characteristics into byte
            ' arrays. Let us assume that strings only contain ASCII codes.
            ' If strings include Unicode characters, use Unicode, UTF7, or UTF8
            ' encoding.
            Dim initVectorBytes As Byte() = Encoding.ASCII.GetBytes(initVector)
            Dim saltValueBytes As Byte() = Encoding.ASCII.GetBytes(saltValue)

            ' Convert our ciphertext into a byte array.
            Dim cipherTextBytes As Byte() = Convert.FromBase64String(cipherText)

            ' First, we must create a password, from which the key will be 
            ' derived. This password will be generated from the specified 
            ' passphrase and salt value. The password will be created using
            ' the specified hash algorithm. Password creation can be done in
            ' several iterations.
            Dim mypassword As New Rfc2898DeriveBytes(passPhrase, saltValueBytes, passwordIterations)
            'Dim mypassword As New PasswordDeriveBytes(passPhrase, saltValueBytes, hashAlgorithm, passwordIterations)

            ' Use the password to generate pseudo-random bytes for the encryption
            ' key. Specify the size of the key in bytes (instead of bits).
            Dim keyBytes As Byte() = mypassword.GetBytes(keySize \ 8)

            ' Create uninitialized Rijndael encryption object.
            Dim symmetricKey As New RijndaelManaged()

            ' It is reasonable to set encryption mode to Cipher Block Chaining
            ' (CBC). Use default options for other symmetric key parameters.
            symmetricKey.Mode = CipherMode.CBC

            ' Generate decryptor from the existing key bytes and initialization 
            ' vector. Key size will be defined based on the number of the key 
            ' bytes.
            Dim decryptor As ICryptoTransform = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)

            ' Define memory stream which will be used to hold encrypted data.
            Dim memoryStream As New MemoryStream(cipherTextBytes)

            ' Define cryptographic stream (always use Read mode for encryption).
            Dim cryptoStream As New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)

            ' Since at this point we don't know what the size of decrypted data
            ' will be, allocate the buffer long enough to hold ciphertext;
            ' plaintext is never longer than ciphertext.
            Dim plainTextBytes As Byte() = New Byte(cipherTextBytes.Length - 1) {}

            ' Start decrypting.
            Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)

            ' Close both streams.
            memoryStream.Close()
            cryptoStream.Close()

            ' Convert decrypted data into a string. 
            ' Let us assume that the original plaintext string was UTF8-encoded.
            plainText = Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount)

            ' Return decrypted string.   
            Return plainText

        Catch ex As Exception
            plainText = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
            Return plainText

        End Try

    End Function

    Public Shared Function GetDriveSerialNumber(ByVal DriveLetter As String) As String
        Try
            Dim disk As ManagementObject = New ManagementObject(String.Format("Win32_Logicaldisk='{0}'", DriveLetter))
            Dim VolumeName As String = disk.Properties("VolumeName").Value.ToString()
            Dim SerialNumber As String = disk.Properties("VolumeSerialnumber").Value.ToString()
            Return SerialNumber
            'Return SerialNumber.Insert(4, "-")

        Catch ex As Exception
            Return ""

        End Try
    End Function

    Public Shared Function is_OfficeSystem() As Boolean
        Dim STS As Boolean = False

        Try

            Common_Procedures.DriveVolumeSerialName = ""
            Try
                Common_Procedures.DriveVolumeSerialName = Common_Procedures.GetDriveSerialNumber("D:")
            Catch ex As Exception
                '---
            End Try

            '---                                                       Server                                                               Mukilan                                                              Gopal               
            If Trim(UCase(Common_Procedures.DriveVolumeSerialName)) = "AEAA0163" Or Trim(UCase(Common_Procedures.DriveVolumeSerialName)) = "F0203A37" Or Trim(UCase(Common_Procedures.DriveVolumeSerialName)) = "424638AA" Then
                STS = True
            End If

            Return STS

        Catch ex As Exception
            Return ""

        End Try
    End Function
    Public Shared Function Scheme_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vColour_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vColour_ID As Integer

        Da = New SqlClient.SqlDataAdapter("select Scheme_IdNo from Scheme_Head where Scheme_Name = '" & Trim(vColour_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vColour_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vColour_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Scheme_NameToIdNo = Val(vColour_ID)

    End Function

    Public Shared Function Scheme_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vColour_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vColour_Nm As String

        Da = New SqlClient.SqlDataAdapter("select Scheme_Name from Scheme_Head where Scheme_IdNo = " & Str(Val(vColour_ID)), Cn1)
        Da.Fill(Dt)

        vColour_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vColour_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        Scheme_IdNoToName = Trim(vColour_Nm)

    End Function

    Public Shared Sub Report_Party_Balance_AgeWise(ByVal Con As SqlClient.SqlConnection, ByVal vAcGrp_IdNo As Integer, ByVal vUPtoDate As Date, ByVal RptCondt As String, ByVal AgingRange As String)
        'Dim Cmd As New SqlClient.SqlCommand
        'Dim Da As New SqlClient.SqlDataAdapter
        'Dim Dt As New DataTable
        'Dim Dt1 As New DataTable
        'Dim Dt2 As New DataTable
        'Dim Bal As String = 0
        'Dim Amt As String = 0
        'Dim I As Long = 0
        'Dim N As Long = 0
        'Dim B() As String
        'Dim S As String, vOLDval As String

        'Cmd.Connection = Con

        'Cmd.Parameters.Clear()
        'Cmd.Parameters.AddWithValue("@uptodate", vUPtoDate)

        'Cmd.CommandText = "Truncate table EntryTemp"
        'Cmd.ExecuteNonQuery()

        'Cmd.CommandText = "Truncate table ReportTempSub"
        'Cmd.ExecuteNonQuery()

        'If Val(vAcGrp_IdNo) = 0 Then vAcGrp_IdNo = 10

        'Cmd.CommandText = "Insert into EntryTemp (int1, name1, currency1 ) Select b.ledger_idno, b.ledger_name, sum(a.voucher_amount) from voucher_details a, ledger_head b, company_head tz where " & RptCondt & IIf(RptCondt <> "", " and ", "") & " a.voucher_date <= @uptodate and b.AccountsGroup_IdNo = " & Str(Val(vAcGrp_IdNo)) & " and a.ledger_idno = b.ledger_idno and a.company_idno = tz.company_idno group by b.ledger_idno, b.ledger_name having sum(a.voucher_amount) <> 0"
        'Cmd.ExecuteNonQuery()

        'Da = New SqlClient.SqlDataAdapter("Select int1 as Ledger_IdNo, name1 as Ledger_Name, abs(sum(currency1)) as BalanceAmount from EntryTemp group by int1, name1 having sum(currency1) < 0 Order by int1, name1", Con)
        'Dt1 = New DataTable
        'Da.Fill(Dt1)

        'Bal = 0
        'If Dt1.Rows.Count > 0 Then

        '    For I = 0 To Dt1.Rows.Count - 1

        '        If Math.Abs(Val(Dt1.Rows(I).Item("BalanceAmount").ToString)) <> 0 Then

        '            Cmd.Parameters.Clear()
        '            Cmd.Parameters.AddWithValue("@uptodate", vUPtoDate)

        '            Cmd.CommandText = "Select a.voucher_date, abs(sum(a.voucher_amount)) as VouAmt from voucher_details a INNER JOIN company_head tz ON a.company_idno = tz.company_idno Where " & RptCondt & IIf(RptCondt <> "", " and ", "") & " a.voucher_date <= @uptodate and a.ledger_idno = " & Str(Val(Dt1.Rows(I).Item("Ledger_IdNo").ToString)) & " and a.voucher_amount < 0 group by a.voucher_date having sum(a.voucher_amount) <> 0 Order by a.voucher_date desc"
        '            Da = New SqlClient.SqlDataAdapter(Cmd)
        '            Dt2 = New DataTable
        '            Da.Fill(Dt2)

        '            Bal = Math.Abs(Val(Dt1.Rows(I).Item("BalanceAmount").ToString))

        '            If Dt2.Rows.Count > 0 Then

        '                For j = 0 To Dt2.Rows.Count - 1

        '                    N = DateDiff(DateInterval.Day, CDate(Dt2.Rows(j).Item("voucher_date")), vUPtoDate)

        '                    Amt = 0
        '                    If Val(Bal) >= Val(Dt2.Rows(j).Item("VouAmt").ToString) Then
        '                        Amt = Val(Dt2.Rows(j).Item("VouAmt").ToString)
        '                    Else
        '                        Amt = Val(Bal)
        '                    End If
        '                    Bal = Format(Val(Bal) - Val(Amt), "##########0.00")

        '                    Cmd.CommandText = "Insert into ReportTempSub ( int1, int2, currency1 ) values ( " & Str(Val(Dt1.Rows(I).Item("Ledger_IdNo").ToString)) & ", " & Str(Val(N)) & ", " & Str(Val(Amt)) & " ) "
        '                    Cmd.ExecuteNonQuery()

        '                    If Val(Bal) <= 0 Then
        '                        Exit For
        '                    End If

        '                Next j
        '            End If
        '            Dt2.Clear()

        '            If Bal <> 0 Then
        '                N = 99999
        '                Amt = Bal
        '                Cmd.CommandText = "Insert into ReportTempSub ( int1, int2, currency1 ) values ( " & Str(Val(Dt1.Rows(I).Item("Ledger_IdNo").ToString)) & ", " & Str(Val(N)) & ", " & Str(Val(Amt)) & " ) "
        '                Cmd.ExecuteNonQuery()
        '            End If

        '        End If

        '    Next I

        'End If
        'Dt1.Clear()

        'Cmd.CommandText = "Truncate table ReportTemp"
        'Cmd.ExecuteNonQuery()

        'If Trim(AgingRange) = "" Then AgingRange = "30,60,90,120"


        'vOLDval = "0"
        'B = Split(AgingRange, ",")

        'For I = 0 To UBound(B)
        '    If Val(B(I)) <> 0 Then
        '        S = Trim(Val(vOLDval)) & " TO " & Trim(Val(B(I)))

        '        Cmd.CommandText = "Insert into ReportTemp ( Int1 ,   Name1      ,             Int2       ,         Name2     , currency1    ) " & _
        '                            " Select        b.Ledger_IdNo, b.Ledger_Name, " & Str(Val(I) + 1) & ", '" & Trim(S) & "',      0    from ReportTempSub a INNER JOIN Ledger_Head b ON a.Int1 = b.Ledger_IdNo group by b.Ledger_IdNo, b.Ledger_Name having sum(a.currency1) <> 0"
        '        Cmd.ExecuteNonQuery()

        '        Cmd.CommandText = "Insert into ReportTemp ( Int1 ,   Name1      ,             Int2       ,         Name2     ,        currency1    ) " & _
        '                            " Select        b.Ledger_IdNo, b.Ledger_Name, " & Str(Val(I) + 1) & ", '" & Trim(S) & "',   sum(a.currency1)   from ReportTempSub a INNER JOIN Ledger_Head b ON a.Int1 = b.Ledger_IdNo where Int2 Between " & Str(Val(vOLDval)) & " and " & Str(Val(B(I))) & " group by b.Ledger_IdNo, b.Ledger_Name"
        '        Cmd.ExecuteNonQuery()

        '        vOLDval = Val(B(I)) + 1

        '    End If
        'Next I

        'If Val(vOLDval) <> 0 Then

        '    S = "ABV " & Trim(Val(vOLDval) - 1)

        '    Cmd.CommandText = "Insert into ReportTemp ( Int1 ,   Name1      ,             Int2       ,         Name2     , currency1    ) " & _
        '                        " Select        b.Ledger_IdNo, b.Ledger_Name, " & Str(Val(I) + 1) & ", '" & Trim(S) & "',      0    from ReportTempSub a INNER JOIN Ledger_Head b ON a.Int1 = b.Ledger_IdNo group by b.Ledger_IdNo, b.Ledger_Name having sum(a.currency1) <> 0"
        '    Cmd.ExecuteNonQuery()

        '    Cmd.CommandText = "Insert into ReportTemp ( Int1 ,   Name1      ,             Int2       ,         Name2     ,        currency1    ) " & _
        '                        " Select        b.Ledger_IdNo, b.Ledger_Name, " & Str(Val(I) + 1) & ", '" & Trim(S) & "',   sum(a.currency1)   from ReportTempSub a INNER JOIN Ledger_Head b ON a.Int1 = b.Ledger_IdNo where Int2 >= " & Str(Val(vOLDval)) & " group by b.Ledger_IdNo, b.Ledger_Name"
        '    Cmd.ExecuteNonQuery()

        'End If

    End Sub
    Public Shared Sub MouseDrag_Form(ByVal Drag_sts As Boolean, ByVal Move_Sts As Boolean, ByRef X_Position As Integer, ByRef Y_Position As Integer)


        'mouse down
        If Drag_sts = True Then
            Common_Procedures.MousePositionX = Windows.Forms.Cursor.Position.X - X_Position
            Common_Procedures.MousePositionY = Windows.Forms.Cursor.Position.Y - Y_Position
        End If

        'mouse move
        If Move_Sts = True And Common_Procedures.MousePositionX <> 0 And Common_Procedures.MousePositionY <> 0 Then
            X_Position = Windows.Forms.Cursor.Position.X - Common_Procedures.MousePositionX
            Y_Position = Windows.Forms.Cursor.Position.Y - Common_Procedures.MousePositionY
        End If

        'mouse up
        If Drag_sts = False And Move_Sts = False Then
            Common_Procedures.MousePositionX = 0
            Common_Procedures.MousePositionY = 0
        End If
    End Sub

    Public Shared Sub reShape(ByVal sender As Object, ByVal CurveSize As Integer)
        Dim p As New Drawing2D.GraphicsPath
        Dim Width As Integer = sender.Width
        Dim Height As Integer = sender.Height

        p.StartFigure()

        p.AddArc(New Rectangle(0, 0, CurveSize, CurveSize), 180, 90)
        p.AddLine(CurveSize, 0, Width - CurveSize, 0)

        p.AddArc(New Rectangle(Width - CurveSize, 0, CurveSize, CurveSize), -90, 90)
        p.AddLine(Width, CurveSize, Width, Height - CurveSize)

        p.AddArc(New Rectangle(Width - CurveSize, Height - CurveSize, CurveSize, CurveSize), 0, 90)
        p.AddLine(Width - CurveSize, Height, CurveSize, Height)

        p.AddArc(New Rectangle(0, Height - CurveSize, CurveSize, CurveSize), 90, 90)

        p.CloseFigure()

        sender.Region = New Region(p)

    End Sub
    Public Shared Sub BackgroundGradient(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs, Optional ByVal Color1 As Drawing.Color = Nothing, Optional ByVal Color2 As Drawing.Color = Nothing)


        Dim g As Graphics = e.Graphics
        Dim p1 As Point = sender.ClientRectangle.Location
        Dim p2 As Point = New Point(sender.ClientRectangle.Right, sender.ClientRectangle.Bottom)


        If Color1 = Nothing And Color2 = Nothing Then

            Using bg As New System.Drawing.Drawing2D.LinearGradientBrush(p1, p2, Color.WhiteSmoke, Color.WhiteSmoke)
                g.FillRectangle(bg, e.ClipRectangle)

            End Using
        Else
            Using bg As New System.Drawing.Drawing2D.LinearGradientBrush(p1, p2, Color1, Color2)
                g.FillRectangle(bg, e.ClipRectangle)

            End Using
        End If




    End Sub
    Public Shared Function USER_NameToIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vUSER_Nm As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal vCmpGrpIdNo As Integer = 0) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vUnit_ID As Integer
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select User_IdNo from " & Trim(vDB_Name) & "User_Head where User_Name = '" & Trim(vUSER_Nm) & "'", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vUnit_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vUnit_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        USER_NameToIdNo = Val(vUnit_ID)

    End Function
    Public Shared Function User_IdNoToName(ByVal Cn1 As SqlClient.SqlConnection, ByVal vUser_ID As Integer) As String
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vUser_Nm As String

        Da = New SqlClient.SqlDataAdapter("select User_Name from User_Head where User_IdNo = " & Str(Val(vUser_ID)), Cn1)
        Da.Fill(Dt)

        vUser_Nm = ""
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vUser_Nm = Trim(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        User_IdNoToName = Trim(vUser_Nm)

    End Function
    Public Shared Function get_Rate_From_PriceList(ByVal Cn1 As SqlClient.SqlConnection, ByVal cur_date As Date, ByVal Item_id As Integer, Optional ByRef Tax_Status As Boolean = False) As Double
        Dim cmd As New SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim cur_rate As Double

        Try
            cmd.Connection = Cn1


            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@PriceDate", cur_date)

            cmd.CommandText = "select a.Tax_Status from Price_List_Head a left outer join Price_List_Details b on a.price_list_idno =b.price_list_idno  where a.Price_List_Date = @PriceDate and b.item_idno =" & Item_id & ""
            dr = cmd.ExecuteReader

            cur_rate = 0
            If dr.HasRows Then
                If dr.Read Then
                    If IsDBNull(dr(0).ToString) = False Then
                        If Val(dr(0).ToString) > 0 Then
                            Tax_Status = True
                        Else
                            Tax_Status = False
                        End If
                    End If
                End If
            End If
            dr.Close()

            cmd.CommandText = "select Rate from Price_List_Details where Price_List_Date = @PriceDate and item_idno =" & Item_id & ""
            dr = cmd.ExecuteReader

            cur_rate = 0
            If dr.HasRows Then
                If dr.Read Then
                    If IsDBNull(dr(0).ToString) = False Then
                        cur_rate = Val(dr(0).ToString)
                    End If
                End If
            End If
            dr.Close()

            ' If Trim(Common_Procedures.settings.CustomerCode) = "4025" Then
            If Val(cur_rate) = 0 Then
                cmd.CommandText = "select top 1 a.Rate, b.Tax_Status from Price_List_Details a  left outer join Price_List_Head b on  a.price_list_idno =b.price_list_idno  where  a.item_idno =" & Item_id & " ORDER BY a.price_list_idno DESC"
                'cmd.CommandText = "select top 1 Rate from Price_List_Details  where  item_idno =" & Item_id & " ORDER BY price_list_idno DESC"
                dr = cmd.ExecuteReader

                cur_rate = 0
                If dr.HasRows Then
                    If dr.Read Then
                        If IsDBNull(dr(0).ToString) = False Then
                            cur_rate = Val(dr(0).ToString)
                        End If
                        If IsDBNull(dr(1).ToString) = False Then
                            Tax_Status = Val(dr(1).ToString)
                        End If
                    End If
                End If
                dr.Close()
            End If


            'End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "FOR  MOVING...", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        get_Rate_From_PriceList = cur_rate

    End Function
    Public Shared Function get_WholeSaleRate_From_PriceList(ByVal Cn1 As SqlClient.SqlConnection, ByVal cur_date As Date, ByVal Item_id As Integer, Optional ByRef Tax_Status As Boolean = False) As Double
        Dim cmd As New SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim cur_rate As Double

        Try
            cmd.Connection = Cn1


            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@PriceDate", cur_date)

            cmd.CommandText = "select a.Tax_Status from Price_List_Head a left outer join Price_List_Details b on a.price_list_idno =b.price_list_idno  where a.Price_List_Date = @PriceDate and b.item_idno =" & Item_id & ""
            dr = cmd.ExecuteReader

            cur_rate = 0
            If dr.HasRows Then
                If dr.Read Then
                    If IsDBNull(dr(0).ToString) = False Then
                        If Val(dr(0).ToString) > 0 Then
                            Tax_Status = True
                        Else
                            Tax_Status = False
                        End If
                    End If
                End If
            End If
            dr.Close()

            cmd.CommandText = "select Wholesale_Rate from Price_List_Details where Price_List_Date = @PriceDate and item_idno =" & Item_id & ""
            dr = cmd.ExecuteReader

            cur_rate = 0
            If dr.HasRows Then
                If dr.Read Then
                    If IsDBNull(dr(0).ToString) = False Then
                        cur_rate = Val(dr(0).ToString)
                    End If
                End If
            End If
            dr.Close()

            ' If Trim(Common_Procedures.settings.CustomerCode) = "4025" Then
            If Val(cur_rate) = 0 Then
                cmd.CommandText = "select top 1 Wholesale_Rate from Price_List_Details where  item_idno =" & Item_id & " ORDER BY price_list_idno DESC"
                dr = cmd.ExecuteReader

                cur_rate = 0
                If dr.HasRows Then
                    If dr.Read Then
                        If IsDBNull(dr(0).ToString) = False Then
                            cur_rate = Val(dr(0).ToString)
                        End If
                    End If
                End If
                dr.Close()
            End If


            'End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "FOR  MOVING...", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        get_WholeSaleRate_From_PriceList = cur_rate

    End Function
    Public Shared Function get_Rate_From_PriceList_TaxStatus(ByVal Cn1 As SqlClient.SqlConnection, ByVal cur_date As Date, ByVal Item_id As Integer) As Boolean
        Dim cmd As New SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim Tax_Status As Boolean = False

        Try
            cmd.Connection = Cn1


            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@PriceDate", cur_date)

            cmd.CommandText = "select a.Tax_Status from Price_List_Head a left outer join Price_List_Details b on a.price_list_idno =b.price_list_idno  where a.Price_List_Date = @PriceDate and b.item_idno =" & Item_id & ""
            dr = cmd.ExecuteReader


            If dr.HasRows Then
                If dr.Read Then
                    If IsDBNull(dr(0).ToString) = False Then
                        If Val(dr(0).ToString) > 0 Then
                            Tax_Status = True
                        Else
                            Tax_Status = False
                        End If
                    End If
                End If
            End If
            dr.Close()


        Catch ex As Exception
            MessageBox.Show(ex.Message, "FOR  MOVING...", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        get_Rate_From_PriceList_TaxStatus = Tax_Status

    End Function
    Public Shared Function item_IdnoToUnitIdNo(ByVal Cn1 As SqlClient.SqlConnection, ByVal vItem_idno As String, Optional ByVal sqltr As SqlClient.SqlTransaction = Nothing, Optional ByVal vCmpGrpIdNo As Integer = 0) As Integer
        Dim Da As New SqlClient.SqlDataAdapter
        Dim Dt As New DataTable
        Dim vUnit_ID As Integer
        Dim vDB_Name As String = ""

        vDB_Name = ""
        If vCmpGrpIdNo <> 0 Then
            vDB_Name = Common_Procedures.get_Company_DataBaseName(vCmpGrpIdNo)
            vDB_Name = vDB_Name & ".."
        End If

        Da = New SqlClient.SqlDataAdapter("select b.Unit_IdNo from " & Trim(vDB_Name) & "item_Head a left outer join unit_head b on a.unit_idno = b.unit_idno where item_idno = " & Val(vItem_idno) & "", Cn1)
        If IsNothing(sqltr) = False Then
            Da.SelectCommand.Transaction = sqltr
        End If
        Da.Fill(Dt)

        vUnit_ID = 0
        If Dt.Rows.Count > 0 Then
            If IsDBNull(Dt.Rows(0)(0).ToString) = False Then
                vUnit_ID = Val(Dt.Rows(0)(0).ToString)
            End If
        End If

        Dt.Dispose()
        Da.Dispose()

        item_IdnoToUnitIdNo = Val(vUnit_ID)

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
    
End Class
