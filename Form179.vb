Public Class Form179
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim Linez As Integer = 0
    Dim sYear1 As Int16 = 0
    Dim sYear2 As Int16 = 0
    Dim sWeek1 As Int16 = 0
    Dim sWeek2 As Int16 = 0
    Dim sMonth1 As Int16 = 0
    Dim sMonth2 As Int16 = 0
    Dim ExcelPath As String = String.Empty
    Dim TotalWeek As Int16 = 0
    Dim TotalMonth As Int16 = 0

    Private Sub Form179_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.GroupBox2.Enabled = False
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        TextBox1.Text = Now.Year & "01"
        TextBox2.Text = Now.Year & "52"
        Label1.Text = "未读入"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            ExcelPath = OpenFileDialog1.FileName
            xExcel = New Microsoft.Office.Interop.Excel.Application
            xWorkBook = xExcel.Workbooks.Open(ExcelPath)
            Ws = xWorkBook.Sheets(5)
            Linez = 3
            Button2.Enabled = True
            Label1.Text = "已读入"
        Else
            Button2.Enabled = False
            Return
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If


        CreateTempDB()
        Label1.Text = "处理临时表中"
        sYear1 = Strings.Left(TextBox1.Text, 4)
        sYear2 = Strings.Left(TextBox2.Text, 4)
        sWeek1 = Strings.Right(TextBox1.Text, 2)
        sWeek2 = Strings.Right(TextBox2.Text, 2)
        TotalWeek = (sYear2 - sYear1) * 53 + (sWeek2 - sWeek1)  ' 0 表示 1週
        Dim BB As Integer = Ws.UsedRange.Rows.Count
        For i As Integer = 3 To BB Step 1
            Label1.Text = i
            Label1.Refresh()
            oRng = Ws.Range(Ws.Cells(i, 1), Ws.Cells(i, 1))
            Dim PND As String = oRng.Value
            If IsNothing(PND) Then
                Continue For
            End If
            oRng = Ws.Range(Ws.Cells(i, 4), Ws.Cells(i, 4))
            Dim Unit1 As String = oRng.Value
            For j As Integer = 0 To TotalWeek Step 1
                oRng = Ws.Range(Ws.Cells(i, 6 + j), Ws.Cells(i, 6 + j))
                If IsNumeric(oRng.Value2) Then
                    Dim Q1 As Integer = Decimal.Round(oRng.Value2, 0, MidpointRounding.AwayFromZero)
                    If Q1 = 0 Then
                        Continue For
                    End If
                    Dim tWeek As Int16 = j + sWeek1
                    Dim tYear As Int16 = 0
                    Select Case tWeek
                        Case Is < 54
                            tYear = sYear1
                        Case 54 To 106
                            tYear = sYear1 + 1
                            tWeek = tWeek - 53
                        Case Is > 106
                            tYear = sYear1 + 2
                            tWeek = tWeek - 106
                    End Select
                    oCommand.CommandText = "INSERT INTO budget2020_1 VALUES ('" & PND & "','" & Unit1 & "'," & tYear & "," & tWeek & "," & Q1 & ")"
                    Try
                        oCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                        Return
                    End Try
                End If
            Next
        Next
        Label1.Text = "临时表已完成"
        'MsgBox("Done")

    End Sub

    Private Sub CreateTempDB()
        oCommand.CommandText = "DROP TABLE budget2020_1"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        oCommand.CommandText = "create table budget2020_1 (pn varchar2(40), Unit1 varchar2(4), year1 number(5,0), week1 number(5,0), quantity number(12, 0))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommander2.Connection = oConnection
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "SELECT COUNT(*) FROM budget2020_1 "
        Dim HasReport As Integer = oCommand.ExecuteScalar()
        If HasReport <= 0 Then
            MsgBox("No Report")
            Return
        End If
        Label1.Text = "处理临时表中"
        sYear1 = Strings.Left(TextBox1.Text, 4)
        sYear2 = Strings.Left(TextBox2.Text, 4)
        sWeek1 = Strings.Right(TextBox1.Text, 2)
        sWeek2 = Strings.Right(TextBox2.Text, 2)
        TotalWeek = (sYear2 - sYear1) * 53 + (sWeek2 - sWeek1)  ' 0 表示 1週
        If IsNothing(xExcel) Then
            If String.IsNullOrEmpty(ExcelPath) Then
                MsgBox("请先读入范例档")
                Return
            End If
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(ExcelPath)
        Ws = xWorkBook.Sheets(6)
        Linez = 3
        oCommand.CommandText = "Select pn,ima02,ima021,unit1,ima44"
        For x As Int16 = 0 To TotalWeek Step 1
            oCommand.CommandText += ",sum(t" & x & ") as t" & x
        Next
        oCommand.CommandText += " from (select pn,ima02,ima021,unit1,ima44"
        For x As Int16 = 0 To TotalWeek Step 1
            Dim tWeek As Int16 = x + sWeek1
            Dim tYear As Int16 = 0
            Select Case tWeek
                Case Is < 54
                    tYear = sYear1
                Case 54 To 106
                    tYear = sYear1 + 1
                    tWeek = tWeek - 53
                Case Is > 106
                    tYear = sYear1 + 2
                    tWeek = tWeek - 106
            End Select
            oCommand.CommandText += ",(case when year1 = " & tYear & " and week1 = " & tWeek & " then quantity else 0 end) as t" & x
        Next
        oCommand.CommandText += " from budget2020_1 left join ima_file on pn = ima01 ) group by pn,ima02,ima021,unit1,ima44 order by pn"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(Linez, 1) = oReader.Item(0)
                Ws.Cells(Linez, 2) = oReader.Item(1)
                Ws.Cells(Linez, 3) = oReader.Item(2)
                Ws.Cells(Linez, 4) = oReader.Item(3)
                Dim ER1 As Decimal = 0
                If oReader.Item(3) <> oReader.Item(4) Then
                    oCommander2.CommandText = "select Round(smd06/smd04,5) from smd_file where smd01 = '" & oReader.Item(0) & "' and smd02 = '" & oReader.Item(3) & "' and smd03 = '" & oReader.Item(4) & "'"
                    oCommander2.CommandText += " union all "
                    oCommander2.CommandText += "select Round(smd04/smd06,5) from smd_file where smd01 = '" & oReader.Item(0) & "' and smd02 = '" & oReader.Item(4) & "' and smd03 = '" & oReader.Item(3) & "'"

                    Try
                        ER1 = oCommander2.ExecuteScalar()
                    Catch ex As Exception
                        ER1 = 1
                    End Try

                Else
                    ER1 = 1
                End If
                Ws.Cells(Linez, 5) = ER1
                Ws.Cells(Linez, 6) = oReader.Item(4)

                ' 20200714
                oCommander2.CommandText = "select pmi05 from pmj_file left join pmi_file on pmj01 = pmi01 where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '"
                oCommander2.CommandText += oReader.Item(0) & "' order by pmj09 desc"
                Dim l_pmj05 As String = oCommander2.ExecuteScalar()
                If l_pmj05 = "Y" Or l_pmj05 = "y" Then
                    oCommander2.CommandText = "select pmi01 from pmj_file left join pmi_file on pmj01 = pmi01 where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '"
                    oCommander2.CommandText += oReader.Item(0) & "' order by pmj09 desc"
                    Dim l_pmi01 As String = oCommander2.ExecuteScalar()
                    oCommander2.CommandText = "select nvl(max(pmr05 * nvl(er,1)),0) from pmr_file left join pmj_file on pmr01 = pmj01 left join exchangeratebyyear s1 on s1.year1 = " & sYear1 & " and s1.currency = pmj05 where pmr01 = '" & l_pmi01 & "'"
                Else
                    oCommander2.CommandText = "select pmj07t*nvl(er,1) from pmj_file left join pmi_file on pmj01 = pmi01 left join exchangeratebyyear on pmj05 = currency and year1 =  " & sYear1
                    oCommander2.CommandText += " where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '" & oReader.Item(0) & "' order by pmj09 desc"
                End If

                'oCommander2.CommandText = "select pmj07t*nvl(er,1) from pmj_file left join pmi_file on pmj01 = pmi01 left join exchangeratebyyear on pmj05 = currency and year1 = 2020 "
                'oCommander2.CommandText += "where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '" & oReader.Item(0) & "' order by pmj09 desc"
                Dim Price1 As Decimal = oCommander2.ExecuteScalar()
                Ws.Cells(Linez, 7) = Price1
                Ws.Cells(Linez, 8) = "=SUM(I" & Linez & ":BS" & Linez & ")"
                For Z As Integer = 0 To TotalWeek Step 1
                    Ws.Cells(Linez, 9 + Z) = oReader.Item(5 + Z) * Price1 * ER1
                    Label1.Text = "1-" & Linez & "-" & Z
                    Label1.Refresh()
                Next
                Linez += 1
                Label1.Text = Linez
                Label1.Refresh()
            End While
        End If
        oReader.Close()

        ' 處理廠商
        CreateTempDB1()
        oCommand.CommandText = "select distinct pn from budget2020_1"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim CountC As Integer = 0
            While oReader.Read()
                oCommander2.CommandText = "INSERT INTO budget2020_2 select pmj03,pmi03,pma08 from ( select pmj03,pmi03 from pmj_file left join pmi_file on pmj01 = pmi01 where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '"
                oCommander2.CommandText += oReader.Item(0) & "' order by pmj09 desc ) left join pmc_file on pmi03 = pmc01 left join pma_file on pmc17 = pma01 where rownum = 1"
                Try
                    oCommander2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                CountC += 1
                Label1.Text = "处理临时表2-" & CountC
                Label1.Refresh()
            End While
        End If
        oReader.Close()

        CreateTempDB2()
        oCommand.CommandText = "insert into budget2020_3 Select AD.pn,AD.unit1,ae.azn02,ae.azn04,ae.azn05,AD.quantity from ( Select aa.pn,aa.unit1,aa.year1,aa.week1,(max(ac.azn01) + nvl(ab.dayadd,0)) as d1,quantity from budget2020_1 aa "
        oCommand.CommandText += "left join budget2020_2 ab on aa.pn= ab.pn left join azn_file ac on aa.year1 = ac.azn02 and aa.week1 = ac.azn05 group by aa.pn,aa.unit1,unit1,dayadd,aa.year1,aa.week1,quantity ) AD left join azn_file ae on AD.d1 = ae.azn01 order by pn"

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        ' 查最小付款年月
        oCommand.CommandText = "Select min(year1 || (case when week1 < 10 then '0' || week1 else to_char(week1) end)) from budget2020_3"
        Dim SS1 As String = oCommand.ExecuteScalar()
        sYear1 = Strings.Left(SS1, 4)
        sWeek1 = Strings.Right(SS1, 2)
        oCommand.CommandText = "Select max(year1 || (case when week1 < 10 then '0' || week1 else to_char(week1) end)) from budget2020_3"
        Dim SS2 As String = oCommand.ExecuteScalar()
        sYear2 = Strings.Left(SS2, 4)
        sWeek2 = Strings.Right(SS2, 2)

        TotalWeek = (sYear2 - sYear1) * 53 + (sWeek2 - sWeek1)  ' 0 表示 1週

        Ws = xWorkBook.Sheets(7)
        Linez = 3


        oCommand.CommandText = "Select pn,ima02,ima021,unit1,ima44"
        For x As Int16 = 0 To TotalWeek Step 1
            oCommand.CommandText += ",sum(t" & x & ") as t" & x
        Next
        oCommand.CommandText += " from (select pn,ima02,ima021,unit1,ima44"
        For x As Int16 = 0 To TotalWeek Step 1
            Dim tWeek As Int16 = x + sWeek1
            Dim tYear As Int16 = 0
            Select Case tWeek
                Case Is < 54
                    tYear = sYear1
                Case 54 To 106
                    tYear = sYear1 + 1
                    tWeek = tWeek - 53
                Case Is > 106
                    tYear = sYear1 + 2
                    tWeek = tWeek - 106
                Case Is > 159
                    tYear = sYear1 + 3
                    tWeek = tWeek - 159
            End Select
            oCommand.CommandText += ",(case when year1 = " & tYear & " and week1 = " & tWeek & " then quantity else 0 end) as t" & x
            Ws.Cells(2, 9 + x) = tYear & "W" & tWeek
        Next
        oCommand.CommandText += " from budget2020_3 left join ima_file on pn = ima01 ) group by pn,ima02,ima021,unit1,ima44 order by pn"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(Linez, 1) = oReader.Item(0)
                Ws.Cells(Linez, 2) = oReader.Item(1)
                Ws.Cells(Linez, 3) = oReader.Item(2)
                Ws.Cells(Linez, 4) = oReader.Item(3)
                Dim ER1 As Decimal = 0
                If oReader.Item(3) <> oReader.Item(4) Then
                    oCommander2.CommandText = "select Round(smd06/smd04,5) from smd_file where smd01 = '" & oReader.Item(0) & "' and smd02 = '" & oReader.Item(3) & "' and smd03 = '" & oReader.Item(4) & "'"
                    oCommander2.CommandText += " union all "
                    oCommander2.CommandText += "select Round(smd04/smd06,5) from smd_file where smd01 = '" & oReader.Item(0) & "' and smd02 = '" & oReader.Item(4) & "' and smd03 = '" & oReader.Item(3) & "'"

                    Try
                        ER1 = oCommander2.ExecuteScalar()
                    Catch ex As Exception
                        ER1 = 1
                    End Try

                Else
                    ER1 = 1
                End If
                Ws.Cells(Linez, 5) = ER1
                Ws.Cells(Linez, 6) = oReader.Item(4)

                ' 20200714
                oCommander2.CommandText = "select pmi05 from pmj_file left join pmi_file on pmj01 = pmi01 where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '"
                oCommander2.CommandText += oReader.Item(0) & "' order by pmj09 desc"
                Dim l_pmj05 As String = oCommander2.ExecuteScalar()
                If l_pmj05 = "Y" Or l_pmj05 = "y" Then
                    oCommander2.CommandText = "select pmi01 from pmj_file left join pmi_file on pmj01 = pmi01 where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '"
                    oCommander2.CommandText += oReader.Item(0) & "' order by pmj09 desc"
                    Dim l_pmi01 As String = oCommander2.ExecuteScalar()
                    oCommander2.CommandText = "select nvl(max(pmr05 * nvl(er,1)),0) from pmr_file left join pmj_file on pmr01 = pmj01 left join exchangeratebyyear s1 on s1.year1 = " & sYear1 & " and s1.currency = pmj05 where pmr01 = '" & l_pmi01 & "'"
                Else
                    oCommander2.CommandText = "select pmj07t*nvl(er,1) from pmj_file left join pmi_file on pmj01 = pmi01 left join exchangeratebyyear on pmj05 = currency and year1 =  " & sYear1
                    oCommander2.CommandText += " where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '" & oReader.Item(0) & "' order by pmj09 desc"
                End If

                'oCommander2.CommandText = "select pmj07t*nvl(er,1) from pmj_file left join pmi_file on pmj01 = pmi01 left join exchangeratebyyear on pmj05 = currency and year1 = 2020 "
                'oCommander2.CommandText += "where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '" & oReader.Item(0) & "' order by pmj09 desc"
                Dim Price1 As Decimal = oCommander2.ExecuteScalar()
                Ws.Cells(Linez, 7) = Price1
                Ws.Cells(Linez, 8) = "=SUM(I" & Linez & ":DD" & Linez & ")"
                For Z As Integer = 0 To TotalWeek Step 1
                    Ws.Cells(Linez, 9 + Z) = oReader.Item(5 + Z) * Price1 * ER1
                    Label1.Text = "2-" & Linez & "-" & Z
                    Label1.Refresh()
                Next
                Linez += 1
                Label1.Text = "2-" & Linez
                Label1.Refresh()
            End While
        End If
        oReader.Close()

        ' 第三頁

        ' 查最小付款年月
        oCommand.CommandText = "Select min(year1 || (case when month1 < 10 then '0' || month1 else to_char(month1) end)) from budget2020_3"
        Dim SS3 As String = oCommand.ExecuteScalar()
        sYear1 = Strings.Left(SS3, 4)
        sMonth1 = Strings.Right(SS3, 2)
        oCommand.CommandText = "Select max(year1 || (case when month1 < 10 then '0' || month1 else to_char(month1) end)) from budget2020_3"
        Dim SS4 As String = oCommand.ExecuteScalar()
        sYear2 = Strings.Left(SS4, 4)
        sMonth2 = Strings.Right(SS4, 2)

        TotalMonth = (sYear2 - sYear1) * 12 + (sMonth2 - sMonth1)  ' 0 表示 1月


        Ws = xWorkBook.Sheets(8)
        Linez = 3

        oCommand.CommandText = "Select pn,ima02,ima021,unit1,ima44"
        For x As Int16 = 0 To TotalMonth Step 1
            oCommand.CommandText += ",sum(t" & x & ") as t" & x
        Next
        oCommand.CommandText += " from (select pn,ima02,ima021,unit1,ima44"
        For x As Int16 = 0 To TotalMonth Step 1
            Dim tMonth As Int16 = x + sMonth1
            Dim tYear As Int16 = 0
            Select Case tMonth
                Case Is < 13
                    tYear = sYear1
                Case 13 To 24
                    tYear = sYear1 + 1
                    tMonth = tMonth - 12
                Case 25 To 36
                    tYear = sYear1 + 2
                    tMonth = tMonth - 24
                Case Is > 36
                    tYear = sYear1 + 3
                    tMonth = tMonth - 36
            End Select
            oCommand.CommandText += ",(case when year1 = " & tYear & " and month1 = " & tMonth & " then quantity else 0 end) as t" & x
            Ws.Cells(2, 9 + x) = tYear & "M" & tMonth
        Next
        oCommand.CommandText += " from budget2020_3 left join ima_file on pn = ima01 ) group by pn,ima02,ima021,unit1,ima44 order by pn"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(Linez, 1) = oReader.Item(0)
                Ws.Cells(Linez, 2) = oReader.Item(1)
                Ws.Cells(Linez, 3) = oReader.Item(2)
                Ws.Cells(Linez, 4) = oReader.Item(3)
                Dim ER1 As Decimal = 0
                If oReader.Item(3) <> oReader.Item(4) Then
                    oCommander2.CommandText = "select Round(smd06/smd04,5) from smd_file where smd01 = '" & oReader.Item(0) & "' and smd02 = '" & oReader.Item(3) & "' and smd03 = '" & oReader.Item(4) & "'"
                    oCommander2.CommandText += " union all "
                    oCommander2.CommandText += "select Round(smd04/smd06,5) from smd_file where smd01 = '" & oReader.Item(0) & "' and smd02 = '" & oReader.Item(4) & "' and smd03 = '" & oReader.Item(3) & "'"

                    Try
                        ER1 = oCommander2.ExecuteScalar()
                    Catch ex As Exception
                        ER1 = 1
                    End Try

                Else
                    ER1 = 1
                End If
                Ws.Cells(Linez, 5) = ER1
                Ws.Cells(Linez, 6) = oReader.Item(4)

                ' 20200714
                oCommander2.CommandText = "select pmi05 from pmj_file left join pmi_file on pmj01 = pmi01 where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '"
                oCommander2.CommandText += oReader.Item(0) & "' order by pmj09 desc"
                Dim l_pmj05 As String = oCommander2.ExecuteScalar()
                If l_pmj05 = "Y" Or l_pmj05 = "y" Then
                    oCommander2.CommandText = "select pmi01 from pmj_file left join pmi_file on pmj01 = pmi01 where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '"
                    oCommander2.CommandText += oReader.Item(0) & "' order by pmj09 desc"
                    Dim l_pmi01 As String = oCommander2.ExecuteScalar()
                    oCommander2.CommandText = "select nvl(max(pmr05 * nvl(er,1)),0) from pmr_file left join pmj_file on pmr01 = pmj01 left join exchangeratebyyear s1 on s1.year1 = " & sYear1 & " and s1.currency = pmj05 where pmr01 = '" & l_pmi01 & "'"
                Else
                    oCommander2.CommandText = "select pmj07t*nvl(er,1) from pmj_file left join pmi_file on pmj01 = pmi01 left join exchangeratebyyear on pmj05 = currency and year1 =  " & sYear1
                    oCommander2.CommandText += " where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '" & oReader.Item(0) & "' order by pmj09 desc"
                End If

                'oCommander2.CommandText = "select pmj07t*nvl(er,1) from pmj_file left join pmi_file on pmj01 = pmi01 left join exchangeratebyyear on pmj05 = currency and year1 = 2020 "
                'oCommander2.CommandText += "where pmj01 = pmi01 and pmiconf = 'Y' and pmj03 = '" & oReader.Item(0) & "' order by pmj09 desc"
                Dim Price1 As Decimal = oCommander2.ExecuteScalar()
                Ws.Cells(Linez, 7) = Price1
                Ws.Cells(Linez, 8) = "=SUM(I" & Linez & ":DD" & Linez & ")"
                For Z As Integer = 0 To TotalMonth Step 1
                    Ws.Cells(Linez, 9 + Z) = oReader.Item(5 + Z) * Price1 * ER1
                    Label1.Text = "3-" & Linez & "-" & Z
                    Label1.Refresh()
                Next
                Linez += 1
                Label1.Text = "3-" & Linez
                Label1.Refresh()
            End While
        End If
        oReader.Close()



        SaveExcel()

    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "主要物料进料需求表-购料资金"
        SaveFileDialog1.DefaultExt = ".xlsx"
        Dim SON As DialogResult = SaveFileDialog1.ShowDialog()
        If SON = DialogResult.OK Then
            Dim SFN As String = SaveFileDialog1.FileName
            Ws.SaveAs(SFN, XlFileFormat.xlOpenXMLWorkbook)
        Else
            MsgBox("没有储存文件", MsgBoxStyle.Critical)
        End If
        xWorkBook.Saved = True
        xWorkBook.Close()
        xExcel.Quit()
        If oConnection.State = ConnectionState.Open Then
            Try
                oConnection.Close()
                'Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub CreateTempDB1()
        oCommand.CommandText = "DROP TABLE budget2020_2"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        oCommand.CommandText = "create table budget2020_2 (pn varchar2(40), Vendor varchar2(10), DayAdd number(5,0))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Private Sub CreateTempDB2()
        oCommand.CommandText = "DROP TABLE budget2020_3"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        oCommand.CommandText = "create table budget2020_3 (pn varchar2(40), Unit1 varchar2(4), year1 number(5,0),month1 number(5,0), week1 number(5,0), quantity number(12, 0))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub
End Class