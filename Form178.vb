Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form178
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand3 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim tYear As Int16 = 0
    Dim tDate1 As Date
    Dim tDate2 As Date
    Dim sYear1 As Int16 = 0
    Dim sYear2 As Int16 = 0
    Dim sMonth1 As Int16 = 0
    Dim sMonth2 As Int16 = 0
    Dim mPeriod1 As String = String.Empty
    Dim mPeriod2 As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form178_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        TextBox1.Text = Today.Year
        TextBox3.Text = Today.ToString("yyyyMM")
        TextBox4.Text = Today.ToString("yyyyMM")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If TextBox1.Text.Length <> 4 Then
            MsgBox("ERROR")
            Return
        End If
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
                oCommand3.Connection = oConnection
                oCommand3.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = TextBox1.Text
        tDate1 = Convert.ToDateTime(tYear & "/01/01")
        tDate2 = tDate1.AddYears(1).AddDays(-1)
        sYear1 = Strings.Left(TextBox3.Text, 4)
        sYear2 = Strings.Left(TextBox4.Text, 4)
        sMonth1 = Strings.Right(TextBox3.Text, 2)
        sMonth2 = Strings.Right(TextBox4.Text, 2)
        mPeriod1 = TextBox3.Text
        mPeriod2 = TextBox4.Text

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "销售成本报表" & tYear
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
                Module1.KillExcelProcess(OldExcel)
                MsgBox("Finished")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "Cost Breakdown %"
        Ws.Activate()
        AdjustExcelFormat1()
        oCommand.CommandText = "Select ccc01,ima02,ima021,ima25,sum(ccc61 * -1),sum(ccc63),sum(ccc62 * -1),sum(ccc61 * ccc23a * -1), sum(ccc61 * ccc23b * -1), sum(ccc61 * ccc23c * -1) , sum(ccc61 * ccc23e * -1),sum(ccc61 * ccc23d * -1) from ccc_file "
        oCommand.CommandText += "left join ima_file on ccc01 = ima01 where ccc02 || (case when ccc03 < 10 then '0' || ccc03 else to_char(ccc03) end) between '"
        oCommand.CommandText += mPeriod1 & "' and '" & mPeriod2 & "' and ccc63 <> 0 and ccc61 <> 0 and ccc62 <> 0  and ima06 = '103' "
        oCommand.CommandText += " group by ccc01,ima02,ima021,ima25 order by ccc01"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 2 + i) = oReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()
        ' 加總
        Ws.Cells(LineZ, 5) = "Total"
        Ws.Cells(LineZ, 6) = "=SUM(F5:F" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 6))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ, 13)), Type:=xlFillDefault)
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 8), Ws.Cells(LineZ + 1, 13))
        oRng.NumberFormat = "0%"
        Ws.Cells(LineZ + 1, 5) = "%"
        Ws.Cells(LineZ + 1, 8) = "=IFERROR(H" & LineZ & "/$G$" & LineZ & ",)"
        oRng = Ws.Range(Ws.Cells(LineZ + 1, 8), Ws.Cells(LineZ + 1, 8))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ + 1, 8), Ws.Cells(LineZ + 1, 13)), Type:=xlFillDefault)

        ' 劃線
        oRng = Ws.Range("B5", Ws.Cells(LineZ + 1, 13))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous



        Ws = xWorkBook.Sheets(2)
        Ws.Name = "Unit Price"
        Ws.Activate()
        AdjustExcelFormat2()

        oCommand.CommandText = "Select tc_prl01,ima02,ima021,ima25,tc_prl06,tc_prl03, tc_prl04,er, tc_prl02 from tc_prl_file left join ima_file on tc_prl01 = ima01 "
        oCommand.CommandText += "left join exchangeratebyyear on tc_prl06 = exchangeratebyyear.currency and year1 = " & tYear & " where tc_prl02 > to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select tc_prn01,ima02,ima021,ima25,tc_prn06,tc_prn03, tc_prn04,er, tc_prn02 from tc_prn_file left join ima_file on tc_prn01 = ima01 "
        oCommand.CommandText += "left join exchangeratebyyear on tc_prn06 = exchangeratebyyear.currency and year1 = " & tYear & " where tc_prn02 > to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oCommand.CommandText += " order by tc_prl01,tc_prl02"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            Dim ASS As Int16 = 1
            Dim OLDoeb04 As String = String.Empty
            While oReader.Read()
                Dim NewOEb04 As String = oReader.Item(0)
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                If NewOEb04 <> OLDoeb04 Then
                    ASS = 1
                Else
                    ASS += 1
                    End If
                Ws.Cells(LineZ, 6) = ASS
                Ws.Cells(LineZ, 7) = oReader.Item(4)
                Ws.Cells(LineZ, 8) = oReader.Item(5)
                Ws.Cells(LineZ, 9) = oReader.Item(6) / 100
                Ws.Cells(LineZ, 10) = oReader.Item(7)
                Ws.Cells(LineZ, 11) = "=H" & LineZ & "*I" & LineZ & "*J" & LineZ
                Ws.Cells(LineZ, 12) = oReader.Item(8)
                LineZ += 1
                OLDoeb04 = NewOEb04
                End While
            End If
        oReader.Close()

            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ - 1, 12))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            ' 第三頁
        TransferDB()
        Ws = xWorkBook.Sheets(3)
        Ws.Name = "Customer Demand Combined Qty"
        Ws.Activate()
        AdjustExcelFormat3()

        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity order by pn,max(azn01)"

        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04 order by tc_prm01,max(azn01)"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, 2 + i) = oReader.Item(i)
                Next
                Ws.Cells(LineZ, 21) = "=SUM(I" & LineZ & ":T" & LineZ & ")"
                LineZ += 1
                End While
            End If
            ' 加總
        Ws.Cells(LineZ, 8) = "Total"
        Ws.Cells(LineZ, 9) = "=SUM(I6:I" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)

        oReader.Close()
            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ, 21))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            ' 第四頁

        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(4)
        Ws.Name = " Cost"
        Ws.Activate()
        AdjustExcelFormat4()
        oCommand.CommandText = "select nvl(round(sum(ccc62 * -1) /sum(ccc63),5),0) from ccc_file left join ima_file on ccc01 = ima01 where ccc02 || (case when ccc03 < 10 then '0' || ccc03 else to_char(ccc03) end) between '"
        oCommand.CommandText += mPeriod1 & "' and '" & mPeriod2 & "' and ccc63 <> 0 and ccc61 <> 0 and ccc62 <> 0  and ima06 = '103' "

        Dim PercentageofSales As Decimal = oCommand.ExecuteScalar()
        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),ccc23,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "left join ccc_file on pn = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0"
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity,ccc23 order by pn,max(azn01)"

        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),ccc23,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "left join ccc_file on tc_prm01 = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04,ccc23 order by tc_prm01,max(azn01)"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 8) = oReader.Item(4)
                Ws.Cells(LineZ, 9) = oReader.Item(5)
                Ws.Cells(LineZ, 10) = oReader.Item(6)
                Dim SALESPRICE As Decimal = 0
                If IsDBNull(oReader.Item(7)) Then
                    Ws.Cells(LineZ, 11) = "总成本与总收入%"
                    oCommand2.CommandText = "select tc_prl03 * er * tc_prl04 / 100 from ( select rownum,tc_prl03,tc_prl06,er,tc_prl02,tc_prl04 from ( select tc_prl03,tc_prl06,tc_prl02,tc_prl04 from tc_prl_file where tc_prl01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prl02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                    oCommand2.CommandText += "union all "
                    oCommand2.CommandText += "select tc_prn03,tc_prn06,tc_prn02,tc_prn04 from tc_prn_file where tc_prn01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prn02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  ) "
                    oCommand2.CommandText += "left join exchangeratebyyear on tc_prl06 = currency and year1 = " & tYear & "order by tc_prl02 ) where rownum = 1"
                    SALESPRICE = oCommand2.ExecuteScalar()
                    Ws.Cells(LineZ, 12) = PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * PercentageofSales * SALESPRICE
                Else
                    If sMonth2 < 10 Then
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/0" & sMonth2
                    Else
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/" & sMonth2
                        End If

                    Ws.Cells(LineZ, 12) = oReader.Item(7)
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * oReader.Item(7)
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * oReader.Item(7)
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * oReader.Item(7)
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * oReader.Item(7)
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * oReader.Item(7)
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * oReader.Item(7)
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * oReader.Item(7)
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * oReader.Item(7)
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * oReader.Item(7)
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * oReader.Item(7)
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * oReader.Item(7)
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * oReader.Item(7)
                    End If
                Ws.Cells(LineZ, 25) = "=SUM(M" & LineZ & ":X" & LineZ & ")"

                LineZ += 1
                End While
            End If
            ' 加總
        Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 13))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 25)), Type:=xlFillDefault)

        oReader.Close()
            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ, 25))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            '第五頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(5)
        Ws.Name = "Cost Material"
        Ws.Activate()
        AdjustExcelFormat4()
        oCommand.CommandText = "select nvl(round(sum(ccc61 * ccc23a * -1) /sum(ccc63),5),0) from ccc_file left join ima_file on ccc01 = ima01 where ccc02 || (case when ccc03 < 10 then '0' || ccc03 else to_char(ccc03) end) between '"
        oCommand.CommandText += mPeriod1 & "' and '" & mPeriod2 & "' and ccc63 <> 0 and ccc61 <> 0 and ccc62 <> 0  and ima06 = '103' "
        PercentageofSales = oCommand.ExecuteScalar()

        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),ccc23a,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "left join ccc_file on pn = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0"
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity,ccc23a order by pn,max(azn01)"
        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),ccc23a,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "left join ccc_file on tc_prm01 = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04,ccc23a order by tc_prm01,max(azn01)"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 8) = oReader.Item(4)
                Ws.Cells(LineZ, 9) = oReader.Item(5)
                Ws.Cells(LineZ, 10) = oReader.Item(6)
                Dim SALESPRICE As Decimal = 0
                If IsDBNull(oReader.Item(7)) Then
                    Ws.Cells(LineZ, 11) = "总成本与总收入%"
                    oCommand2.CommandText = "select tc_prl03 * er * tc_prl04 / 100 from ( select rownum,tc_prl03,tc_prl06,er,tc_prl02,tc_prl04 from ( select tc_prl03,tc_prl06,tc_prl02,tc_prl04 from tc_prl_file where tc_prl01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prl02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                    oCommand2.CommandText += "union all "
                    oCommand2.CommandText += "select tc_prn03,tc_prn06,tc_prn02,tc_prn04 from tc_prn_file where tc_prn01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prn02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  ) "
                    oCommand2.CommandText += "left join exchangeratebyyear on tc_prl06 = currency and year1 = " & tYear & "order by tc_prl02 ) where rownum = 1"
                    SALESPRICE = oCommand2.ExecuteScalar()
                    Ws.Cells(LineZ, 12) = PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * PercentageofSales * SALESPRICE
                Else
                    If sMonth2 < 10 Then
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/0" & sMonth2
                    Else
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/" & sMonth2
                        End If
                    Ws.Cells(LineZ, 12) = oReader.Item(7)
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * oReader.Item(7)
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * oReader.Item(7)
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * oReader.Item(7)
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * oReader.Item(7)
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * oReader.Item(7)
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * oReader.Item(7)
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * oReader.Item(7)
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * oReader.Item(7)
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * oReader.Item(7)
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * oReader.Item(7)
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * oReader.Item(7)
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * oReader.Item(7)
                    End If
                Ws.Cells(LineZ, 25) = "=SUM(M" & LineZ & ":X" & LineZ & ")"

                LineZ += 1
                End While
            End If
            ' 加總
        Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 13))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 25)), Type:=xlFillDefault)

        oReader.Close()
            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ, 25))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            '第六頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(6)
        Ws.Name = "Cost  LB"
        Ws.Activate()
        AdjustExcelFormat4()
        oCommand.CommandText = "select nvl(round(sum(ccc61 * ccc23b * -1) /sum(ccc63),5),0) from ccc_file left join ima_file on ccc01 = ima01 where ccc02 || (case when ccc03 < 10 then '0' || ccc03 else to_char(ccc03) end) between '"
        oCommand.CommandText += mPeriod1 & "' and '" & mPeriod2 & "' and ccc63 <> 0 and ccc61 <> 0 and ccc62 <> 0  and ima06 = '103' "
        PercentageofSales = oCommand.ExecuteScalar()

        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),ccc23b,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "left join ccc_file on pn = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0"
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity,ccc23b order by pn,max(azn01)"
        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),ccc23b,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "left join ccc_file on tc_prm01 = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04,ccc23b order by tc_prm01,max(azn01)"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 8) = oReader.Item(4)
                Ws.Cells(LineZ, 9) = oReader.Item(5)
                Ws.Cells(LineZ, 10) = oReader.Item(6)
                Dim SALESPRICE As Decimal = 0
                If IsDBNull(oReader.Item(7)) Then
                    Ws.Cells(LineZ, 11) = "总成本与总收入%"
                    oCommand2.CommandText = "select tc_prl03 * er * tc_prl04 / 100 from ( select rownum,tc_prl03,tc_prl06,er,tc_prl02,tc_prl04 from ( select tc_prl03,tc_prl06,tc_prl02,tc_prl04 from tc_prl_file where tc_prl01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prl02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                    oCommand2.CommandText += "union all "
                    oCommand2.CommandText += "select tc_prn03,tc_prn06,tc_prn02,tc_prn04 from tc_prn_file where tc_prn01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prn02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  ) "
                    oCommand2.CommandText += "left join exchangeratebyyear on tc_prl06 = currency and year1 = " & tYear & "order by tc_prl02 ) where rownum = 1"
                    SALESPRICE = oCommand2.ExecuteScalar()
                    Ws.Cells(LineZ, 12) = PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * PercentageofSales * SALESPRICE
                Else
                    If sMonth2 < 10 Then
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/0" & sMonth2
                    Else
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/" & sMonth2
                        End If
                    Ws.Cells(LineZ, 12) = oReader.Item(7)
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * oReader.Item(7)
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * oReader.Item(7)
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * oReader.Item(7)
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * oReader.Item(7)
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * oReader.Item(7)
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * oReader.Item(7)
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * oReader.Item(7)
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * oReader.Item(7)
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * oReader.Item(7)
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * oReader.Item(7)
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * oReader.Item(7)
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * oReader.Item(7)
                    End If
                Ws.Cells(LineZ, 25) = "=SUM(M" & LineZ & ":X" & LineZ & ")"

                LineZ += 1
                End While
            End If
            ' 加總
        Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 13))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 25)), Type:=xlFillDefault)

        oReader.Close()
            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ, 25))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            '第七頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(7)
        Ws.Name = "Cost  OVH 1"
        Ws.Activate()
        AdjustExcelFormat4()
        oCommand.CommandText = "select nvl(round(sum(ccc61 * ccc23c * -1) /sum(ccc63),5),0) from ccc_file left join ima_file on ccc01 = ima01 where ccc02 || (case when ccc03 < 10 then '0' || ccc03 else to_char(ccc03) end) between '"
        oCommand.CommandText += mPeriod1 & "' and '" & mPeriod2 & "' and ccc63 <> 0 and ccc61 <> 0 and ccc62 <> 0  and ima06 = '103' "
        PercentageofSales = oCommand.ExecuteScalar()

        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),ccc23c,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "left join ccc_file on pn = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0"
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity,ccc23c order by pn,max(azn01)"
        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),ccc23c,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "left join ccc_file on tc_prm01 = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04,ccc23c order by tc_prm01,max(azn01)"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 8) = oReader.Item(4)
                Ws.Cells(LineZ, 9) = oReader.Item(5)
                Ws.Cells(LineZ, 10) = oReader.Item(6)
                Dim SALESPRICE As Decimal = 0
                If IsDBNull(oReader.Item(7)) Then
                    Ws.Cells(LineZ, 11) = "总成本与总收入%"
                    oCommand2.CommandText = "select tc_prl03 * er * tc_prl04 / 100 from ( select rownum,tc_prl03,tc_prl06,er,tc_prl02,tc_prl04 from ( select tc_prl03,tc_prl06,tc_prl02,tc_prl04 from tc_prl_file where tc_prl01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prl02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                    oCommand2.CommandText += "union all "
                    oCommand2.CommandText += "select tc_prn03,tc_prn06,tc_prn02,tc_prn04 from tc_prn_file where tc_prn01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prn02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  ) "
                    oCommand2.CommandText += "left join exchangeratebyyear on tc_prl06 = currency and year1 = " & tYear & "order by tc_prl02 ) where rownum = 1"
                    SALESPRICE = oCommand2.ExecuteScalar()
                    Ws.Cells(LineZ, 12) = PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * PercentageofSales * SALESPRICE
                Else
                    If sMonth2 < 10 Then
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/0" & sMonth2
                    Else
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/" & sMonth2
                        End If
                    Ws.Cells(LineZ, 12) = oReader.Item(7)
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * oReader.Item(7)
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * oReader.Item(7)
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * oReader.Item(7)
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * oReader.Item(7)
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * oReader.Item(7)
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * oReader.Item(7)
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * oReader.Item(7)
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * oReader.Item(7)
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * oReader.Item(7)
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * oReader.Item(7)
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * oReader.Item(7)
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * oReader.Item(7)
                    End If
                Ws.Cells(LineZ, 25) = "=SUM(M" & LineZ & ":X" & LineZ & ")"

                LineZ += 1
                End While
            End If
            ' 加總
        Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 13))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 25)), Type:=xlFillDefault)

        oReader.Close()
            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ, 25))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            '第八頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(8)
        Ws.Name = "Cost  OVH 2"
        Ws.Activate()
        AdjustExcelFormat4()
        oCommand.CommandText = "select nvl(round(sum(ccc61 * ccc23e * -1) /sum(ccc63),5),0) from ccc_file left join ima_file on ccc01 = ima01 where ccc02 || (case when ccc03 < 10 then '0' || ccc03 else to_char(ccc03) end) between '"
        oCommand.CommandText += mPeriod1 & "' and '" & mPeriod2 & "' and ccc63 <> 0 and ccc61 <> 0 and ccc62 <> 0  and ima06 = '103' "
        PercentageofSales = oCommand.ExecuteScalar()

        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),ccc23e,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "left join ccc_file on pn = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0"
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity,ccc23e order by pn,max(azn01)"
        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),ccc23e,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "left join ccc_file on tc_prm01 = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04,ccc23e order by tc_prm01,max(azn01)"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 8) = oReader.Item(4)
                Ws.Cells(LineZ, 9) = oReader.Item(5)
                Ws.Cells(LineZ, 10) = oReader.Item(6)
                Dim SALESPRICE As Decimal = 0
                If IsDBNull(oReader.Item(7)) Then
                    Ws.Cells(LineZ, 11) = "总成本与总收入%"
                    oCommand2.CommandText = "select tc_prl03 * er * tc_prl04 / 100 from ( select rownum,tc_prl03,tc_prl06,er,tc_prl02,tc_prl04 from ( select tc_prl03,tc_prl06,tc_prl02,tc_prl04 from tc_prl_file where tc_prl01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prl02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                    oCommand2.CommandText += "union all "
                    oCommand2.CommandText += "select tc_prn03,tc_prn06,tc_prn02,tc_prn04 from tc_prn_file where tc_prn01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prn02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  ) "
                    oCommand2.CommandText += "left join exchangeratebyyear on tc_prl06 = currency and year1 = " & tYear & "order by tc_prl02 ) where rownum = 1"
                    SALESPRICE = oCommand2.ExecuteScalar()
                    Ws.Cells(LineZ, 12) = PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * PercentageofSales * SALESPRICE
                Else
                    If sMonth2 < 10 Then
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/0" & sMonth2
                    Else
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/" & sMonth2
                        End If
                    Ws.Cells(LineZ, 12) = oReader.Item(7)
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * oReader.Item(7)
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * oReader.Item(7)
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * oReader.Item(7)
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * oReader.Item(7)
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * oReader.Item(7)
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * oReader.Item(7)
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * oReader.Item(7)
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * oReader.Item(7)
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * oReader.Item(7)
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * oReader.Item(7)
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * oReader.Item(7)
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * oReader.Item(7)
                    End If
                Ws.Cells(LineZ, 25) = "=SUM(M" & LineZ & ":X" & LineZ & ")"

                LineZ += 1
                End While
            End If
            ' 加總
        Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 13))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 25)), Type:=xlFillDefault)

        oReader.Close()
            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ, 25))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            '第九頁
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(9)
        Ws.Name = "Cost Outsourced"
        Ws.Activate()
        AdjustExcelFormat4()
        oCommand.CommandText = "select nvl(round(sum(ccc61 * ccc23d * -1) /sum(ccc63),5),0) from ccc_file left join ima_file on ccc01 = ima01 where ccc02 || (case when ccc03 < 10 then '0' || ccc03 else to_char(ccc03) end) between '"
        oCommand.CommandText += mPeriod1 & "' and '" & mPeriod2 & "' and ccc63 <> 0 and ccc61 <> 0 and ccc62 <> 0  and ima06 = '103' "
        PercentageofSales = oCommand.ExecuteScalar()

        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),ccc23d,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "left join ccc_file on pn = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0"
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity,ccc23d order by pn,max(azn01)"
        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),ccc23d,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "left join ccc_file on tc_prm01 = ccc01 and ccc02 = " & sYear2 & " and ccc03 = " & sMonth2 & " and ccc23 > 0 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04,ccc23d order by tc_prm01,max(azn01)"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 8) = oReader.Item(4)
                Ws.Cells(LineZ, 9) = oReader.Item(5)
                Ws.Cells(LineZ, 10) = oReader.Item(6)
                Dim SALESPRICE As Decimal = 0
                If IsDBNull(oReader.Item(7)) Then
                    Ws.Cells(LineZ, 11) = "总成本与总收入%"
                    oCommand2.CommandText = "select tc_prl03 * er * tc_prl04 / 100 from ( select rownum,tc_prl03,tc_prl06,er,tc_prl02,tc_prl04 from ( select tc_prl03,tc_prl06,tc_prl02,tc_prl04 from tc_prl_file where tc_prl01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prl02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
                    oCommand2.CommandText += "union all "
                    oCommand2.CommandText += "select tc_prn03,tc_prn06,tc_prn02,tc_prn04 from tc_prn_file where tc_prn01 = '"
                    oCommand2.CommandText += oReader.Item(0) & "' and tc_prn02 > to_date('"
                    oCommand2.CommandText += Convert.ToDateTime(oReader.Item(6)).ToString("yyyy/MM/dd") & "','yyyy/mm/dd')  ) "
                    oCommand2.CommandText += "left join exchangeratebyyear on tc_prl06 = currency and year1 = " & tYear & "order by tc_prl02 ) where rownum = 1"
                    SALESPRICE = oCommand2.ExecuteScalar()
                    Ws.Cells(LineZ, 12) = PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * PercentageofSales * SALESPRICE
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * PercentageofSales * SALESPRICE
                Else
                    If sMonth2 < 10 Then
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/0" & sMonth2
                    Else
                        Ws.Cells(LineZ, 11) = "实际成本年月" & sYear2 & "/" & sMonth2
                        End If
                    Ws.Cells(LineZ, 12) = oReader.Item(7)
                    Ws.Cells(LineZ, 13) = oReader.Item(8) * oReader.Item(7)
                    Ws.Cells(LineZ, 14) = oReader.Item(9) * oReader.Item(7)
                    Ws.Cells(LineZ, 15) = oReader.Item(10) * oReader.Item(7)
                    Ws.Cells(LineZ, 16) = oReader.Item(11) * oReader.Item(7)
                    Ws.Cells(LineZ, 17) = oReader.Item(12) * oReader.Item(7)
                    Ws.Cells(LineZ, 18) = oReader.Item(13) * oReader.Item(7)
                    Ws.Cells(LineZ, 19) = oReader.Item(14) * oReader.Item(7)
                    Ws.Cells(LineZ, 20) = oReader.Item(15) * oReader.Item(7)
                    Ws.Cells(LineZ, 21) = oReader.Item(16) * oReader.Item(7)
                    Ws.Cells(LineZ, 22) = oReader.Item(17) * oReader.Item(7)
                    Ws.Cells(LineZ, 23) = oReader.Item(18) * oReader.Item(7)
                    Ws.Cells(LineZ, 24) = oReader.Item(19) * oReader.Item(7)
                    End If
                Ws.Cells(LineZ, 25) = "=SUM(M" & LineZ & ":X" & LineZ & ")"

                LineZ += 1
                End While
            End If
            ' 加總
        Ws.Cells(LineZ, 13) = "=SUM(M6:M" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 13))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 13), Ws.Cells(LineZ, 25)), Type:=xlFillDefault)

        oReader.Close()
            ' 劃線
        oRng = Ws.Range("B4", Ws.Cells(LineZ, 25))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "Currency：RMB"
        Ws.Cells(4, 2) = "Period Time:" & mPeriod1 & "-" & mPeriod2
        Ws.Cells(5, 2) = "ERP料号"
        Ws.Cells(5, 3) = "品名"
        Ws.Cells(5, 4) = "规格"
        Ws.Cells(5, 5) = "单位"
        Ws.Cells(5, 6) = "累计销售数量"
        Ws.Cells(5, 7) = "累计销售金额"
        Ws.Cells(5, 8) = "累计销售成本-合计"
        Ws.Cells(5, 9) = "累计销售成本-材料"
        Ws.Cells(5, 10) = "累计销售成本-人工"
        Ws.Cells(5, 11) = "累计销售成本-制费1"
        Ws.Cells(5, 12) = "累计销售成本-制费2"
        Ws.Cells(5, 13) = "累计销售成本-委外"
        Ws.Cells(6, 2) = "Part No."
        Ws.Cells(6, 3) = "Part Name"
        Ws.Cells(6, 4) = "Spec."
        Ws.Cells(6, 5) = "Unit"
        Ws.Cells(6, 6) = "Sales Qty"
        Ws.Cells(6, 7) = "Sales Amount"
        Ws.Cells(6, 8) = "COGS"
        Ws.Cells(6, 9) = "Cost-Material"
        Ws.Cells(6, 10) = "Cost-DL"
        Ws.Cells(6, 11) = "Cost-OVH 1"
        Ws.Cells(6, 12) = "Cost-OVH 2"
        Ws.Cells(6, 13) = "Cost Outsourced"
        oRng = Ws.Range("B2")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("F1", "M1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "

        LineZ = 7
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "Currency：RMB"
        Ws.Cells(4, 2) = "ERP料号"
        Ws.Cells(4, 3) = "品名"
        Ws.Cells(4, 4) = "规格"
        Ws.Cells(4, 5) = "单位"
        Ws.Cells(4, 6) = "项次"
        Ws.Cells(4, 7) = "原币"
        Ws.Cells(4, 8) = "原币售价"
        Ws.Cells(4, 9) = "东莞取价百分比"
        Ws.Cells(4, 10) = "汇率"
        Ws.Cells(4, 11) = "本币售价"
        Ws.Cells(4, 12) = "截止日期"
        Ws.Cells(5, 2) = "Part No."
        Ws.Cells(5, 3) = "Part Name"
        Ws.Cells(5, 4) = "Spec."
        Ws.Cells(5, 5) = "Unit"
        Ws.Cells(5, 6) = "Positi"
        Ws.Cells(5, 7) = "Currency"
        Ws.Cells(5, 8) = "Price"
        Ws.Cells(5, 9) = "TP %"
        Ws.Cells(5, 10) = "Exchange"
        Ws.Cells(5, 11) = "Price"
        Ws.Cells(5, 12) = "Closing Date "
        oRng = Ws.Range("B2")
        oRng.EntireColumn.NumberFormat = "@"
        oRng = Ws.Range("K2")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "

        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "Currency：RMB"
        Ws.Cells(4, 2) = "ERP料号"
        Ws.Cells(4, 3) = "品名"
        Ws.Cells(4, 4) = "规格"
        Ws.Cells(4, 5) = "单位"
        Ws.Cells(4, 6) = "年"
        Ws.Cells(4, 7) = "周"
        Ws.Cells(4, 8) = "交货日期"
        Ws.Cells(5, 2) = "Part No."
        Ws.Cells(5, 3) = "Part Name"
        Ws.Cells(5, 4) = "Spec."
        Ws.Cells(5, 5) = "Unit"
        Ws.Cells(5, 6) = "Year"
        Ws.Cells(5, 7) = "Week"
        Ws.Cells(5, 8) = "Delivery Date "
        oRng = Ws.Range("I1", "U1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "
        oRng = Ws.Range("I4", "T4")
        oRng.NumberFormat = "[$-en-US]mmm/yy;@"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(4, 8 + i) = tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(4, 21) = "合计"
        Ws.Cells(5, 21) = "Total"
        LineZ = 6
    End Sub
    Private Sub TransferDB()
        oCommand.CommandText = "DROP TABLE budget2020"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        oCommand.CommandText = "Create table Budget2020 (PN varchar2(40), Year1 number(5,0), Week1 number(5,0), Quantity number(18,3))"
        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

        For i As Int16 = 53 To 104 Step 1
            oCommand.CommandText = "insert into budget2020 select pn," & tYear & "," & i - 52 & ",sum(w" & i & ") from ship_temp WHERE w" & i & " <> 0 and etype in (1,2) "
            oCommand.CommandText += " group by pn"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        Next

    End Sub
    Private Sub AdjustExcelFormat4()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "Currency：RMB"
        Ws.Cells(4, 2) = "ERP料号"
        Ws.Cells(4, 3) = "品名"
        Ws.Cells(4, 4) = "规格"
        Ws.Cells(4, 5) = "单位"
        Ws.Cells(4, 6) = "销售区域"
        Ws.Cells(4, 7) = "销售客户"
        Ws.Cells(4, 8) = "年"
        Ws.Cells(4, 9) = "周"
        Ws.Cells(4, 10) = "交货日期"
        Ws.Cells(4, 11) = "备注"
        Ws.Cells(4, 12) = "单位成本"
        Ws.Cells(5, 2) = "Part No."
        Ws.Cells(5, 3) = "Part Name"
        Ws.Cells(5, 4) = "Spec."
        Ws.Cells(5, 5) = "Unit"
        Ws.Cells(5, 6) = "Area"
        Ws.Cells(5, 7) = "Customer"
        Ws.Cells(5, 8) = "Year"
        Ws.Cells(5, 9) = "Week"
        Ws.Cells(5, 10) = "Delivery Date"
        Ws.Cells(5, 11) = "Remark"
        Ws.Cells(5, 12) = "Cost per unit"
        oRng = Ws.Range("L1", "Y1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "

        oRng = Ws.Range("M4", "X4")
        oRng.NumberFormat = "[$-en-US]mmm/yy;@"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(4, 12 + i) = tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(4, 25) = "合计"
        Ws.Cells(5, 25) = "Total"
        LineZ = 6
    End Sub
End Class