Imports Microsoft.Office.Interop.Excel.XlFileFormat
Public Class Form190
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim pYear As String = String.Empty
    Dim tDate1 As Date
    Dim tDate2 As Date
    Dim tDate3 As Date
    Dim C1 As String = String.Empty
    Dim C2 As String = String.Empty
    Dim LineZ As Integer = 0

    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form190_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        pYear = NumericUpDown1.Value
        tDate1 = Convert.ToDateTime(pYear & "/01/01")
        tDate2 = tDate1.AddYears(1).AddDays(-1)
        tDate3 = Now.Date
        C1 = tDate3.AddMonths(-1).Year
        C2 = tDate3.AddMonths(-1).Month
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
                oCommand2.Connection = oConnection
                oCommand2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
       
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Price_Monitor"
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
        Ws.Activate()
        AdjustExcelFormat()
        
        oCommand.CommandText = "Select distinct pmm09,pmc03,pmn04,ima02,ima021,ima44,ima44_fac,ima25,ima08,ima06 ,(case when ima08 ='S' then stb09a else (stb07 + stb08 + stb09 + stb09a) end) as t1 "
        oCommand.CommandText += "from pmm_file left join pmn_file on pmm01 = pmn01 left join pmc_file on pmm09 = pmc01 left join ima_file on pmn04 = ima01 left join stb_file on stb01 = pmn04 and stb02 = " & C1 & " and stb03 = " & C2
        oCommand.CommandText += " where pmm18 = 'Y' and pmm04 between to_date('"
        oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmn53 > 0 order by pmn04"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()

                Ws.Cells(LineZ, 1) = oReader.Item("pmm09")
                Ws.Cells(LineZ, 2) = oReader.Item("pmn04")
                Ws.Cells(LineZ, 3) = oReader.Item("ima02")
                Ws.Cells(LineZ, 4) = oReader.Item("ima021")
                Ws.Cells(LineZ, 5) = oReader.Item("ima44")
                Ws.Cells(LineZ, 6) = oReader.Item("ima44_fac")
                Ws.Cells(LineZ, 7) = oReader.Item("ima25")
                Ws.Cells(LineZ, 8) = oReader.Item("ima08")
                Ws.Cells(LineZ, 9) = oReader.Item("ima06")
                Ws.Cells(LineZ, 10) = oReader.Item("t1")
                oCommand2.CommandText = "Select pmm01,pmn31t, pmm22, pmm42, Round((pmn44 / pmn09),2) as t1  from ( Select * from pmm_file left join pmn_file on pmm01 = pmn01 "
                oCommand2.CommandText += "where pmm09 = '" & oReader.Item("pmm09") & "' and pmn04 = '" & oReader.Item("pmn04") & "' and pmm04 between to_date('"
                oCommand2.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand2.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmn53 >0 order by pmm04 desc ) where rownum = 1"
                oReader2 = oCommand2.ExecuteReader()
                If oReader2.HasRows() Then
                    While oReader2.Read()
                        Ws.Cells(LineZ, 11) = oReader2.Item("pmm01")
                        Ws.Cells(LineZ, 12) = oReader2.Item("pmn31t")
                        Ws.Cells(LineZ, 13) = oReader2.Item("pmm22")
                        Ws.Cells(LineZ, 14) = oReader2.Item("pmm42")
                        Ws.Cells(LineZ, 16) = oReader2.Item("t1")
                    End While
                End If
                oReader2.Close()
                oCommand2.CommandText = "Select nvl(Round(avg(pmn31 * pmm42 / pmn09),2),0) from pmm_file left join pmn_file on pmm01 = pmn01 where pmm09 = '" & oReader.Item("pmm09") & "' and pmn04 = '"
                oCommand2.CommandText += oReader.Item("pmn04") & "' and pmm04 between to_date('"
                oCommand2.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
                oCommand2.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and pmm18 = 'Y' and pmn53 >0 "
                Dim XXA As Decimal = oCommand2.ExecuteScalar()
                Ws.Cells(LineZ, 15) = XXA
                ' 入庫量
                oCommand2.CommandText = "Select nvl(Round(sum(rvv17 * rvv35_fac),2),0) from rvu_file left join rvv_file on rvu01 = rvv01 where rvu04 = '" & oReader.Item("pmm09") & "' and rvv31 = '"
                oCommand2.CommandText += oReader.Item("pmn04") & "' and rvuconf = 'Y' and year(rvu03) = " & tDate3.Year
                Dim XXB As Decimal = oCommand2.ExecuteScalar
                Ws.Cells(LineZ, 18) = XXB
                ' 最後一次入庫單價
                oCommand2.CommandText = "Select * from ( Select nvl(Round((rvv38  * pmm42 / rvv35_fac),2),0) as t1 from rvu_file left join rvv_file on rvu01 = rvv01 left join pmm_file on rvv36 = pmm01 "
                oCommand2.CommandText += "where rvu04 = '" & oReader.Item("pmm09") & "' and rvv31 = '" & oReader.Item("pmn04") & "' and rvuconf = 'Y' and year(rvu03) = " & tDate3.Year & " order by rvu03 desc ) where rownum = 1"
                Dim XXC As Decimal = oCommand2.ExecuteScalar()
                Ws.Cells(LineZ, 17) = XXC

                Ws.Cells(LineZ, 19) = "=Q" & LineZ & "-J" & LineZ
                Ws.Cells(LineZ, 20) = "=S" & LineZ & "*R" & LineZ
                Ws.Cells(LineZ, 21) = "=T" & LineZ & "/(J" & LineZ & "*R" & LineZ & ")"
                Ws.Cells(LineZ, 22) = "=Q" & LineZ & "-O" & LineZ
                Ws.Cells(LineZ, 23) = "=V" & LineZ & "*R" & LineZ
                Ws.Cells(LineZ, 24) = "=W" & LineZ & "/(O" & LineZ & "*R" & LineZ & ")"
                Ws.Cells(LineZ, 25) = "=Q" & LineZ & "-P" & LineZ
                Ws.Cells(LineZ, 26) = "=Y" & LineZ & "*R" & LineZ
                Ws.Cells(LineZ, 27) = "=Z" & LineZ & "/(P" & LineZ & "*R" & LineZ & ")"
                Ws.Cells(LineZ, 28) = oReader.Item("pmc03")
                LineZ += 1
                Label2.Text = LineZ
            End While
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.ColumnWidth = 21.11
        Ws.Cells(1, 1) = "廠商"
        Ws.Cells(1, 2) = "料件編號"
        Ws.Cells(1, 3) = "品名"
        Ws.Cells(1, 4) = "規格"
        Ws.Cells(1, 5) = "採購單位"
        Ws.Cells(1, 6) = "換算率"
        Ws.Cells(1, 7) = "庫存單位"
        Ws.Cells(1, 8) = "來源碼"
        Ws.Cells(1, 9) = "分群碼"
        Ws.Cells(1, 10) = C1 & "標準單價 本币"
        Ws.Cells(1, 11) = pYear & "年最后一张订单採購單號"
        Ws.Cells(1, 12) = pYear & "年最后一张订单的含税价(原幣)"
        Ws.Cells(1, 13) = "币别"
        Ws.Cells(1, 14) = "匯率"
        Ws.Cells(1, 15) = pYear & "年年度平均采购单价(不含税价本币)"
        Ws.Cells(1, 16) = pYear & "年最后一张订单的不含税价(本幣)"
        Ws.Cells(1, 17) = tDate3.Year & "未稅單價(本幣不含税)"
        Ws.Cells(1, 18) = tDate3.Year & "年入库量"
        Ws.Cells(1, 19) = tDate3.Year & "未税單價与" & tDate3.Year & "标准价格单价差"
        Ws.Cells(1, 20) = "金額差1"
        Ws.Cells(1, 21) = "降比1"
        Ws.Cells(1, 22) = tDate3.Year & "未税单价与" & pYear & "年平均采购价格差"
        Ws.Cells(1, 23) = "金額差2"
        Ws.Cells(1, 24) = "降比2"
        Ws.Cells(1, 25) = tDate3.Year & "未税单价与" & pYear & "年最后一张订单的不含税价差"
        Ws.Cells(1, 26) = "金額差3"
        Ws.Cells(1, 27) = "降比3"
        Ws.Cells(1, 28) = "供应商简称"
        oRng = Ws.Range("B1", "B1")
        oRng.EntireColumn.NumberFormat = "@"

        LineZ = 2
    End Sub
End Class