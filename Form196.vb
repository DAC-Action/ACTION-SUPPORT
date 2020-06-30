Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form196
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim lYear As Int16 = 0
    Dim DStartN As Date
    Dim DstartE As Date
    Dim DstartE1 As Date
    Dim LineZ As Integer = 0
    Dim LineX As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        
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
        lYear = NumericUpDown1.Value
        DStartN = Convert.ToDateTime(lYear & "/01/01")
        DstartE = DStartN.AddYears(5).AddDays(-1)
        DstartE1 = DStartN.AddYears(3).AddDays(-1)
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form196_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Purchase_Report"
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
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub ExportToExcel()

        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat(lYear)

        oCommand.CommandText = "Select rvu04,rvu05,pmccrat,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t1+t2+t3+t4+t5) as t6,"
        oCommand.CommandText += "pmc081,pma02,pmc091,pmcud03,pmc10,pmc12 from (select rvu04,rvu05,pmccrat,pmc081, pma02,pmc091,pmcud03,pmc10,pmc12,"
        oCommand.CommandText += "(case when year(rvu03) = " & lYear & " then rvv39 * pmm42 else 0 end) as t1,(case when year(rvu03) = " & lYear + 1 & " then rvv39 * pmm42 else 0 end) as t2,"
        oCommand.CommandText += "(case when year(rvu03) = " & lYear + 2 & " then rvv39 * pmm42 else 0 end) as t3,(case when year(rvu03) = " & lYear + 3 & " then rvv39 * pmm42 else 0 end) as t4,"
        oCommand.CommandText += "(case when year(rvu03) = " & lYear + 4 & " then rvv39 * pmm42 else 0 end) as t5 from rvu_file left join rvv_file on rvu01 =rvv01 "
        oCommand.CommandText += "left join pmc_file on rvu04 = pmc01 left join pma_file on pmc17 = pma01 left join pmm_file on rvv36 = pmm01 where rvu03 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and rvuconf = 'Y' "
        oCommand.CommandText += ") group by rvu04,rvu05,pmccrat,pmc081,pma02,pmc091,pmcud03,pmc10,pmc12"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
                Label2.Text = LineZ
                Label2.Refresh()
            End While
        End If
        oReader.Close()

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat1(lYear)

        oCommand.CommandText = "Select rvv31,c1,sum(t1) as t1,sum(t1a) as t1a,sum(t2) as t2,sum(t2a) as t2a,sum(t3) as t3,sum(t3a) as t3a,"
        oCommand.CommandText += "sum(t1+t2+t3) as t4,sum(t1a+t2a+t3a) as t4a,ima02,ima021,ima901,ima25,rvv35_fac,rvv35,ima48,ima46,ima45 from ( "
        oCommand.CommandText += "select rvv31,(rvu04 || rvu05) as c1,ima02,ima021,ima901,ima25,rvv35_fac,rvv35,ima48,ima46,ima45,"
        oCommand.CommandText += "(case when year(rvu03) = 2016 then rvv17 else 0 end) as t1,(case when year(rvu03) = 2017 then rvv17 else 0 end) as t2,"
        oCommand.CommandText += "(case when year(rvu03) = 2018 then rvv17 else 0 end) as t3,(case when year(rvu03) = 2016 then rvv39 * pmm42 else 0 end) as t1a,"
        oCommand.CommandText += "(case when year(rvu03) = 2017 then rvv39 * pmm42 else 0 end) as t2a,(case when year(rvu03) = 2018 then rvv39 * pmm42 else 0 end) as t3a "
        oCommand.CommandText += "from rvu_file left join rvv_file on rvu01 =rvv01 left join ima_file on rvv31 = ima01 left join pmm_file on rvv36 = pmm01 where rvu03 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('" & DstartE1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and rvuconf = 'Y' "
        oCommand.CommandText += ") group by rvv31,c1,ima02,ima021,ima901,ima25,rvv35_fac,rvv35,ima48,ima46,ima45"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                GetCustomer(oReader.Item(0))
                GetMasterItem(oReader.Item(0))
                GetSector(oReader.Item(0))
                LineZ += 1
                Label2.Text = LineZ
                Label2.Refresh()
            End While
        End If
        oReader.Close()

    End Sub
    Private Sub AdjustExcelFormat(ByVal sYear As Int16)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "供应商金额汇总表"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "供应商代码"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 2) = "供应商简称"
        Ws.Cells(1, 3) = "录入ERP时间"
        Ws.Cells(1, 4) = sYear & "未税金额RMB"
        Ws.Cells(1, 5) = sYear + 1 & "未税金额RMB"
        Ws.Cells(1, 6) = sYear + 2 & "未税金额RMB"
        Ws.Cells(1, 7) = sYear + 3 & "未税金额RMB"
        Ws.Cells(1, 8) = sYear + 4 & "未税金额RMB"
        Ws.Cells(1, 9) = sYear & "-" & sYear + 4 & "金额汇总"
        Ws.Cells(1, 10) = "供应商名称"
        Ws.Cells(1, 11) = "付款条件"
        Ws.Cells(1, 12) = "公司地址"
        Ws.Cells(1, 13) = "联络人"
        Ws.Cells(1, 14) = "联系电话"
        Ws.Cells(1, 15) = "电子邮箱"
        LineZ = 2
    End Sub

    Private Sub AdjustExcelFormat1(ByVal sYear As Int16)
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "物料金额汇总表"
        Ws.Columns.EntireColumn.ColumnWidth = 20
        Ws.Cells(1, 1) = "元件料号"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormat = "@"
        Ws.Cells(1, 2) = "厂商简称"
        Ws.Cells(1, 3) = sYear & "入庫量（采购单位）"
        Ws.Cells(1, 4) = sYear & "2017未税金额(本幣)"
        Ws.Cells(1, 5) = sYear + 1 & "入庫量（采购单位）"
        Ws.Cells(1, 6) = sYear + 1 & "2017未税金额(本幣)"
        Ws.Cells(1, 7) = sYear + 2 & "入庫量（采购单位）"
        Ws.Cells(1, 8) = sYear + 2 & "2017未税金额(本幣)"
        Ws.Cells(1, 9) = "入库数量汇总（采购单位）"
        Ws.Cells(1, 10) = "未税金额汇总(本幣)"
        Ws.Cells(1, 11) = "品名"
        Ws.Cells(1, 12) = "规格"
        Ws.Cells(1, 13) = "生效日期"
        Ws.Cells(1, 14) = "库存单位"
        Ws.Cells(1, 15) = "换算率"
        Ws.Cells(1, 16) = "采购单位"
        Ws.Cells(1, 17) = "前置期"
        Ws.Cells(1, 18) = "MOQ"
        Ws.Cells(1, 19) = "MPQ"
        Ws.Cells(1, 20) = "产品客户代码"
        Ws.Cells(1, 21) = "主件简号"
        Ws.Cells(1, 22) = "应用生产工序"
        LineZ = 2
    End Sub
    Private Sub GetCustomer(ByVal bmb03 As String)
        Dim CS1 As String = String.Empty
        oCommand2.CommandText = "select distinct substr(bmb01,4,2) as c1 from bmb_file,bma_file where bmb01 = bma01 and  bmb03 = '" & bmb03 & "' and bmb05 is null and bma10 = '2' and bmaacti = 'Y'"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                CS1 += oReader2.Item("c1") & "|"
            End While
            CS1 = CS1.Remove(CS1.Length() - 1, 1)
            Ws.Cells(LineZ, 20) = CS1
        End If
        oReader2.Close()
    End Sub
    Private Sub GetMasterItem(ByVal bmb03 As String)
        Dim CS1 As String = String.Empty
        oCommand2.CommandText = "select distinct substr(bmb01,4,6) as c1 from bmb_file,bma_file where bmb01 = bma01 and  bmb03 = '" & bmb03 & "' and bmb05 is null and bma10 = '2' and bmaacti = 'Y'"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                CS1 += oReader2.Item("c1") & "|"
            End While
            CS1 = CS1.Remove(CS1.Length() - 1, 1)
            Ws.Cells(LineZ, 21) = CS1
        End If
        oReader2.Close()
    End Sub
    Private Sub GetSector(ByVal bmb03 As String)
        Dim CS1 As String = String.Empty
        oCommand2.CommandText = "select distinct  (case when substr(bmb01,length(bmb01),1) = 'A' then substr(bmb01,length(bmb01)-2,3) else "
        oCommand2.CommandText += "substr(bmb01,length(bmb01)-1,2) end) as c1 from bmb_file,bma_file where bmb01 = bma01 and bma10 = '2' and  bmb03 = '"
        oCommand2.CommandText += bmb03 & "' and bmaacti = 'Y'"
        oReader2 = oCommand2.ExecuteReader()
        If oReader2.HasRows() Then
            While oReader2.Read()
                Select Case oReader2.Item("c1")
                    Case "31"
                        CS1 += "裁纱" & "|"
                    Case "32"
                        CS1 += "预型" & "|"
                    Case "32A"
                        CS1 += "二次预型" & "|"
                    Case "35"
                        CS1 += "成型" & "|"
                    Case "35A"
                        CS1 += "二次成型" & "|"
                    Case "36"
                        CS1 += "CNC" & "|"
                    Case "36A"
                        CS1 += "二次CNC" & "|"
                    Case "61"
                        CS1 += "补土" & "|"
                    Case "61A"
                        CS1 += "二次补土" & "|"
                    Case "63"
                        CS1 += "涂装" & "|"
                    Case "63A"
                        CS1 += "二次涂装" & "|"
                    Case "64"
                        CS1 += "胶合" & "|"
                    Case "64A"
                        CS1 += "二次胶合" & "|"
                    Case "65"
                        CS1 += "抛光" & "|"
                    Case "65A"
                        CS1 += "二次抛光" & "|"
                    Case "66"
                        CS1 += "包装" & "|"
                End Select
            End While
            If CS1.Length > 0 Then
                CS1 = CS1.Remove(CS1.Length() - 1, 1)
            End If
            Ws.Cells(LineZ, 22) = CS1
        End If
        oReader2.Close()
    End Sub
End Class