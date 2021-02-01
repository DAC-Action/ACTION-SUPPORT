Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form30
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader99 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim kConnection As New SqlClient.SqlConnection
    Dim kCommander As New SqlClient.SqlCommand
    Dim kReader As SqlClient.SqlDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim DStartN As Date
    Dim DstartE As Date
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim CYear As String = String.Empty
    Dim CMonth As String = String.Empty
    Dim g_pja01 As String = String.Empty
    Dim LineZ As Integer = 0
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
        kConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        If kConnection.State <> ConnectionState.Open Then
            kConnection.Open()
            kCommander.Connection = kConnection
            kCommander.CommandType = CommandType.Text
        End If
        TYear = Strings.Left(TextBox1.Text, 4)
        TMonth = Strings.Right(TextBox1.Text, 2)
        DStartN = DateTimePicker1.Value
        DstartE = DateTimePicker2.Value
        'DStartN = Convert.ToDateTime(TYear & "/" & TMonth & "/01")
        'DstartE = DStartN.AddMonths(1).AddDays(-1)
        CYear = Strings.Left(TextBox2.Text, 4)
        CMonth = Strings.Right(TextBox2.Text, 2)
        'add by cloud 20170922
        g_pja01 = TextBox3.Text
        'ExportToExcel()
        'SaveExcel()
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form30_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If Now.Month < 10 Then
            TextBox1.Text = Now.Year & "0" & Now.Month
        Else
            TextBox1.Text = Now.Year & Now.Month
        End If
        TextBox2.Text = TextBox1.Text
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Project_Report"
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
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "monthly sum"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 30
        oRng = Ws.Range("A1", "D1")
        oRng.Merge()
        oRng = Ws.Range("A2", "D2")
        oRng.Merge()
        Ws.Cells(1, 1) = "Dongguan Action Composites LTD Co."
        Ws.Cells(2, 1) = "Sum cost by Project"
        Ws.Cells(3, 1) = "币别（Currency)：RMB"
        Ws.Cells(3, 2) = "Month:"
        oRng = Ws.Range("A4", "A5")
        oRng.Merge()
        oRng = Ws.Range("B4", "B5")
        oRng.Merge()
        oRng = Ws.Range("C4", "C5")
        oRng.Merge()
        oRng = Ws.Range("D4", "D5")
        oRng.Merge()
        oRng = Ws.Range("E4", "E5")
        oRng.Merge()
        'oRng = Ws.Range("F4", "H4")
        'oRng.Merge()
        'oRng = Ws.Range("I4", "I5")
        'oRng.Merge()
        oRng = Ws.Range("K4", "K5")
        oRng.Merge()
        oRng = Ws.Range("L4", "L5")
        oRng.Merge()
        oRng = Ws.Range("M4", "M5")
        oRng.Merge()
        Ws.Cells(4, 1) = "项目编号（Project Code）"
        Ws.Cells(4, 2) = "项目名称（Project Name）"
        Ws.Cells(4, 3) = "项目类别(type)"
        Ws.Cells(4, 4) = "项目立项日期（Project Start Date）"
        Ws.Cells(4, 5) = "项目负责人（Project Leader） "
        Ws.Cells(4, 6) = "材料（Material）"
        Ws.Cells(4, 7) = "RD人工（RD labour cost）"
        Ws.Cells(4, 8) = "CMM人工（CMM labour cost）"
        Ws.Cells(4, 9) = "PL人工（PL labour cost）"
        Ws.Cells(4, 10) = "QE人工（QE labour cost）"
        Ws.Cells(4, 11) = "厂内付费模具费用(DAC mold)"
        Ws.Cells(4, 12) = "客户付费模具费用(mold paid by customer)"
        Ws.Cells(4, 13) = "Total"
        Ws.Cells(5, 6) = "金额（amount）"
        Ws.Cells(5, 7) = "金额（amount）"
        Ws.Cells(5, 8) = "金额（amount）"
        Ws.Cells(5, 9) = "金额（amount）"
        Ws.Cells(5, 10) = "金额（amount）"
        'Ws.Cells(5, 6) = ""
        'Ws.Cells(5, 7) = ""
        'Ws.Cells(5, 6) = "工时（labor hour）"
        'Ws.Cells(5, 7) = "工费率（LB rate）"
        'Ws.Cells(5, 8) = "金额（amount）"
        LineZ = 6
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        xWorkBook.Sheets.Add()
        Ws = xWorkBook.Sheets(1)

        Ws.Activate()
        AdjustExcelFormat1()
        'oCommand.CommandText = "select nvl(round(sum(tlf10*tlf12*(stb07+stb08+stb09)),2),0) as t1 ,tlf20,pja01,pja02,pja05,pja08 from pja_file "
        oCommand.CommandText = "select nvl(round(sum(tlf10*tlf12*(ccc23)),2),0) as t1 ,tlf20,pja01,pja02,pja05,pja08, pjq02 from pja_file "
        oCommand.CommandText += "LEFT JOIN tlf_file on pja01 = tlf20 and tlf20 is not null and tlf20 <> ' ' and tlf06 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tlf907 = -1 and tlf13 like 'aimt3%' "
        oCommand.CommandText += "left join pjq_file on pja07 = pjq01 "
        oCommand.CommandText += "left join ima_file on tlf01 = ima01 and ima06 <> '105' "
        'oCommand.CommandText += "left join stb_file on stb01 = tlf01 and stb01 = ima01 and stb02 = " & CYear & " and stb03 = " & CMonth
        oCommand.CommandText += "left join ccc_file on ccc01 = tlf01 and ccc01 = ima01 and ccc02 = " & CYear & " and ccc03 = " & CMonth
        If Not String.IsNullOrEmpty(g_pja01) Then
            oCommand.CommandText += " WHERE pja01 LIKE '" & g_pja01 & "%' "
        End If
        oCommand.CommandText += " group by tlf20,pja01,pja02,pja05,pja08, pjq02 order by tlf20 "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item("pja01")
                Ws.Cells(LineZ, 2) = oReader.Item("pja02")
                Ws.Cells(LineZ, 3) = oReader.Item("pjq02")
                Ws.Cells(LineZ, 4) = oReader.Item("pja05")
                Ws.Cells(LineZ, 5) = oReader.Item("pja08")
                Ws.Cells(LineZ, 6) = oReader.Item("t1")
                'Dim MoldFee As Decimal = GetMoldFee(oReader.Item("pja01"))
                'Ws.Cells(LineZ, 9) = MoldFee
                'Ws.Cells(LineZ, 6) = GetLaborHour(oReader.Item("pja01"))
                GetLaborHour(oReader.Item("pja01"))
                GetMoldFee(oReader.Item("pja01"))
                'Ws.Cells(LineZ, 7) = 35
                'Ws.Cells(LineZ, 8) = "=F" & LineZ & "*G" & LineZ
                Ws.Cells(LineZ, 13) = "=F" & LineZ & "+G" & LineZ & "+H" & LineZ & "+I" & LineZ & "+J" & LineZ & "+L" & LineZ
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 第二頁 20170909
        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat2()
        oCommand.CommandText = "select tlf19,gem02,tlf11,tlf06,tlf905,tlf01,ima02,ima06,imz02,imz39,aag02,ima021,tlf902,imd02,ccc23,tlf10*tlf12,(tlf10*tlf12*ccc23),ina07,inbud03,ima11,azf03,tlf20 from pja_file "
        oCommand.CommandText += "left join tlf_file on pja01 = tlf20 left join gem_file on tlf19 = gem01 left join ima_file on tlf01 = ima01 left join imz_file on ima06 = imz01 left join aag_file on imz39 = aag01 "
        oCommand.CommandText += "left join imd_file on tlf902 = imd01 left join ina_file on tlf905 = ina01 left join inb_file on tlf905 = inb01 and ina01 = inb01 and tlf906 = inb03 left join ccc_File on tlf01 = ccc01 and ccc02 = " & CYear & " and ccc03 = " & CMonth
        oCommand.CommandText += " left join azf_file on ima11 = azf01 and azf02 = 'F' where tlf20 is not null and tlf20 <> ' ' and tlf06 between to_date('"
        oCommand.CommandText += DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ima06 <> '105' and tlf907 = -1 and tlf13 like 'aimt3%' "
        If Not String.IsNullOrEmpty(g_pja01) Then
            oCommand.CommandText += " AND pja01 LIKE '" & g_pja01 & "%' "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
            End While
        End If
        oReader.Close()

        ' 第三頁 20170913
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        AdjustExcelFormat3()
        'kCommander.CommandText = "select  edate,euser,eusername,EProject ,ehour from ProjectHR WHERE eDate between '" & DStartN.ToString("yyyy/MM/dd") & "' AND '"
        kCommander.CommandText = "select  edate,euser,eusername,EProject ,ehour, EDepartNo, isnull(ProjectHourCost.HourCost,0) as HourCost  from ProjectHR left join ProjectHourCost on SUBSTRING(EDepartNo , 1, 2) = DepartmentCode  WHERE eDate between '" & DStartN.ToString("yyyy/MM/dd") & "' AND '"
        kCommander.CommandText += DstartE.ToString("yyyy/MM/dd") & "' "
        If Not String.IsNullOrEmpty(g_pja01) Then
            oCommand.CommandText += " and Eproject LIKE '" & g_pja01 & "%' "
        End If
        kReader = kCommander.ExecuteReader()
        If kReader.HasRows() Then
            While kReader.Read()
                Dim EB As Date = kReader.Item("edate")
                Ws.Cells(LineZ, 1) = kReader.Item("edate")
                Ws.Cells(LineZ, 2) = EB.ToString("yyyy-MM")
                Ws.Cells(LineZ, 3) = kReader.Item("euser")
                Ws.Cells(LineZ, 4) = kReader.Item("eusername")
                Ws.Cells(LineZ, 5) = kReader.Item("EProject")
                Ws.Cells(LineZ, 6) = GetProjectName(kReader.Item("EProject"))
                Ws.Cells(LineZ, 7) = kReader.Item("ehour")
                Dim l_hourcost As Decimal = 0
                l_hourcost = kReader.Item("HourCost")
                Ws.Cells(LineZ, 8) = l_hourcost
                Ws.Cells(LineZ, 9) = kReader.Item("ehour") * l_hourcost
                Ws.Cells(LineZ, 10) = kReader.Item("EDepartNo")
                LineZ += 1
            End While
        End If
        kReader.Close()

        ' 第四頁 20170913
        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        AdjustExcelFormat4()
        oCommand.CommandText = "select tlf19,pmc03,rva06,rva01,tlf905,rvv36,tlf01,ima02,ima021,tlf11,rvb08,tlf10,rvv38t,pmm22,pmm42,rvv39 * pmm42,(rvv39t - rvv39)*pmm42,rvv39t * pmm42,rvv39,(rvv39t - rvv39),rvv39t,tlf20, rvv37 from TLF_FILE "
        oCommand.CommandText += "left join rvv_file on tlf905=rvv01 and tlf906 = rvv02 left join rvb_file on rvv04 = rvb01 and rvv05  = rvb02 left join rva_file on rvb01 = rva01 left join pmm_file on rvv36 = pmm01 "
        oCommand.CommandText += "left join pmc_file on tlf19 = pmc01 left join ima_file on tlf01 = ima01 where tlf13 = 'apmt150' and tlf01 like '7%' and tlf20 is not null and tlf20 <> ' ' and tlf06 between to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(g_pja01) Then
            oCommand.CommandText += " and tlf20 LIKE '" & g_pja01 & "%' "
        End If
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Ws.Cells(LineZ, 1) = oReader.Item(0) & " " & oReader.Item(1)
                For i As Int16 = 2 To oReader.FieldCount - 2 Step 1
                    Ws.Cells(LineZ, i) = oReader.Item(i)
                Next
                oCommand2.CommandText = "Select nvl(Pmnud04,'') from pmn_file Where pmn01 = '" & oReader.Item("rvv36") & "' and pmn02 = " & oReader.Item("rvv37")
                Dim l_pmnud04 As String = String.Empty
                l_pmnud04 = oCommand2.ExecuteScalar()
                Select Case l_pmnud04
                    Case 1
                        Ws.Cells(LineZ, 22) = "1.客户付费"
                    Case 2
                        Ws.Cells(LineZ, 22) = "2:厂内付费"
                    Case 3
                        Ws.Cells(LineZ, 22) = "3:客户与厂内分摊"
                    Case Else
                        Ws.Cells(LineZ, 22) = l_pmnud04
                End Select
                LineZ += 1
            End While
        End If
        oReader.Close()
    End Sub
    'Private Function GetMoldFee(ByVal pja01 As String)
    '    Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
    '    oCommander99.Connection = oConnection
    '    oCommander99.CommandType = CommandType.Text
    '    oCommander99.CommandText = "select nvl(sum(rvv39 * pmm42),0) from TLF_FILE left join rvv_file on tlf905=rvv01 and tlf906 = rvv02 left join pmm_file on rvv36 = pmm01 "
    '    oCommander99.CommandText += "where tlf13 = 'apmt150' and tlf01 like '7%' and tlf20 is not null and tlf20 <> ' ' and tlf20 = '"
    '    oCommander99.CommandText += pja01 & "' and tlf06 between to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
    '    oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
    '    Dim MF As Decimal = oCommander99.ExecuteScalar()
    '    Return MF
    'End Function
    'Private Function GetLaborHour(ByVal pja01 As String)
    '    kCommander.CommandText = "select isnull(sum(eHour),0) from ProjectHR where eproject = '"
    '    kCommander.CommandText += pja01 & "' and eDate between '" & DStartN.ToString("yyyy/MM/dd") & "' AND '"
    '    kCommander.CommandText += DstartE.ToString("yyyy/MM/dd") & "' "
    '    Dim TH As Decimal = kCommander.ExecuteScalar()
    '    'If IsDBNull(TH) Then
    '    ' TH = 0
    '    'End If
    '    'kConnection.Close()
    '    Return TH
    'End Function
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "材料成本"
        Ws.Columns.EntireColumn.ColumnWidth = 13.44
        Ws.Cells(1, 1) = "部门编号"
        Ws.Cells(1, 2) = "部门名称"
        Ws.Cells(1, 3) = "单位"
        Ws.Cells(1, 4) = "单据日期"
        Ws.Cells(1, 5) = "单据编号"
        Ws.Cells(1, 6) = "料件编号"
        Ws.Cells(1, 7) = "品名"
        Ws.Cells(1, 8) = "分群码"
        Ws.Cells(1, 9) = "说明"
        Ws.Cells(1, 10) = "料件所属会计科目"
        Ws.Cells(1, 11) = "存货科目名称"
        Ws.Cells(1, 12) = "规格"
        Ws.Cells(1, 13) = "仓库"
        Ws.Cells(1, 14) = "仓库名称"
        Ws.Cells(1, 15) = "单价"
        Ws.Cells(1, 16) = "数量"
        Ws.Cells(1, 17) = "总金额"
        Ws.Cells(1, 18) = "单头备注"
        Ws.Cells(1, 19) = "单身备注"
        Ws.Cells(1, 20) = "其他分群码 三"
        Ws.Cells(1, 21) = "说明内容"
        Ws.Cells(1, 22) = "项目号码"

        LineZ = 2
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "人工成本"
        Ws.Columns.EntireColumn.ColumnWidth = 13.44
        Ws.Cells(1, 1) = "日期"
        Ws.Cells(1, 2) = "月份"
        Ws.Cells(1, 3) = "人员编号"
        Ws.Cells(1, 4) = "人员姓名"
        Ws.Cells(1, 5) = "專案編號"
        Ws.Cells(1, 6) = "专案名称"
        Ws.Cells(1, 7) = "工时"
        Ws.Cells(1, 8) = "平均人工"
        Ws.Cells(1, 9) = "人工成本"
        Ws.Cells(1, 10) = "部门"

        LineZ = 2
    End Sub
    Private Function GetProjectName(ByVal pja01 As String)
        oCommand2.CommandText = "select pja02 from pja_File where pja01 = '" & pja01 & "'"
        Dim PN As String = oCommand2.ExecuteScalar()
        Return PN
    End Function
    Private Sub AdjustExcelFormat4()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "模具费成本"
        Ws.Columns.EntireColumn.ColumnWidth = 13.44
        Ws.Cells(1, 1) = "廠商"
        Ws.Cells(1, 2) = "收貨日期"
        Ws.Cells(1, 3) = "收貨單號"
        Ws.Cells(1, 4) = "入庫單號"
        Ws.Cells(1, 5) = "採購單號"
        Ws.Cells(1, 6) = "料件編號"
        Ws.Cells(1, 7) = "品名"
        Ws.Cells(1, 8) = "規格"
        Ws.Cells(1, 9) = "單位"
        Ws.Cells(1, 10) = "收貨量"
        Ws.Cells(1, 11) = "入庫量"
        Ws.Cells(1, 12) = "含稅單價"
        Ws.Cells(1, 13) = "币别"
        Ws.Cells(1, 14) = "匯率"
        Ws.Cells(1, 15) = "未税金额(本幣)"
        Ws.Cells(1, 16) = "税额(本幣)"
        Ws.Cells(1, 17) = "含稅金額(本幣)"
        Ws.Cells(1, 18) = "未税金额(原幣)"
        Ws.Cells(1, 19) = "税额(原幣)"
        Ws.Cells(1, 20) = "含稅金額(原幣)"
        Ws.Cells(1, 21) = "项目编号"
        Ws.Cells(1, 22) = "是否客户付费"
        LineZ = 2
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
    Private Sub GetLaborHour(ByVal pja01 As String)
        kCommander.CommandText = "Select sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4 from ( "
        kCommander.CommandText += "select (case when SUBSTRING(EDepartNo , 1, 2) = 'RD' then sum(EHour * HourCost ) else 0 end) as t1, "
        kCommander.CommandText += "(case when SUBSTRING(EDepartNo , 1, 2) = 'QC' then sum(EHour * HourCost ) else 0 end) as t2, "
        kCommander.CommandText += "(case when SUBSTRING(EDepartNo , 1, 2) = 'PL' then sum(EHour * HourCost ) else 0 end) as t3, "
        kCommander.CommandText += "(case when SUBSTRING(EDepartNo , 1, 2) = 'QE' then sum(EHour * HourCost ) else 0 end) as t4 "
        kCommander.CommandText += "from ProjectHR left join ProjectHourCost on SUBSTRING(EDepartNo, 1, 2) = ProjectHourCost.DepartmentCode where eproject = '"
        kCommander.CommandText += pja01 & "' and eDate between '" & DStartN.ToString("yyyy/MM/dd") & "' AND '"
        kCommander.CommandText += DstartE.ToString("yyyy/MM/dd") & "' group by SUBSTRING( EDepartNo , 1, 2) ) as AD"
        kReader = kCommander.ExecuteReader()
        If kReader.HasRows() Then
            While kReader.Read()
                Ws.Cells(LineZ, 7) = kReader.Item("t1")
                Ws.Cells(LineZ, 8) = kReader.Item("t2")
                Ws.Cells(LineZ, 9) = kReader.Item("t3")
                Ws.Cells(LineZ, 10) = kReader.Item("t4")
            End While
        End If
        kReader.Close()
    End Sub
    Private Sub GetMoldFee(ByVal pja01 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection
        oCommander99.CommandType = CommandType.Text
        oCommander99.CommandText = "select (case when pmnud04 = 2 then nvl(sum(rvv39 * pmm42),0) else 0 end ) as t1 ,(case when pmnud04 = 1 then nvl(sum(rvv39 * pmm42),0) else 0 end ) as t2 "
        oCommander99.CommandText += "from TLF_FILE left join rvv_file on tlf905=rvv01 and tlf906 = rvv02 left join pmm_file on rvv36 = pmm01 left join pmn_file on rvv36 =pmn01 and rvv37 =pmn02 "
        oCommander99.CommandText += "where tlf13 = 'apmt150' and tlf01 like '7%' and tlf20 is not null and tlf20 <> ' ' and tlf20 = '"
        oCommander99.CommandText += pja01 & "' and tlf06 between to_date('" & DStartN.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommander99.CommandText += DstartE.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') group by pmnud04 "
        oReader99 = oCommander99.ExecuteReader()
        If oReader99.HasRows() Then
            While oReader99.Read()
                Ws.Cells(LineZ, 11) = oReader99.Item("t1")
                Ws.Cells(LineZ, 12) = oReader99.Item("t2")
            End While
        End If
        oReader99.Close()
    End Sub
End Class