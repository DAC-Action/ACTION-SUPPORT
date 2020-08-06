Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form186
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim tYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim xPath As String = String.Empty
    Dim Record1 As Boolean = False  ' 若 True 就不記錄
    Dim ReportType As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form186_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        Me.NumericUpDown1.Value = Today.Year
        Me.NumericUpDown2.Value = Today.Month
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If RadioButton1.Checked = True Then
            ReportType = 1
        Else
            ReportType = 2
        End If

        If ReportType = 1 Then
            xPath = "C:\temp\STD_GM_Template.xlsx"
        Else
            xPath = "C:\temp\ACT_GM_Template.xlsx"
        End If

        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
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
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        tYear = Me.NumericUpDown1.Value
        tMonth = Me.NumericUpDown2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        Dim Name1 As String = String.Empty
        If ReportType = 1 Then
            Name1 = "标准毛利率报表"
        Else
            Name1 = "实际毛利率报表"
        End If
        SaveFileDialog1.FileName = Name1
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
        If ReportType = 1 Then
            XPath = "C:\temp\STD_GM_Template.xlsx"
        Else
            XPath = "C:\temp\ACT_GM_Template.xlsx"
        End If
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(2, 4 + i) = tYear & "/" & i & "/01"
        Next
        LineZ = 4
        ' 20200806 Cloud
        'Dim EndDate As Date = Convert.ToDateTime(tYear & "/" & tMonth & "/01")
        'EndDate = EndDate.AddMonths(1).AddDays(-1)
        oCommand.CommandText = "select bma01,ima02,tqa02,ima25, bma05 from bma_file left join ima_file on bma01 = ima01 left join tqa_file on tqa03 = '2' and ima1005 = tqa01 where ima06 = '103' and bma10 = 2 and bmaacti = 'Y' " 'and bma05 < to_date('" & EndDate.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                Record1 = False
                DetailedData(tYear, tMonth, oReader.Item(0), oReader.Item(4))
                If Record1 = False Then
                    Ws.Cells(LineZ, 1) = oReader.Item(0)
                    Ws.Cells(LineZ, 2) = oReader.Item(1)
                    Ws.Cells(LineZ, 3) = oReader.Item(2)
                    Ws.Cells(LineZ, 4) = oReader.Item(3)
                    '销售总额 (20/01/13 Brady memo)
                    DetailData2(tYear, tMonth, oReader.Item(0))
                    '平均毛利 (20/01/13 Brady memo)
                    Ws.Cells(LineZ, 18) = "=(SUM(E" & LineZ & ":P" & LineZ & ")-SUMIF(E" & LineZ & ":P" & LineZ & ",""1"",E" & LineZ & ":P" & LineZ & "))/(COUNTA(E" & LineZ & ":P" & LineZ & ")-COUNTIF(E" & LineZ & ":P" & LineZ & ",""=0"")-COUNTIF(E" & LineZ & ":P" & LineZ & ",""=1""))"
                    LineZ += 1
                    Label3.Text = LineZ
                    Label3.Refresh()
                End If
            End While
        End If
        oReader.Close()
        oRng = Ws.Range("A1", "R1")
        oRng.EntireColumn.AutoFit()
    End Sub
    Private Sub DetailedData(ByVal Year1 As Int16, ByVal Month1 As Int16, ByVal ima01 As String, ByVal bma05 As Date)
        For i As Int16 = Month1 To 1 Step -1
            If ReportType = 1 Then
                ' 20200806 Cloud
                Dim T1 As Date = Convert.ToDateTime(Year1 & "/" & i & "/01")
                Dim T2 As Date = Convert.ToDateTime(bma05.Year & "/" & bma05.Month & "/01")
                If T1 < T2 Then
                    Exit For
                End If
                oCommand2.CommandText = "Select nvl(Round((stb07+stb08+stb09+stb09a) /ex1.er,4),0) from stb_file left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD' "
                oCommand2.CommandText += "where stb01 = '" & ima01 & "' and stb02 = " & Year1 & " and stb03 = " & i
            Else
                oCommand2.CommandText = "Select nvl(Round((ccc23) /ex1.er,4),0) from ccc_file left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD' "
                oCommand2.CommandText += "where ccc01 = '" & ima01 & "' and ccc02 = " & Year1 & " and ccc03 = " & i
            End If
            'oCommand2.CommandText = "Select nvl(Round((stb07+stb08+stb09+stb09a) /ex1.er,4),0) from stb_file left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD' "
            'oCommand2.CommandText += "where stb01 = '" & ima01 & "' and stb02 = " & Year1 & " and stb03 = " & i
            Dim STDCostUSD As Decimal = oCommand2.ExecuteScalar()
            If i = Month1 And STDCostUSD = 0 Then
                Record1 = True
                Exit For    '當月沒資料就離開 (20/01/13 Brady memo)
            End If
            Dim TempDate As Date = Convert.ToDateTime(Year1 & "/" & i & "/01")
            Dim TempDate1 As Date = TempDate.AddMonths(1).AddDays(-1)
            oCommand2.CommandText = " Select (case when tc_prl06 ='USD' then tc_prl03 * tc_prl04 /100 else nvl(Round(tc_prl03 * tc_prl04 /100 * ex1.er / ex2.er,4),0) end) from tc_prl_file  left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = tc_prl06 "
            oCommand2.CommandText += "left join exchangeratebyyear ex2 on ex2.year1 = " & Year1 & " and ex2.currency = 'USD' where tc_prl01  = '" & ima01 & "' and tc_prl02 >= to_date('" & TempDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by tc_prl02"
            oReader2 = oCommand2.ExecuteReader
            Dim SaleSRMB As Decimal = 0
            If oReader2.HasRows() Then
                oReader2.Read()
                SaleSRMB = oReader2.Item(0)
            Else
                If i = Month1 Then
                    Record1 = True
                    Exit For
                End If
            End If
            oReader2.Close()
            Dim Perce1 As Decimal = Decimal.Round((SaleSRMB - STDCostUSD) / SaleSRMB, 4)
            Ws.Cells(LineZ, i + 4) = Perce1
        Next
    End Sub
    Private Sub DetailData2(ByVal Year1 As Int16, ByVal Month1 As Int16, ByVal ima01 As String)
        oCommand2.CommandText = " Select nvl(Round(sum(ccc63)/ex1.er,4),0) from ccc_file  left join exchangeratebyyear ex1 on ex1.year1 = " & Year1 & " and ex1.currency = 'USD'  where ccc01 = '"
        oCommand2.CommandText += ima01 & "' and ccc02 = " & Year1 & " and ccc03 <= " & Month1 & " group by ex1.er"
        Dim SS As Decimal = oCommand2.ExecuteScalar()
        Ws.Cells(LineZ, 17) = SS
    End Sub
End Class