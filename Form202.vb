Public Class Form202
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
    Dim Time1 As Date
    Dim Time2 As Date
    Dim CYear1 As Int16 = 0
    Dim CYear2 As Int16 = 0
    Dim CMonth1 As Int16 = 0
    Dim CMonth2 As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form202_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Time1 = DateTimePicker1.Value
        Time2 = DateTimePicker2.Value
        CYear1 = Time1.Year
        CYear2 = Time2.Year
        CMonth1 = Time1.Month
        CMonth2 = Time2.Month
        If CYear1 = CYear2 And CMonth1 = CMonth2 Then
            MsgBox("Date Error")
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
        

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "直接材料成本差异"
        AdjustExcelFormat()
        LineZ = 2
        oCommand.CommandText = "Select * from ( "
        oCommand.CommandText += "Select bma01,xx1,bma06,bmb03,xx2,ima08,bmb10,ima25,t1,sum(n1) as n1,sum(n2) as n2 ,sum(m1) as m1,sum(m2) as m2,bma05, bmaacti from ( "
        oCommand.CommandText += "Select bma01,s1.ima02 as xx1,bma06, bmb03, s2.ima02 as xx2,s2.ima08, bmb10, s2.ima25,Round((1/bmb10_fac),2) as t1, Round(sum(bmb06/bmb07 * (1+bmb08 /100)),2) as n1, stb07 / Round((1/bmb10_fac),2) as n2, 0 as m1, 0 as m2,bma05, bmaacti "
        oCommand.CommandText += "from bma_file left join bmb_file on bma01 = bmb01 and bma06 = bmb29 left join ima_file s1 on bma01 = s1.ima01 and bmb01 = s1.ima01 left join ima_file s2 on bmb03 = s2.ima01 "
        oCommand.CommandText += "left join stb_file on bmb03 = stb01 and stb02 = " & CYear1 & " and stb03 = " & CMonth1 & " where s1.ima08 in ('M','S') and bmaacti = 'Y' and bma05 <= to_date('"
        oCommand.CommandText += Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and not (length(bmb03) in (15, 16) and s2.ima08 = 'M') and bmb04 <= to_date('"
        oCommand.CommandText += Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 is null or bmb05 >= to_date('"
        oCommand.CommandText += Time1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) and bmb07 <> 0 group by bma01,s1.ima02,bma06, bmb03, s2.ima02,s2.ima08, bmb10, s2.ima25,(1/bmb10_fac), stb07,bma05, bmaacti "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select bma01,s1.ima02,bma06, bmb03, s2.ima02,s2.ima08, bmb10, s2.ima25,Round((1/bmb10_fac),2) as t1,0,0, Round(sum(bmb06/bmb07 * (1+bmb08 /100)),2), stb07 / Round((1/bmb10_fac),2),bma05, bmaacti "
        oCommand.CommandText += "from bma_file left join bmb_file on bma01 = bmb01 and bma06 = bmb29 left join ima_file s1 on bma01 = s1.ima01 and bmb01 = s1.ima01 left join ima_file s2 on bmb03 = s2.ima01 "
        oCommand.CommandText += "left join stb_file on bmb03 = stb01 and stb02 = " & CYear2 & " and stb03 = " & CMonth2 & " where s1.ima08 in ('M','S') and bmaacti = 'Y' and bma05 <= to_date('"
        oCommand.CommandText += Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and not (length(bmb03) in (15, 16) and s2.ima08 = 'M') and bmb04 <= to_date('"
        oCommand.CommandText += Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and (bmb05 is null or bmb05 >= to_date('"
        oCommand.CommandText += Time2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) and bmb07 <> 0 group by bma01,s1.ima02,bma06, bmb03, s2.ima02,s2.ima08, bmb10, s2.ima25,(1/bmb10_fac), stb07,bma05, bmaacti "
        oCommand.CommandText += ") group by bma01,xx1,bma06,bmb03,xx2,ima08,bmb10,ima25,t1,bma05, bmaacti ) where n1 <> m1 or n2 <> m2 order by bma01, bmb03"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To 12 Step 1
                    Ws.Cells(LineZ, 1 + i) = oReader.Item(i)
                Next
                Ws.Cells(LineZ, 14) = "=(L" & LineZ & "-J" & LineZ & ") * K" & LineZ
                Ws.Cells(LineZ, 15) = "=(M" & LineZ & "-K" & LineZ & ") * L" & LineZ
                Ws.Cells(LineZ, 16) = "=(L" & LineZ & "*M" & LineZ & ")-(J" & LineZ & "*K" & LineZ & ")"
                Ws.Cells(LineZ, 17) = oReader.Item("bma05")
                Ws.Cells(LineZ, 18) = oReader.Item("bmaacti")
                LineZ += 1
            End While
        End If


        oRng = Ws.Range("A1", "R1")
        oRng.EntireColumn.AutoFit()


    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.EntireRow.RowHeight = 16
        Ws.Cells(1, 1) = "主件料号"
        Ws.Cells(1, 2) = "品名"
        Ws.Cells(1, 3) = "特性代码"
        Ws.Cells(1, 4) = "元件料号"
        Ws.Cells(1, 5) = "品名"
        oRng = Ws.Range("D1", "D1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 6) = "来源码"
        Ws.Cells(1, 7) = "BOM表单位"
        Ws.Cells(1, 8) = "库存单位"
        Ws.Cells(1, 9) = "换算率"
        oRng = Ws.Range("I1", "P1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        Ws.Cells(1, 10) = "上期用量"
        Ws.Cells(1, 11) = "上期标准成本单价"
        Ws.Cells(1, 12) = "本期用量"
        Ws.Cells(1, 13) = "本期标准成本单价"
        Ws.Cells(1, 14) = "用量差异"
        Ws.Cells(1, 15) = "标准单价差异"
        Ws.Cells(1, 16) = "成本差异"
        Ws.Cells(1, 17) = "BOM表发放日期"
        Ws.Cells(1, 18) = "BOM表有效否"
        oRng = Ws.Range("B2", "B2")
        oRng.Select()
        'Ws.FreezePanes(2, 5)
        xExcel.ActiveWindow.FreezePanes = True
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "产品上下期标准成本-直接材料成本变动明细表"
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
End Class