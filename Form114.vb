Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlChartType
Imports Microsoft.Office.Core.MsoChartElementType
'Imports Microsoft.Office.Interop.Excel.XlRowCol
Public Class Form114
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommand2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader2 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim TYear As String = String.Empty
    Dim TMonth As String = String.Empty
    Dim CYear As String = String.Empty
    Dim CMonth As String = String.Empty
    Dim g_oga03 As String = String.Empty
    Dim LineZ As Integer = 0
    Dim LineS1 As Int16 = 0
    Dim LineS2 As Int16 = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form114_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        Me.DateTimePicker1.Value = Today
        Me.DateTimePicker2.Value = Today

    End Sub

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
        g_oga03 = String.Empty
        If Not String.IsNullOrEmpty(TextBox1.Text) Then
            g_oga03 = TextBox1.Text
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "HAC_Sale_WeekReport"
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
        AdjustExcelFormat1()

        '191029 add by Brady CS要求調整程式邏輯
        '因为发现当收款客户相同，但不同的账款客户做在同一张出货单上时，香港端的出货单会把所有不同的账款客户默认为其中一个账款客户，
        '但实际出货应该是属于不同的账款客户的—>由此HAC的销售周报有时体现的数据不真实。
        '从取香港端axmt820a账款客户编号 改取 香港端axmt810客户编号(oea03)
        'oCommand.CommandText = "select oga03,oga032"
        'For i As Int16 = 1 To 53 Step 1
        '    oCommand.CommandText += ",sum(t" & i & ") as t" & i
        'Next
        'oCommand.CommandText += " from ( select oga03,oga032"
        'For i As Int16 = 1 To 53 Step 1
        '    oCommand.CommandText += ",(case when azn05 = " & i & " then ogb14t * oga24 else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from hkacttest.oga_file left join hkacttest.ogb_file on oga01 = ogb01 "
        'oCommand.CommandText += "left join azn_file on oga02 = azn01 where ogapost = 'Y' and oga04 <> 'D0003' and oga02 between to_date('"
        'oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        'oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb04 <> 'AC0000000000' "
        'If Not String.IsNullOrEmpty(g_oga03) Then
        '    oCommand.CommandText += " AND oga03 ='" & g_oga03 & "' "
        'End If
        '' 20180312 add oha ohb
        'oCommand.CommandText += "union all "
        'oCommand.CommandText += "select oha03,oha032"
        'For i As Int16 = 1 To 53 Step 1
        '    oCommand.CommandText += ",(case when azn05 = " & i & " then ohb14t * oha24 * (-1) else 0 end) as t" & i
        'Next
        'oCommand.CommandText += " from hkacttest.oha_file left join hkacttest.ohb_file on oha01 = ohb01 "
        'oCommand.CommandText += "left join azn_file on oha02 = azn01 where ohapost = 'Y' and oha04 <> 'D0003' and oha02 between to_date('"
        'oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        'oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb04 <> 'AC0000000000' "
        'If Not String.IsNullOrEmpty(g_oga03) Then
        '    oCommand.CommandText += " AND oha03 ='" & g_oga03 & "' "
        'End If
        'oCommand.CommandText += " ) group by oga03,oga032 order by oga03"
        'oReader = oCommand.ExecuteReader()
        'If oReader.HasRows() Then
        '    While oReader.Read()
        '        For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
        '            Ws.Cells(LineZ, i + 1) = oReader.Item(i)
        '        Next
        '        Ws.Cells(LineZ, 56) = "=SUM(C" & LineZ & ":BC" & LineZ & ")"
        '        LineZ += 1
        '    End While
        '    oRng = Ws.Range("C6", Ws.Cells(LineZ, 56))
        '    oRng.NumberFormatLocal = "[$$-en-CA]#,##0.00;-[$$-en-CA]#,##0.00"
        '    oRng = Ws.Range("BD6", Ws.Cells(LineZ, 56))
        '    oRng.Interior.Color = Color.DarkGray
        '    Ws.Cells(LineZ, 1) = "Total"
        '    Ws.Cells(LineZ, 3) = "=SUM(C6:C" & LineZ - 1 & ")"
        '    oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
        '    oRng.Interior.Color = Color.DarkGray
        '    oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 56)), Type:=xlFillDefault)
        '    'oRng.AutoFill(Destination:=Range("C17:BB17"), Type:=xlFillDefault)
        '    oRng = Ws.Range("A6", Ws.Cells(LineZ, 56))
        '    oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        '    oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        '    oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        '    oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        '    oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '    oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        'End If
        'oReader.Close()
        oCommand.CommandText = "select oga03,oga032"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( select oea03 as oga03,oea032 as oga032"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then ogb14t * oga24 else 0 end) as t" & i
        Next
        oCommand.CommandText += " from hkacttest.oea_file,hkacttest.oeb_file,hkacttest.ogb_file,hkacttest.oga_file "
        oCommand.CommandText += "left join azn_file on oga02 = azn01 where ogapost = 'Y' and oga04 <> 'D0003' and oga02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb04 <> 'AC0000000000' "
        oCommand.CommandText += " and oga01 = ogb01 and ogb31 = oeb01 and ogb32 = oeb03 and oea01 = oeb01 "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND oga03 ='" & g_oga03 & "' "
        End If
        ' 20180312 add oha ohb
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select oha03,oha032"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then ohb14t * oha24 * (-1) else 0 end) as t" & i
        Next
        oCommand.CommandText += " from hkacttest.oha_file left join hkacttest.ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join azn_file on oha02 = azn01 where ohapost = 'Y' and oha04 <> 'D0003' and oha02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb04 <> 'AC0000000000' "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND oha03 ='" & g_oga03 & "' "
        End If
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select customercode, customername"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then Amount else 0 end) "
        Next
        oCommand.CommandText += " from vac_shipment_2 left join azn_file on shipmentdate = azn01 where shipmentdate between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND customercode ='" & g_oga03 & "' "
        End If
        oCommand.CommandText += " ) group by oga03,oga032 order by oga03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                Ws.Cells(LineZ, 56) = "=SUM(C" & LineZ & ":BC" & LineZ & ")"
                LineZ += 1
                End While
            oRng = Ws.Range("C6", Ws.Cells(LineZ, 56))
            oRng.NumberFormatLocal = "[$$-en-CA]#,##0.00;-[$$-en-CA]#,##0.00"
            oRng = Ws.Range("BD6", Ws.Cells(LineZ, 56))
            oRng.Interior.Color = Color.DarkGray
            Ws.Cells(LineZ, 1) = "Total"
            Ws.Cells(LineZ, 3) = "=SUM(C6:C" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
            oRng.Interior.Color = Color.DarkGray
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 56)), Type:=xlFillDefault)
                'oRng.AutoFill(Destination:=Range("C17:BB17"), Type:=xlFillDefault)
            oRng = Ws.Range("A6", Ws.Cells(LineZ, 56))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
            End If
        oReader.Close()
            '191029 add by Brady END

        LineS1 = LineZ
        LineZ += 2
        AdjustExcelFormat2()
        oCommand.CommandText = "select oga03,oga032"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( select oea03 as oga03,oea032 as oga032"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then ogb14t * oga24 else 0 end) as t" & i
        Next
        oCommand.CommandText += " from hkacttest.oea_file,hkacttest.oeb_file,hkacttest.ogb_file,hkacttest.oga_file "
        oCommand.CommandText += "left join azn_file on oga02 = azn01 where ogapost = 'Y' and oga04 <> 'D0003' and oga02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ogb04 = 'AC0000000000' "
        oCommand.CommandText += " and oga01 = ogb01 and ogb31 = oeb01 and ogb32 = oeb03 and oea01 = oeb01 "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND oga03 ='" & g_oga03 & "' "
            End If
            ' 20180312 add oha ohb
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select oha03,oha032"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then ohb14t * oha24 * (-1) else 0 end) as t" & i
        Next
        oCommand.CommandText += " from hkacttest.oha_file left join hkacttest.ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join azn_file on oha02 = azn01 where ohapost = 'Y' and oha04 <> 'D0003' and oha02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ohb04 = 'AC0000000000' "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND oha03 ='" & g_oga03 & "' "
            End If
        oCommand.CommandText += " ) group by oga03,oga032 order by oga03"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                Ws.Cells(LineZ, 56) = "=SUM(C" & LineZ & ":BC" & LineZ & ")"
                LineZ += 1
                End While
            oRng = Ws.Range(Ws.Cells(LineS1 + 4, 3), Ws.Cells(LineZ, 56))
            oRng.NumberFormatLocal = "[$$-en-CA]#,##0.00;-[$$-en-CA]#,##0.00"
            oRng = Ws.Range(Ws.Cells(LineS1 + 4, 56), Ws.Cells(LineZ, 56))
            oRng.Interior.Color = Color.DarkGray
            Ws.Cells(LineZ, 1) = "Total"
            Ws.Cells(LineZ, 3) = "=SUM(C" & LineS1 + 4 & ":C" & LineZ - 1 & ")"
            oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 3))
            oRng.Interior.Color = Color.DarkGray
            oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ, 56)), Type:=xlFillDefault)
                'oRng.AutoFill(Destination:=Range("C17:BB17"), Type:=xlFillDefault)
            oRng = Ws.Range(Ws.Cells(LineS1 + 4, 1), Ws.Cells(LineZ, 56))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
            End If
        oReader.Close()
        LineS2 = LineZ
            ' 試作 圖表 20180315
            'LineZ += 2
        Dim XA As Excel.Chart = Ws.Shapes.AddChart(xlColumnClustered, 50, 600, 2500, 600).Chart
        oRng = Ws.Range("B5", Ws.Cells(LineS1 - 1, 55))
        XA.SetSourceData(oRng)

        XA.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        XA.ChartTitle.Text = "Weekly HAC Part sales"
        XA.SetElement(msoElementLegendNone)
        XA.SetElement(msoElementDataTableWithLegendKeys)

        Dim XB As Excel.Chart = Ws.Shapes.AddChart(xlColumnClustered, 50, 1300, 2500, 400).Chart
        oRng = Ws.Range(Ws.Cells(LineS1 + 3, 2), Ws.Cells(LineZ - 1, 55))
        XB.SetSourceData(oRng)

        XB.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        XB.ChartTitle.Text = "Weekly HAC Tooling sales"
        XB.SetElement(msoElementLegendNone)
        XB.SetElement(msoElementDataTableWithLegendKeys)
            ' 先試看看
        LineZ = 115
        For i As Int16 = 6 To LineS1 - 1 Step 1
            Ws.Cells(LineZ, 2) = "=B" & i
            Ws.Cells(LineZ, 3) = "=BD" & i
            LineZ += 1
        Next
        Ws.Cells(LineZ, 3) = "=SUM(C110:C" & LineZ - 1 & ")"

        Dim XC As Excel.Chart = Ws.Shapes.AddChart(xlPie, 350, 1800, 400, 400).Chart
        oRng = Ws.Range(Ws.Cells(110, 2), Ws.Cells(LineZ - 1, 3))
        XC.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)

        XC.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        XC.ChartTitle.Text = "HAC Part sales"
        XC.SetElement(msoElementLegendNone)
        XC.SetElement(msoElementDataLabelBestFit)
            'XC.SetElement(msoElementDataLabelShow)
        XC.SeriesCollection(1).DataLabels.ShowCategoryName = True
        XC.SeriesCollection(1).DataLabels.ShowValue = False
        XC.SeriesCollection(1).DataLabels.ShowPercentage = True

            ' 150行開始
        LineZ = 150
        For i As Int16 = LineS1 + 4 To LineS2 - 1 Step 1
            Ws.Cells(LineZ, 2) = "=B" & i
            Ws.Cells(LineZ, 3) = "=BD" & i
            LineZ += 1
        Next
        Ws.Cells(LineZ, 3) = "=SUM(C150:C" & LineZ - 1 & ")"

        Dim XD As Excel.Chart = Ws.Shapes.AddChart(xlPie, 350, 2300, 400, 400).Chart
        oRng = Ws.Range(Ws.Cells(150, 2), Ws.Cells(LineZ - 1, 3))
        XD.SetSourceData(oRng, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)

        XD.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        XD.ChartTitle.Text = "HAC Tooling sales"
        XD.SetElement(msoElementLegendNone)
        XD.SetElement(msoElementDataLabelBestFit)
            'XC.SetElement(msoElementDataLabelShow)
        XD.SeriesCollection(1).DataLabels.ShowCategoryName = True
        XD.SeriesCollection(1).DataLabels.ShowValue = False
        XD.SeriesCollection(1).DataLabels.ShowPercentage = True

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        AdjustExcelFormat3()
        oCommand.CommandText = "select 1"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",sum(t" & i & ") as t" & i
        Next
        oCommand.CommandText += " from ( select 1"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then ogb14t * oga24 else 0 end) as t" & i
        Next
        oCommand.CommandText += " from hkacttest.oga_file left join hkacttest.ogb_file on oga01 = ogb01 "
        oCommand.CommandText += "left join azn_file on oga02 = azn01 where ogapost = 'Y' and oga04 <> 'D0003' and oga02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND oga03 ='" & g_oga03 & "' "
            End If
            ' 20180312 add oha ohb
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select 1"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then ohb14t * oha24 * (-1) else 0 end) as t" & i
        Next
        oCommand.CommandText += " from hkacttest.oha_file left join hkacttest.ohb_file on oha01 = ohb01 "
        oCommand.CommandText += "left join azn_file on oha02 = azn01 where ohapost = 'Y' and oha04 <> 'D0003' and oha02 between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND oha03 ='" & g_oga03 & "' "
        End If
        oCommand.CommandText += "union all "
        oCommand.CommandText += "Select 1"
        For i As Int16 = 1 To 53 Step 1
            oCommand.CommandText += ",(case when azn05 = " & i & " then Amount else 0 end) "
        Next
        oCommand.CommandText += " from vac_shipment_2 left join azn_file on shipmentdate = azn01 where shipmentdate between to_date('"
        oCommand.CommandText += Me.DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        oCommand.CommandText += Me.DateTimePicker2.Value.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') "
        If Not String.IsNullOrEmpty(g_oga03) Then
            oCommand.CommandText += " AND customercode ='" & g_oga03 & "' "
        End If
        oCommand.CommandText += " )"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                For i As Int16 = 1 To oReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = oReader.Item(i)
                Next
                LineZ += 1
                End While
            End If
        oReader.Close()
            'oCommand.CommandText = "select sum(t1) from ( select (case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ) as t1 from hkacttest.tc_bud_file where tc_bud02 = "
            'oCommand.CommandText += Me.DateTimePicker2.Value.Year & " and tc_bud06 in ('Alex','Howard') ) "
            'Dim TotalBudget As Decimal = oCommand.ExecuteScalar()
            'oCommand.CommandText = "select count(azn01) from azn_file where azn02 = " & Me.DateTimePicker2.Value.Year
            'Dim TotalDAys As Decimal = oCommand.ExecuteScalar()
        For i As Int16 = 1 To 53
            oCommand.CommandText = "select azn04,count(azn01) as t1 FROM azn_file where azn02 = " & Me.DateTimePicker2.Value.Year & " and azn05 = " & i & " group by azn04"
            oReader = oCommand.ExecuteReader()
            Dim WeekMoney As Decimal = 0
            If oReader.HasRows() Then
                While oReader.Read()
                    oCommand2.CommandText = "select nvl(sum(t1),0) from ( select (case when tc_bud14 = 'EUR' then tc_bud13 * 1.2 else tc_bud13 end ) as t1 from hkacttest.tc_bud_file where tc_bud02 = "
                        'oCommand2.CommandText += Me.DateTimePicker2.Value.Year & " and tc_bud03 = " & oReader.Item("azn04") & " and tc_bud06 in ('Alex','Howard') ) "
                    oCommand2.CommandText += Me.DateTimePicker2.Value.Year & " and tc_bud03 = " & oReader.Item("azn04") & " and tc_bud06 in ('USA/Japan') ) "
                    Dim MonthBudget As Decimal = oCommand2.ExecuteScalar()
                    oCommand2.CommandText = "select nvl(count(azn01),0) as t1 FROM azn_file where azn02 = " & Me.DateTimePicker2.Value.Year & " and azn04 = " & oReader.Item("azn04") & " group by azn04"
                    Dim MonthTotalDay As Decimal = oCommand2.ExecuteScalar()
                    If MonthTotalDay <> 0 Then
                        WeekMoney += (MonthBudget / MonthTotalDay) * oReader.Item("t1")
                        End If
                    End While
                End If
            oReader.Close()


                'oCommand.CommandText = "select count(azn01) from azn_file where azn02 = " & Me.DateTimePicker2.Value.Year
                'oCommand.CommandText += " and azn05 = " & i
                'Dim dayonweek As Int16 = oCommand.ExecuteScalar()
                'Ws.Cells(LineZ, i + 1) = TotalBudget / TotalDAys * dayonweek
            Ws.Cells(LineZ, i + 1) = WeekMoney
        Next
        LineZ += 1
        LineS1 = LineZ
        oRng = Ws.Range("B3", "BB4")
        oRng.NumberFormatLocal = "[$$-en-CA]#,##0.00;-[$$-en-CA]#,##0.00"
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        LineZ = 7
        Ws.Cells(7, 2) = "=B3"
        Ws.Cells(7, 3) = "=C3+B7"
        oRng = Ws.Range("C7", "C7")
        oRng.AutoFill(Destination:=Ws.Range("C7", "BB7"), Type:=xlFillDefault)

        Ws.Cells(8, 2) = "=B4"
        Ws.Cells(8, 3) = "=C4+B8"
        oRng = Ws.Range("C8", "C8")
        oRng.AutoFill(Destination:=Ws.Range("C8", "BB8"), Type:=xlFillDefault)

        oRng = Ws.Range("B7", "BB8")
        oRng.NumberFormatLocal = "[$$-en-CA]#,##0.00;-[$$-en-CA]#,##0.00"
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        Ws.Cells(10, 2) = "=(B8-B7)*-1"
        oRng = Ws.Range("B10", "B10")
        oRng.AutoFill(Destination:=Ws.Range("B10", "BB10"), Type:=xlFillDefault)
        oRng = Ws.Range("B10", "BB10")
        oRng.NumberFormat = "[$$-en-CA]#,##0.00;[Red]-[$$-en-CA]#,##0.00"

        Ws.Cells(11, 2) = "=B10/B8"
        oRng = Ws.Range("B11", "B11")
        oRng.AutoFill(Destination:=Ws.Range("B11", "BB11"), Type:=xlFillDefault)
        oRng = Ws.Range("B11", "BB11")
        oRng.NumberFormatLocal = "0%"

        oRng = Ws.Range("A10", "BB11")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            ' 開始劃圖 A
        Dim YA As Excel.Chart = Ws.Shapes.AddChart(xlColumnClustered, 20, 300, 4000, 400).Chart
        oRng = Ws.Range("A2", "BB4")
        YA.SetSourceData(oRng)
        YA.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        YA.ChartTitle.Text = "Weekly HAC Sales VS. Target"
        YA.SetElement(msoElementLegendNone)
        YA.SetElement(msoElementDataTableWithLegendKeys)

            '圖B
        Dim YB As Excel.Chart = Ws.Shapes.AddChart(xlLine, 20, 720, 4000, 400).Chart
        oRng = Ws.Range("A2", "BB4")
        YB.SetSourceData(oRng)
        YB.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        YB.ChartTitle.Text = "Weekly HAC Sales VS. Target"
        YB.SetElement(msoElementLegendNone)
        YB.SetElement(msoElementDataTableWithLegendKeys)

            '圖C
        Dim YC As Excel.Chart = Ws.Shapes.AddChart(xlLine, 20, 1140, 4000, 400).Chart
        oRng = Ws.Range("A6", "BB8")
        YC.SetSourceData(oRng)
        YC.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        YC.ChartTitle.Text = "YTD HAC Sales VS. Target"
        YC.SetElement(msoElementLegendNone)
        YC.SetElement(msoElementDataTableWithLegendKeys)

            '圖D
        Dim YD As Excel.Chart = Ws.Shapes.AddChart(xlLine, 20, 1560, 4000, 400).Chart
        oRng = Ws.Range("A10", "BB10")
        YD.SetSourceData(oRng)
        YD.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        YD.ChartTitle.Text = "YTD Comparsion Budget / Current USD " & Me.DateTimePicker2.Value.Year
        YD.SetElement(msoElementLegendNone)
        YD.SetElement(msoElementDataTableWithLegendKeys)

            '圖E
        Dim YE As Excel.Chart = Ws.Shapes.AddChart(xlLine, 20, 1980, 4000, 400).Chart
        oRng = Ws.Range("A11", "BB11")
        YE.SetSourceData(oRng)
        YE.SetElement(msoElementChartTitleAboveChart)
            'XA.ActiveChart.SetElement(msoElementLegendNone)
        YE.ChartTitle.Text = "YTD Comparsion Budget / Current " & Me.DateTimePicker2.Value.Year & " %"
        YE.SetElement(msoElementLegendNone)
        YE.SetElement(msoElementDataTableWithLegendKeys)
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Sales Split Customer HAC"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 7.9
        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.ColumnWidth = 19.3
        oRng = Ws.Range("A1", "BD1")
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(1, 1) = "Sales Split by Customers HAC"
        oRng = Ws.Range("A2", "BD2")
        oRng.Merge()
        oRng.Font.Size = 16
        Ws.Cells(2, 1) = Me.DateTimePicker1.Value.Year & "年"
        oRng = Ws.Range("A4", "BD4")
        oRng.Merge()
        oRng.Font.Size = 12
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(4, 1) = "HAC Part"
        Ws.Cells(5, 1) = "客户编号Customer Code"
        Ws.Cells(5, 2) = "客户简称C_SName"
        For i As Int16 = 1 To 53 Step 1
            Ws.Cells(5, i + 2) = "W" & i
        Next
        Ws.Cells(5, 56) = "Total"
        oRng = Ws.Range("A5", "BD5")
        oRng.Interior.Color = Color.DarkGray
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        LineZ = 6
    End Sub
    Private Sub AdjustExcelFormat2()
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 56))
        oRng.Merge()
        oRng.HorizontalAlignment = xlLeft
        Ws.Cells(LineZ, 1) = "HAC Tooling"
        LineZ += 1
        Ws.Cells(LineZ, 1) = "客户编号Customer Code"
        Ws.Cells(LineZ, 2) = "客户简称C_SName"
        For i As Int16 = 1 To 53 Step 1
            Ws.Cells(LineZ, i + 2) = "W" & i
        Next
        Ws.Cells(LineZ, 56) = "Total"
        oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ, 56))
        oRng.Interior.Color = Color.DarkGray
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        LineZ += 1
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Sales Compare to Budget HAC"
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.ColumnWidth = 12.8
        Ws.Rows.RowHeight = 25.2
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 28.4
        oRng = Ws.Range("A1", "BB1")
        oRng.Merge()
        oRng.Font.Size = 20
        oRng.EntireRow.RowHeight = 39
        Ws.Cells(1, 1) = "Sales Compare to Budget HAC " & Me.DateTimePicker1.Value.Year
        Ws.Cells(3, 1) = "Current Sales"
        Ws.Cells(4, 1) = "Sales Target"
        For i As Int16 = 1 To 53 Step 1
            Ws.Cells(2, i + 1) = "W" & i
            Ws.Cells(6, i + 1) = "W" & i
        Next
        oRng = Ws.Range("B2", "BB2")
        oRng.Interior.Color = Color.FromArgb(196, 189, 151)
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("A3", "A4")
        oRng.Interior.Color = Color.FromArgb(196, 189, 151)
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        Ws.Cells(6, 1) = "YTD"
        Ws.Cells(7, 1) = "current sales"
        Ws.Cells(8, 1) = "Sales Target"
        oRng = Ws.Range("A6", "BB6")
        oRng.Interior.Color = Color.FromArgb(149, 179, 215)
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        oRng = Ws.Range("A7", "A8")
        oRng.Interior.Color = Color.FromArgb(149, 179, 215)
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        Ws.Cells(10, 1) = "YTD Comparison Budget / Current USD"
        Ws.Cells(11, 1) = "YTD Comparison Budget / Current %"

        LineZ = 3

    End Sub
End Class