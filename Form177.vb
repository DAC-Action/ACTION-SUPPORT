Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form177
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
    Dim l_oeb04 As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form177_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        TextBox1.Text = Today.Year
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
        l_oeb04 = String.Empty
        If Not String.IsNullOrEmpty(TextBox2.Text) Then
            l_oeb04 = TextBox2.Text
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'ExportToExcel()
        TransferDB()
        ExportToExcel()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "DAC销售预算报表" & tYear
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
        Ws.Name = "Customer Demand Combined Qty"
        Ws.Activate()
        AdjustExcelFormat1()
        'oCommand.CommandText = "select oeb04,ima02,ima021,ima25,tc_cif_05,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,"
        'oCommand.CommandText += "sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ("
        'oCommand.CommandText += "select oeb04,ima02,ima021,ima25,tc_cif_05,(case when month(tc_cif_05) = 1 then tc_cif_04 else 0 end ) as t1,"
        'oCommand.CommandText += "(case when month(tc_cif_05) = 2 then tc_cif_04 else 0 end ) as t2,(case when month(tc_cif_05) = 3 then tc_cif_04 else 0 end ) as t3,"
        'oCommand.CommandText += "(case when month(tc_cif_05) = 4 then tc_cif_04 else 0 end ) as t4,(case when month(tc_cif_05) = 5 then tc_cif_04 else 0 end ) as t5,"
        'oCommand.CommandText += "(case when month(tc_cif_05) = 6 then tc_cif_04 else 0 end ) as t6,(case when month(tc_cif_05) = 7 then tc_cif_04 else 0 end ) as t7,"
        'oCommand.CommandText += "(case when month(tc_cif_05) = 8 then tc_cif_04 else 0 end ) as t8,(case when month(tc_cif_05) = 9 then tc_cif_04 else 0 end ) as t9,"
        'oCommand.CommandText += "(case when month(tc_cif_05) = 10 then tc_cif_04 else 0 end ) as t10,(case when month(tc_cif_05) = 11 then tc_cif_04 else 0 end ) as t11,"
        'oCommand.CommandText += "(case when month(tc_cif_05) = 12 then tc_cif_04 else 0 end ) as t12 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        'oCommand.CommandText += "left join ima_file on oeb04 = ima01 where tc_cif_05 between to_date('"
        'oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
        'oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 not like 'FC%'"
        'If Not String.IsNullOrEmpty(l_oeb04) Then
        '    oCommand.CommandText += " AND oeb04 like '%" & l_oeb04 & "%' "
        'End If
        'oCommand.CommandText += ") group by oeb04,ima02,ima021,ima25,tc_cif_05"
        'oCommand.CommandText = "select pn,ima02,ima021,ima25,year1,week1,max(azn01),(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4, (case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6, (case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'If Not String.IsNullOrEmpty(l_oeb04) Then
        '    oCommand.CommandText += " WHERE pn like '%" & l_oeb04 & "%' "
        'End If
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity order by pn,max(azn01)"
        oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01),(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, (case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4, (case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6, (case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        If Not String.IsNullOrEmpty(l_oeb04) Then
            oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
        End If
        oCommand.CommandText += "group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04 order by tc_prm01,max(azn01)"

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
        Ws.Cells(LineZ, 5) = "Total"
        Ws.Cells(LineZ, 9) = "=SUM(I5:I" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)

        oReader.Close()
        ' 劃線
        oRng = Ws.Range("B3", Ws.Cells(LineZ, 21))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 第二頁
        Ws = xWorkBook.Sheets(2)
        Ws.Name = "Amount（Detail）"
        Ws.Activate()
        AdjustExcelFormat2()
        'oCommand.CommandText = "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03 , max(A2.azn01) as c1,(case when month(max(A2.azn01))= 1 then tc_prm04 else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(A2.azn01))= 2 then tc_prm04 else 0 end) as t2,(case when month(max(A2.azn01))= 3 then tc_prm04 else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(A2.azn01))= 4 then tc_prm04 else 0 end) as t4,(case when month(max(A2.azn01))= 5 then tc_prm04 else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(A2.azn01))= 6 then tc_prm04 else 0 end) as t6,(case when month(max(A2.azn01))= 7 then tc_prm04 else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(A2.azn01))= 8 then tc_prm04 else 0 end) as t8,(case when month(max(A2.azn01))= 9 then tc_prm04 else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(A2.azn01))= 10 then tc_prm04 else 0 end) as t10,(case when month(max(A2.azn01))= 11 then tc_prm04 else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(A2.azn01))= 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 "
        'oCommand.CommandText += "left join (select oeb04,max(A1.azn02) as c1,max(A1.azn05) as c2 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
        'oCommand.CommandText += "left join azn_file A1 on tc_cif_05 = A1.azn01  where year(tc_cif_05) = " & tYear & " and tc_cif_01 not like 'FC%' and tc_cif_04 <> 0 "
        'oCommand.CommandText += " group by oeb04 ) X1 ON tc_prm01 = X1.oeb04 left join azn_file A2 on tc_prm02 = A2.azn02 and tc_prm03 = A2.azn05 where tc_prm02 = " & tYear
        'oCommand.CommandText += " And (X1.c2 Is null Or tc_prm03 > X1.C2) "
        'If Not String.IsNullOrEmpty(l_oeb04) Then
        '    oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
        'End If
        'oCommand.CommandText += "group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03, tc_prm04"
        'oCommand.CommandText = "Select pn,ima02,ima021,ima25,C1,(case when tc_prl02 is null and tc_prn02 is null then '3' when tc_prl02 is null then '2' else '1' end) as XS1,"
        'oCommand.CommandText += "(case when tc_prl02 is null then tc_prn02 else tc_prl02 end) as C2,(case when tc_prl02 is null then tc_prn06 else tc_prl06 end) as C3,"
        'oCommand.CommandText += "(case when tc_prl02 is null then tc_prn03 else tc_prl03 end) as C4,(case when tc_prl02 is null then t1 * tc_prn03 else t1 * tc_prl03 end) as t1,"
        'oCommand.CommandText += "(case when tc_prl02 is null then t2 * tc_prn03 else t2 * tc_prl03 end) as t2,(case when tc_prl02 is null then t3 * tc_prn03 else t3 * tc_prl03 end) as t3,"
        'oCommand.CommandText += "(case when tc_prl02 is null then t4 * tc_prn03 else t4 * tc_prl03 end) as t4,(case when tc_prl02 is null then t5 * tc_prn03 else t5 * tc_prl03 end) as t5,"
        'oCommand.CommandText += "(case when tc_prl02 is null then t6 * tc_prn03 else t6 * tc_prl03 end) as t6,(case when tc_prl02 is null then t7 * tc_prn03 else t7 * tc_prl03 end) as t7,"
        'oCommand.CommandText += "(case when tc_prl02 is null then t8 * tc_prn03 else t8 * tc_prl03 end) as t8,(case when tc_prl02 is null then t9 * tc_prn03 else t9 * tc_prl03 end) as t9,"
        'oCommand.CommandText += "(case when tc_prl02 is null then t10 * tc_prn03 else t10 * tc_prl03 end) as t10,(case when tc_prl02 is null then t11 * tc_prn03 else t11 * tc_prl03 end) as t11,"
        'oCommand.CommandText += "(case when tc_prl02 is null then t12 * tc_prn03 else t12 * tc_prl03 end) as t12 from ( "
        'oCommand.CommandText += "select pn,ima02,ima021,ima25,year1,week1,max(azn01) as c1,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, (case when month(max(azn01)) = 3 then quantity else 0 end) as t3,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 4 then quantity else 0 end) as t4,(case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6,(case when month(max(azn01)) = 7 then quantity else 0 end) as t7,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 "
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,year1,week1,quantity order by pn,max(azn01) ) AB left join tc_prl_file XX1 on AB.pn = XX1.tc_prl01 AND AB.C1 <= XX1.tc_prl02  "
        'oCommand.CommandText += "left join tc_prn_file XX3 on AB.pn = XX3.TC_PRn01 WHERE (XX1.tc_prl02 = (Select min(XX2.tc_prl02) from tc_prl_file XX2 where AB.PN = XX2.tc_prl01 and AB.C1 <= XX2.tc_prl02) or tc_prl02 is null)"
        'If Not String.IsNullOrEmpty(l_oeb04) Then
        '    oCommand.CommandText += " AND pn like '%" & l_oeb04 & "%' "
        'End If
        'oCommand.CommandText += " order by pn, c1"

        oCommand.CommandText = "Select tc_prm01,ima02,ima021,ima25,C1,(case when tc_prl02 is null and tc_prn02 is null then '3' when tc_prl02 is null then '2' else '1' end) as XS1,"
        oCommand.CommandText += "(case when tc_prl02 is null then tc_prn02 else tc_prl02 end) as C2,(case when tc_prl02 is null then tc_prn06 else tc_prl06 end) as C3,"
        oCommand.CommandText += "(case when tc_prl02 is null then tc_prn03 else tc_prl03 end) as C4,(case when tc_prl02 is null then t1 * tc_prn03 else t1 * tc_prl03 end) as t1,"
        oCommand.CommandText += "(case when tc_prl02 is null then t2 * tc_prn03 else t2 * tc_prl03 end) as t2,(case when tc_prl02 is null then t3 * tc_prn03 else t3 * tc_prl03 end) as t3,"
        oCommand.CommandText += "(case when tc_prl02 is null then t4 * tc_prn03 else t4 * tc_prl03 end) as t4,(case when tc_prl02 is null then t5 * tc_prn03 else t5 * tc_prl03 end) as t5,"
        oCommand.CommandText += "(case when tc_prl02 is null then t6 * tc_prn03 else t6 * tc_prl03 end) as t6,(case when tc_prl02 is null then t7 * tc_prn03 else t7 * tc_prl03 end) as t7,"
        oCommand.CommandText += "(case when tc_prl02 is null then t8 * tc_prn03 else t8 * tc_prl03 end) as t8,(case when tc_prl02 is null then t9 * tc_prn03 else t9 * tc_prl03 end) as t9,"
        oCommand.CommandText += "(case when tc_prl02 is null then t10 * tc_prn03 else t10 * tc_prl03 end) as t10,(case when tc_prl02 is null then t11 * tc_prn03 else t11 * tc_prl03 end) as t11,"
        oCommand.CommandText += "(case when tc_prl02 is null then t12 * tc_prn03 else t12 * tc_prl03 end) as t12 from ( "
        oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01) as c1,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,"
        oCommand.CommandText += "(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2,(case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,"
        oCommand.CommandText += "(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4,(case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6,(case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,"
        oCommand.CommandText += "(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,"
        oCommand.CommandText += "(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        If Not String.IsNullOrEmpty(l_oeb04) Then
            oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
        End If
        oCommand.CommandText += "group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04 order by tc_prm01,max(azn01) ) AB left join tc_prl_file XX1 on AB.tc_prm01 = XX1.tc_prl01 AND AB.C1 <= XX1.tc_prl02  "
        oCommand.CommandText += "left join tc_prn_file XX3 on AB.tc_prm01 = XX3.TC_PRn01 WHERE (XX1.tc_prl02 = (Select min(XX2.tc_prl02) from tc_prl_file XX2 where AB.tc_prm01 = XX2.tc_prl01 and AB.C1 <= XX2.tc_prl02) or tc_prl02 is null) order by tc_prm01,c1"

        oReader = oCommand.ExecuteReader
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
                Ws.Cells(LineZ, 8) = ASS
                Ws.Cells(LineZ, 9) = oReader.Item(4) '交貨日期
                If Not IsDBNull(oReader.Item(5)) Then
                    If oReader.Item(5) = "1" Then
                        Ws.Cells(LineZ, 10) = "大货售价"
                        End If
                    If oReader.Item(5) = "2" Then
                        Ws.Cells(LineZ, 10) = "预估售价"
                        End If
                    End If

                If IsDBNull(oReader.Item(6)) Then
                    Ws.Cells(LineZ, 11) = "有销售无售价"
                Else
                    Ws.Cells(LineZ, 11) = oReader.Item(6)
                    End If
                Ws.Cells(LineZ, 12) = oReader.Item(7)
                Ws.Cells(LineZ, 13) = oReader.Item(8)
                    ' 1月
                Ws.Cells(LineZ, 14) = oReader.Item(9)
                Ws.Cells(LineZ, 15) = oReader.Item(10)
                Ws.Cells(LineZ, 16) = oReader.Item(11)
                Ws.Cells(LineZ, 17) = oReader.Item(12)
                Ws.Cells(LineZ, 18) = oReader.Item(13)
                Ws.Cells(LineZ, 19) = oReader.Item(14)
                Ws.Cells(LineZ, 20) = oReader.Item(15)
                Ws.Cells(LineZ, 21) = oReader.Item(16)
                Ws.Cells(LineZ, 22) = oReader.Item(17)
                Ws.Cells(LineZ, 23) = oReader.Item(18)
                Ws.Cells(LineZ, 24) = oReader.Item(19)
                Ws.Cells(LineZ, 25) = oReader.Item(20)
                    ' 合計
                Ws.Cells(LineZ, 26) = "=SUM(N" & LineZ & ":Y" & LineZ & ")"
                LineZ += 1
                OLDoeb04 = NewOEb04
                End While
            End If
        oReader.Close()
            ' 加總
        Ws.Cells(LineZ, 13) = "Total"
        Ws.Cells(LineZ, 14) = "=SUM(N5:N" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 14), Ws.Cells(LineZ, 14))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 14), Ws.Cells(LineZ, 26)), Type:=xlFillDefault)

            ' 劃線
        oRng = Ws.Range("B3", Ws.Cells(LineZ, 26))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            ' 第三頁
        Ws = xWorkBook.Sheets(3)
        Ws.Name = "Amount"
        Ws.Activate()
        AdjustExcelFormat3()

            'oCommand.CommandText = "select oeb04,ima02,ima021,ima25,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,"
            'oCommand.CommandText += "sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ("
            'oCommand.CommandText += "select oeb04,ima02,ima021,ima25,(case when month(tc_cif_05) = 1 then tc_cif_04 else 0 end ) as t1,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 2 then tc_cif_04 else 0 end ) as t2,(case when month(tc_cif_05) = 3 then tc_cif_04 else 0 end ) as t3,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 4 then tc_cif_04 else 0 end ) as t4,(case when month(tc_cif_05) = 5 then tc_cif_04 else 0 end ) as t5,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 6 then tc_cif_04 else 0 end ) as t6,(case when month(tc_cif_05) = 7 then tc_cif_04 else 0 end ) as t7,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 8 then tc_cif_04 else 0 end ) as t8,(case when month(tc_cif_05) = 9 then tc_cif_04 else 0 end ) as t9,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 10 then tc_cif_04 else 0 end ) as t10,(case when month(tc_cif_05) = 11 then tc_cif_04 else 0 end ) as t11,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 12 then tc_cif_04 else 0 end ) as t12 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            'oCommand.CommandText += "left join ima_file on oeb04 = ima01 where tc_cif_05 between to_date('"
            'oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            'oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 not like 'FC%'"
            'If Not String.IsNullOrEmpty(l_oeb04) Then
            '    oCommand.CommandText += " AND oeb04 like '%" & l_oeb04 & "%' "
            'End If
            'oCommand.CommandText += " union all "
            'oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,(case when month(max(A2.azn01))= 1 then tc_prm04 else 0 end) as t1,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 2 then tc_prm04 else 0 end) as t2,(case when month(max(A2.azn01))= 3 then tc_prm04 else 0 end) as t3,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 4 then tc_prm04 else 0 end) as t4,(case when month(max(A2.azn01))= 5 then tc_prm04 else 0 end) as t5,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 6 then tc_prm04 else 0 end) as t6,(case when month(max(A2.azn01))= 7 then tc_prm04 else 0 end) as t7,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 8 then tc_prm04 else 0 end) as t8,(case when month(max(A2.azn01))= 9 then tc_prm04 else 0 end) as t9,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 10 then tc_prm04 else 0 end) as t10,(case when month(max(A2.azn01))= 11 then tc_prm04 else 0 end) as t11,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 "
            'oCommand.CommandText += "left join (select oeb04,max(A1.azn02) as c1,max(A1.azn05) as c2 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            'oCommand.CommandText += "left join azn_file A1 on tc_cif_05 = A1.azn01  where year(tc_cif_05) = " & tYear & " and tc_cif_01 not like 'FC%' and tc_cif_04 <> 0 "
            'oCommand.CommandText += " group by oeb04 ) X1 ON tc_prm01 = X1.oeb04 left join azn_file A2 on tc_prm02 = A2.azn02 and tc_prm03 = A2.azn05 where tc_prm02 = " & tYear
            'oCommand.CommandText += " And (X1.c2 Is null Or tc_prm03 > X1.C2) "
            'If Not String.IsNullOrEmpty(l_oeb04) Then
            '    oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
            'End If
            'oCommand.CommandText += "group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03, tc_prm04 "
            'oCommand.CommandText += ") group by oeb04,ima02,ima021,ima25"

        'oCommand.CommandText = "Select pn,ima02,ima021,ima25,(case when XX1.tc_prl02 is null then tc_prn06 else tc_prl06 end) as C3,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t1 * nvl(tc_prn03,0) else t1 * nvl(tc_prl03,0) end) as t1,sum(case when tc_prl02 is null then t2 * nvl(tc_prn03,0) else t2 * nvl(tc_prl03,0) end) as t2,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t3 * nvl(tc_prn03,0) else t3 * nvl(tc_prl03,0) end) as t3,sum(case when tc_prl02 is null then t4 * nvl(tc_prn03,0) else t4 * nvl(tc_prl03,0) end) as t4,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t5 * nvl(tc_prn03,0) else t5 * nvl(tc_prl03,0) end) as t5,sum(case when tc_prl02 is null then t6 * nvl(tc_prn03,0) else t6 * nvl(tc_prl03,0) end) as t6,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t7 * nvl(tc_prn03,0) else t7 * nvl(tc_prl03,0) end) as t7,sum(case when tc_prl02 is null then t8 * nvl(tc_prn03,0) else t8 * nvl(tc_prl03,0) end) as t8,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t9 * nvl(tc_prn03,0) else t9 * nvl(tc_prl03,0) end) as t9,sum(case when tc_prl02 is null then t10 * nvl(tc_prn03,0) else t10 * nvl(tc_prl03,0) end) as t10,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t11 * nvl(tc_prn03,0) else t11 * nvl(tc_prl03,0) end) as t11,sum(case when tc_prl02 is null then t12 * nvl(tc_prn03,0) else t12 * nvl(tc_prl03,0) end) as t12 from ( "
        'oCommand.CommandText += "select pn,ima02,ima021,ima25,year1,week1,max(azn01) as c1,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, "
        'oCommand.CommandText += "(case when month(max(azn01)) = 3 then quantity else 0 end) as t3,(case when month(max(azn01)) = 4 then quantity else 0 end) as t4,(case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6,(case when month(max(azn01)) = 7 then quantity else 0 end) as t7,(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join azn_file on year1 = azn02 and week1 = azn05 group by pn,ima02,ima021,ima25,year1,week1,quantity order by pn,max(azn01) "
        'oCommand.CommandText += ") AB left join tc_prl_file XX1 on AB.pn = XX1.tc_prl01 AND AB.C1 <= XX1.tc_prl02  left join tc_prn_file XX3 on AB.pn = XX3.TC_PRn01 WHERE (XX1.tc_prl02 = (Select min(XX2.tc_prl02) from tc_prl_file XX2 where AB.PN = XX2.tc_prl01 and AB.C1 <= XX2.tc_prl02) or tc_prl02 is null) "
        'If Not String.IsNullOrEmpty(l_oeb04) Then
        '    oCommand.CommandText += " AND pn like '%" & l_oeb04 & "%' "
        '    End If
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25, XX1.tc_prl02, XX3.TC_PRN06, XX1.tc_prl06 order by pn"

        oCommand.CommandText = "Select tc_prm01,ima02,ima021,ima25,(case when XX1.tc_prl02 is null then tc_prn06 else tc_prl06 end) as C3,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t1 * nvl(tc_prn03,0) else t1 * nvl(tc_prl03,0) end) as t1,sum(case when tc_prl02 is null then t2 * nvl(tc_prn03,0) else t2 * nvl(tc_prl03,0) end) as t2,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t3 * nvl(tc_prn03,0) else t3 * nvl(tc_prl03,0) end) as t3,sum(case when tc_prl02 is null then t4 * nvl(tc_prn03,0) else t4 * nvl(tc_prl03,0) end) as t4,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t5 * nvl(tc_prn03,0) else t5 * nvl(tc_prl03,0) end) as t5,sum(case when tc_prl02 is null then t6 * nvl(tc_prn03,0) else t6 * nvl(tc_prl03,0) end) as t6,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t7 * nvl(tc_prn03,0) else t7 * nvl(tc_prl03,0) end) as t7,sum(case when tc_prl02 is null then t8 * nvl(tc_prn03,0) else t8 * nvl(tc_prl03,0) end) as t8,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t9 * nvl(tc_prn03,0) else t9 * nvl(tc_prl03,0) end) as t9,sum(case when tc_prl02 is null then t10 * nvl(tc_prn03,0) else t10 * nvl(tc_prl03,0) end) as t10,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t11 * nvl(tc_prn03,0) else t11 * nvl(tc_prl03,0) end) as t11,sum(case when tc_prl02 is null then t12 * nvl(tc_prn03,0) else t12 * nvl(tc_prl03,0) end) as t12 from ( "
        oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,max(azn01) as c1,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, "
        oCommand.CommandText += "(case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4,(case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6,(case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,"
        oCommand.CommandText += "(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "Where tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        If Not String.IsNullOrEmpty(l_oeb04) Then
        oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
        End If
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tc_prm02,tc_prm03,tc_prm04 order by tc_prm01,max(azn01) "
        oCommand.CommandText += ") AB left join tc_prl_file XX1 on AB.tc_prm01 = XX1.tc_prl01 AND AB.C1 <= XX1.tc_prl02  left join tc_prn_file XX3 on AB.tc_prm01 = XX3.TC_PRn01 WHERE (XX1.tc_prl02 = (Select min(XX2.tc_prl02) from tc_prl_file XX2 where AB.tc_prm01 = XX2.tc_prl01 and AB.C1 <= XX2.tc_prl02) or tc_prl02 is null) "
        oCommand.CommandText += "Group by tc_prm01,ima02,ima021,ima25, XX1.tc_prl02, XX3.TC_PRN06, XX1.tc_prl06 order by tc_prm01"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                    '        For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                    '            Ws.Cells(LineZ, 2 + i) = oReader.Item(i)
                    '        Next
                    '        Ws.Cells(LineZ, 18) = "=SUM(F" & LineZ & ":Q" & LineZ & ")"
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 8) = oReader.Item(4)
                    '各月份
                Ws.Cells(LineZ, 9) = oReader.Item(5)
                Ws.Cells(LineZ, 10) = oReader.Item(6)
                Ws.Cells(LineZ, 11) = oReader.Item(7)
                Ws.Cells(LineZ, 12) = oReader.Item(8)
                Ws.Cells(LineZ, 13) = oReader.Item(9)
                Ws.Cells(LineZ, 14) = oReader.Item(10)
                Ws.Cells(LineZ, 15) = oReader.Item(11)
                Ws.Cells(LineZ, 16) = oReader.Item(12)
                Ws.Cells(LineZ, 17) = oReader.Item(13)
                Ws.Cells(LineZ, 18) = oReader.Item(14)
                Ws.Cells(LineZ, 19) = oReader.Item(15)
                Ws.Cells(LineZ, 20) = oReader.Item(16)
                    ' 合計
                Ws.Cells(LineZ, 21) = "=SUM(I" & LineZ & ":T" & LineZ & ")"
                LineZ += 1
                End While
            End If
        oReader.Close()
            ' 加總
        Ws.Cells(LineZ, 8) = "Total"
        Ws.Cells(LineZ, 9) = "=SUM(I5:I" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)

            ' 劃線
        oRng = Ws.Range("B3", Ws.Cells(LineZ, 21))
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            '' 第4頁
            'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
            'Ws = xWorkBook.Sheets(4)
            'Ws.Name = "Calloff+Forecast Amount（Detail）"
            'Ws.Activate()
            'AdjustExcelFormat4()
            ''oCommand.CommandText = "select oeb04,ima02,ima021,ima25,tc_prl02, tc_prl06,tc_prl03,(Z1.t1 * tc_prl03) as t1,(Z1.t2 * tc_prl03) as t2,(Z1.t3 * tc_prl03) as t3,(Z1.t4 * tc_prl03) as t4,(Z1.t5 * tc_prl03) as t5,"
            ''oCommand.CommandText += "(Z1.t6 * tc_prl03) as t6,(Z1.t7 * tc_prl03) as t7,(Z1.t8 * tc_prl03) as t8,(Z1.t9 * tc_prl03) as t9,(Z1.t10 * tc_prl03) as t10,(Z1.t11 * tc_prl03) as t11,(Z1.t12 * tc_prl03) as t12 "
            ''oCommand.CommandText += "from ( "
            ''oCommand.CommandText += "select oeb04,ima02,ima021,ima25,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,"
            ''oCommand.CommandText += "sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ("
            ''oCommand.CommandText += "select oeb04,ima02,ima021,ima25,(case when month(tc_cif_05) = 1 then tc_cif_04 else 0 end ) as t1,"
            ''oCommand.CommandText += "(case when month(tc_cif_05) = 2 then tc_cif_04 else 0 end ) as t2,(case when month(tc_cif_05) = 3 then tc_cif_04 else 0 end ) as t3,"
            ''oCommand.CommandText += "(case when month(tc_cif_05) = 4 then tc_cif_04 else 0 end ) as t4,(case when month(tc_cif_05) = 5 then tc_cif_04 else 0 end ) as t5,"
            ''oCommand.CommandText += "(case when month(tc_cif_05) = 6 then tc_cif_04 else 0 end ) as t6,(case when month(tc_cif_05) = 7 then tc_cif_04 else 0 end ) as t7,"
            ''oCommand.CommandText += "(case when month(tc_cif_05) = 8 then tc_cif_04 else 0 end ) as t8,(case when month(tc_cif_05) = 9 then tc_cif_04 else 0 end ) as t9,"
            ''oCommand.CommandText += "(case when month(tc_cif_05) = 10 then tc_cif_04 else 0 end ) as t10,(case when month(tc_cif_05) = 11 then tc_cif_04 else 0 end ) as t11,"
            ''oCommand.CommandText += "(case when month(tc_cif_05) = 12 then tc_cif_04 else 0 end ) as t12 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            ''oCommand.CommandText += "left join ima_file on oeb04 = ima01 where tc_cif_05 between to_date('"
            ''oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            ''oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 not like 'FC%'"
            ''If Not String.IsNullOrEmpty(l_oeb04) Then
            ''    oCommand.CommandText += " AND oeb04 like '%" & l_oeb04 & "%' "
            ''End If
            ''oCommand.CommandText += " union all "
            ''oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,(case when month(max(A2.azn01))= 1 then tc_prm04 else 0 end) as t1,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 2 then tc_prm04 else 0 end) as t2,(case when month(max(A2.azn01))= 3 then tc_prm04 else 0 end) as t3,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 4 then tc_prm04 else 0 end) as t4,(case when month(max(A2.azn01))= 5 then tc_prm04 else 0 end) as t5,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 6 then tc_prm04 else 0 end) as t6,(case when month(max(A2.azn01))= 7 then tc_prm04 else 0 end) as t7,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 8 then tc_prm04 else 0 end) as t8,(case when month(max(A2.azn01))= 9 then tc_prm04 else 0 end) as t9,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 10 then tc_prm04 else 0 end) as t10,(case when month(max(A2.azn01))= 11 then tc_prm04 else 0 end) as t11,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 "
            ''oCommand.CommandText += "left join (select oeb04,max(A1.azn02) as c1,max(A1.azn05) as c2 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            ''oCommand.CommandText += "left join azn_file A1 on tc_cif_05 = A1.azn01  where year(tc_cif_05) = " & tYear & " and tc_cif_01 not like 'FC%' "
            ''oCommand.CommandText += " group by oeb04 ) X1 ON tc_prm01 = X1.oeb04 left join azn_file A2 on tc_prm02 = A2.azn02 and tc_prm03 = A2.azn05 where tc_prm02 = " & tYear
            ''oCommand.CommandText += " And (X1.c2 Is null Or tc_prm03 > X1.C2) "
            ''If Not String.IsNullOrEmpty(l_oeb04) Then
            ''    oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
            ''End If
            ''oCommand.CommandText += "group by tc_prm01,ima02,ima021,ima25,tc_prm04 "
            ''oCommand.CommandText += ") group by oeb04,ima02,ima021,ima25 "
            ''oCommand.CommandText += ") Z1 left join tc_prl_file on Z1.oeb04 = tc_prl01 and tc_prl02 >= to_date('" & tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') order by oeb04"

            ''oCommand.CommandText = "select oeb04,ima02,ima021,ima25,tc_prl02,tc_prl06,tc_prl03,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,"
            ''oCommand.CommandText += " sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
            ''oCommand.CommandText += " select oeb04,ima02,ima021,ima25,tc_prl02,tc_prl06,tc_prl03,(case when month(tc_cif_05) = 1 then tc_cif_04 * tc_prl03 else 0 end ) as t1,"
            ''oCommand.CommandText += " (case when month(tc_cif_05) = 2 then tc_cif_04 * tc_prl03 else 0 end ) as t2,(case when month(tc_cif_05) = 3 then tc_cif_04 * tc_prl03 else 0 end ) as t3,"
            ''oCommand.CommandText += " (case when month(tc_cif_05) = 4 then tc_cif_04 * tc_prl03 else 0 end ) as t4,(case when month(tc_cif_05) = 5 then tc_cif_04 * tc_prl03 else 0 end ) as t5,"
            ''oCommand.CommandText += " (case when month(tc_cif_05) = 6 then tc_cif_04 * tc_prl03 else 0 end ) as t6,(case when month(tc_cif_05) = 7 then tc_cif_04 * tc_prl03 else 0 end ) as t7,"
            ''oCommand.CommandText += " (case when month(tc_cif_05) = 8 then tc_cif_04 * tc_prl03 else 0 end ) as t8,(case when month(tc_cif_05) = 9 then tc_cif_04 * tc_prl03 else 0 end ) as t9,"
            ''oCommand.CommandText += " (case when month(tc_cif_05) = 10 then tc_cif_04 * tc_prl03 else 0 end ) as t10,(case when month(tc_cif_05) = 11 then tc_cif_04 * tc_prl03 else 0 end ) as t11,"
            ''oCommand.CommandText += " (case when month(tc_cif_05) = 12 then tc_cif_04 * tc_prl03 else 0 end ) as t12 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            ''oCommand.CommandText += "  left join ima_file on oeb04 = ima01  left join tc_prl_file on tc_cif_05 < tc_prl02 and oeb04 = tc_prl01  where tc_cif_05 between to_date('"
            ''oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            ''oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 not like 'FC%' "
            ''oCommand.CommandText += " and (tc_prl02 = (select min(tc_prl02) from tc_prl_file where oeb04 = tc_prl01 and tc_prl02 >= tc_cif_05) or tc_prl02 is null) "
            ''If Not String.IsNullOrEmpty(l_oeb04) Then
            ''    oCommand.CommandText += " AND oeb04 like '%" & l_oeb04 & "%' "
            ''End If
            ''oCommand.CommandText += "union all "
            ''oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,tc_prl02,tc_prl06,tc_prl03,(case when month(max(A2.azn01))= 1 then tc_prm04 * tc_prl03 else 0 end) as t1,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 2 then tc_prm04 * tc_prl03 else 0 end) as t2,(case when month(max(A2.azn01))= 3 then tc_prm04 * tc_prl03 else 0 end) as t3,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 4 then tc_prm04 * tc_prl03 else 0 end) as t4,(case when month(max(A2.azn01))= 5 then tc_prm04 * tc_prl03 else 0 end) as t5,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 6 then tc_prm04 * tc_prl03 else 0 end) as t6,(case when month(max(A2.azn01))= 7 then tc_prm04 * tc_prl03 else 0 end) as t7,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 8 then tc_prm04 * tc_prl03 else 0 end) as t8,(case when month(max(A2.azn01))= 9 then tc_prm04 * tc_prl03 else 0 end) as t9,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 10 then tc_prm04 * tc_prl03 else 0 end) as t10,(case when month(max(A2.azn01))= 11 then tc_prm04 * tc_prl03 else 0 end) as t11,"
            ''oCommand.CommandText += "(case when month(max(A2.azn01))= 12 then tc_prm04 * tc_prl03 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 "
            ''oCommand.CommandText += "left join (select oeb04,max(A1.azn02) as c1,max(A1.azn05) as c2 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            ''oCommand.CommandText += "left join azn_file A1 on tc_cif_05 = A1.azn01  where year(tc_cif_05) = " & tYear & " and tc_cif_01 not like 'FC%' "
            ''oCommand.CommandText += "group by oeb04 ) X1 ON tc_prm01 = X1.oeb04 left join azn_file A2 on tc_prm02 = A2.azn02 and tc_prm03 = A2.azn05 left join tc_prl_file on tc_prm01 = tc_prl01 "
            ''oCommand.CommandText += "where tc_prm02 = " & tYear & " And (X1.c2 Is null Or tc_prm03 > X1.C2)   and (tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prm01 = tc_prl01 and tc_prl02 >= A2.azn01) or tc_prl02 is null) "
            ''If Not String.IsNullOrEmpty(l_oeb04) Then
            ''    oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
            ''End If
            ''oCommand.CommandText += "having max(A2.azn01) < tc_prl02 or tc_prl02 is null  group by tc_prm01,ima02,ima021,ima25,tc_prm03,tc_prm04 ,tc_prl02,tc_prl06,tc_prl03  ) group by oeb04,ima02,ima021,ima25,tc_prl02,tc_prl06,tc_prl03 order by oeb04"
            '' modify by cloud 20191017
            'oCommand.CommandText = "select oeb04,ima02,ima021,ima25,tc_cif_05, xx1,tc_prl02,tc_prl06,tc_prl03,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,"
            'oCommand.CommandText += "sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
            'oCommand.CommandText += "select oeb04,ima02,ima021,ima25,tc_cif_05,'1' as xx1,tc_prl02,tc_prl06,tc_prl03,(case when month(tc_cif_05) = 1 then tc_cif_04 * tc_prl03 else 0 end ) as t1,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 2 then tc_cif_04 * tc_prl03 else 0 end ) as t2,(case when month(tc_cif_05) = 3 then tc_cif_04 * tc_prl03 else 0 end ) as t3,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 4 then tc_cif_04 * tc_prl03 else 0 end ) as t4,(case when month(tc_cif_05) = 5 then tc_cif_04 * tc_prl03 else 0 end ) as t5,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 6 then tc_cif_04 * tc_prl03 else 0 end ) as t6,(case when month(tc_cif_05) = 7 then tc_cif_04 * tc_prl03 else 0 end ) as t7,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 8 then tc_cif_04 * tc_prl03 else 0 end ) as t8,(case when month(tc_cif_05) = 9 then tc_cif_04 * tc_prl03 else 0 end ) as t9,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 10 then tc_cif_04 * tc_prl03 else 0 end ) as t10,(case when month(tc_cif_05) = 11 then tc_cif_04 * tc_prl03 else 0 end ) as t11,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 12 then tc_cif_04 * tc_prl03 else 0 end ) as t12 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            'oCommand.CommandText += "left join ima_file on oeb04 = ima01  left join tc_prl_file on tc_cif_05 < tc_prl02 and oeb04 = tc_prl01 where tc_cif_05 between to_date('"
            'oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            'oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 not like 'FC%' "
            'oCommand.CommandText += "and (tc_prl02 = (select min(tc_prl02) from tc_prl_file where oeb04 = tc_prl01 and tc_prl02 >= tc_cif_05) or tc_prl02 is null) "
            'If Not String.IsNullOrEmpty(l_oeb04) Then
            '    oCommand.CommandText += " AND oeb04 like '%" & l_oeb04 & "%' "
            'End If
            'oCommand.CommandText += "union all "
            'oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,max(A2.azn01),(case when tc_prn01 is null and tc_prl02 is not null then '1' when tc_prn01 is null and tc_prl02 is null then '' when tc_prn01 is not null then '2' end),"
            'oCommand.CommandText += "(case when tc_prl02 is null then tc_prn02 else tc_prl02 end),(case when tc_prl06 is null then tc_prn06 else tc_prl06 end),(case when tc_prl03 is null then tc_prn03 else tc_prl03 end),"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 1 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t1,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 2 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t2,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 3 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t3,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 4 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t4,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 5 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t5,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 6 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t6,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 7 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t7,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 8 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t8,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 9 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t9,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 10 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t10,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 11 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t11,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 12 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t12 "
            'oCommand.CommandText += "from tc_prm_file left join ima_file on tc_prm01 = ima01 left join (select oeb04,max(A1.azn02) as c1,max(A1.azn05) as c2 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            'oCommand.CommandText += "left join azn_file A1 on tc_cif_05 = A1.azn01  where year(tc_cif_05) = " & tYear & " and tc_cif_01 not like 'FC%' and tc_cif_04 <> 0  group by oeb04 ) X1 ON tc_prm01 = X1.oeb04 left join azn_file A2 on tc_prm02 = A2.azn02 and tc_prm03 = A2.azn05 "
            'oCommand.CommandText += "left join tc_prl_file on tc_prm01 = tc_prl01 left join tc_prn_file on tc_prm01 = tc_prn01 "
            'oCommand.CommandText += "where tc_prm02 = " & tYear & " And (X1.c2 Is null Or tc_prm03 > X1.C2)   and (tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prm01 = tc_prl01 and tc_prl02 >= A2.azn01) or tc_prl02 is null) "
            'If Not String.IsNullOrEmpty(l_oeb04) Then
            '    oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
            'End If
            'oCommand.CommandText += "having max(A2.azn01) < tc_prl02 or tc_prl02 is null  group by tc_prm01,ima02,ima021,ima25,tc_prm03,tc_prm04 ,tc_prl02,tc_prl06,tc_prl03,tc_prn01,tc_prn02,tc_prn06,tc_prn03  "
            'oCommand.CommandText += ") group by oeb04,ima02,ima021,ima25,tc_cif_05, xx1,tc_prl02,tc_prl06,tc_prl03 order by oeb04"

            'oReader = oCommand.ExecuteReader()
            'If oReader.HasRows() Then
            '    Dim ASS As Int16 = 1
            '    Dim OLDoeb04 As String = String.Empty
            '    While oReader.Read()
            '        Dim NewOEb04 As String = oReader.Item(0)
            '        Ws.Cells(LineZ, 2) = oReader.Item(0)
            '        Ws.Cells(LineZ, 3) = oReader.Item(1)
            '        Ws.Cells(LineZ, 4) = oReader.Item(2)
            '        Ws.Cells(LineZ, 5) = oReader.Item(3)
            '        If NewOEb04 <> OLDoeb04 Then
            '            ASS = 1
            '        Else
            '            ASS += 1
            '        End If
            '        Ws.Cells(LineZ, 8) = ASS
            '        Ws.Cells(LineZ, 9) = oReader.Item(4) '交貨日期
            '        If Not IsDBNull(oReader.Item(5)) Then
            '            If oReader.Item(5) = "1" Then
            '                Ws.Cells(LineZ, 10) = "大货售价"
            '            End If
            '            If oReader.Item(5) = "2" Then
            '                Ws.Cells(LineZ, 10) = "预估售价"
            '            End If
            '        End If

            '        If IsDBNull(oReader.Item(6)) Then
            '            Ws.Cells(LineZ, 11) = "有销售无售价"
            '        Else
            '            Ws.Cells(LineZ, 11) = oReader.Item(6)
            '        End If
            '        Ws.Cells(LineZ, 12) = oReader.Item(7)
            '        Ws.Cells(LineZ, 13) = oReader.Item(8)
            '        ' 1月
            '        Ws.Cells(LineZ, 14) = oReader.Item(9)
            '        Ws.Cells(LineZ, 15) = oReader.Item(10)
            '        Ws.Cells(LineZ, 16) = oReader.Item(11)
            '        Ws.Cells(LineZ, 17) = oReader.Item(12)
            '        Ws.Cells(LineZ, 18) = oReader.Item(13)
            '        Ws.Cells(LineZ, 19) = oReader.Item(14)
            '        Ws.Cells(LineZ, 20) = oReader.Item(15)
            '        Ws.Cells(LineZ, 21) = oReader.Item(16)
            '        Ws.Cells(LineZ, 22) = oReader.Item(17)
            '        Ws.Cells(LineZ, 23) = oReader.Item(18)
            '        Ws.Cells(LineZ, 24) = oReader.Item(19)
            '        Ws.Cells(LineZ, 25) = oReader.Item(20)
            '        ' 合計
            '        Ws.Cells(LineZ, 26) = "=SUM(N" & LineZ & ":Y" & LineZ & ")"
            '        LineZ += 1
            '        OLDoeb04 = NewOEb04
            '    End While
            'End If
            'oReader.Close()
            '' 加總
            'Ws.Cells(LineZ, 13) = "Total"
            'Ws.Cells(LineZ, 14) = "=SUM(N5:N" & LineZ - 1 & ")"
            'oRng = Ws.Range(Ws.Cells(LineZ, 14), Ws.Cells(LineZ, 14))
            'oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 14), Ws.Cells(LineZ, 26)), Type:=xlFillDefault)

            '' 劃線
            'oRng = Ws.Range("B3", Ws.Cells(LineZ, 26))
            'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

            '' 第5頁
            'Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
            'Ws = xWorkBook.Sheets(5)
            'Ws.Name = "Calloff+Forecast Amount"
            'Ws.Activate()
            'AdjustExcelFormat5()

            'oCommand.CommandText = "select oeb04,ima02,ima021,ima25,tc_prl06,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,"
            'oCommand.CommandText += "sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
            'oCommand.CommandText += "select oeb04,ima02,ima021,ima25,tc_cif_05,'1' as xx1,tc_prl02,tc_prl06,tc_prl03,(case when month(tc_cif_05) = 1 then tc_cif_04 * tc_prl03 else 0 end ) as t1,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 2 then tc_cif_04 * tc_prl03 else 0 end ) as t2,(case when month(tc_cif_05) = 3 then tc_cif_04 * tc_prl03 else 0 end ) as t3,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 4 then tc_cif_04 * tc_prl03 else 0 end ) as t4,(case when month(tc_cif_05) = 5 then tc_cif_04 * tc_prl03 else 0 end ) as t5,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 6 then tc_cif_04 * tc_prl03 else 0 end ) as t6,(case when month(tc_cif_05) = 7 then tc_cif_04 * tc_prl03 else 0 end ) as t7,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 8 then tc_cif_04 * tc_prl03 else 0 end ) as t8,(case when month(tc_cif_05) = 9 then tc_cif_04 * tc_prl03 else 0 end ) as t9,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 10 then tc_cif_04 * tc_prl03 else 0 end ) as t10,(case when month(tc_cif_05) = 11 then tc_cif_04 * tc_prl03 else 0 end ) as t11,"
            'oCommand.CommandText += "(case when month(tc_cif_05) = 12 then tc_cif_04 * tc_prl03 else 0 end ) as t12 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            'oCommand.CommandText += "left join ima_file on oeb04 = ima01  left join tc_prl_file on tc_cif_05 < tc_prl02 and oeb04 = tc_prl01 where tc_cif_05 between to_date('"
            'oCommand.CommandText += tDate1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and to_date('"
            'oCommand.CommandText += tDate2.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and tc_cif_01 not like 'FC%' "
            'oCommand.CommandText += "and (tc_prl02 = (select min(tc_prl02) from tc_prl_file where oeb04 = tc_prl01 and tc_prl02 >= tc_cif_05) or tc_prl02 is null) "
            'If Not String.IsNullOrEmpty(l_oeb04) Then
            '    oCommand.CommandText += " AND oeb04 like '%" & l_oeb04 & "%' "
            'End If
            'oCommand.CommandText += "union all "
            'oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,max(A2.azn01),(case when tc_prn01 is null and tc_prl02 is not null then '1' when tc_prn01 is null and tc_prl02 is null then '' when tc_prn01 is not null then '2' end),"
            'oCommand.CommandText += "(case when tc_prl02 is null then tc_prn02 else tc_prl02 end),(case when tc_prl06 is null then tc_prn06 else tc_prl06 end),(case when tc_prl03 is null then tc_prn03 else tc_prl03 end),"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 1 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t1,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 2 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t2,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 3 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t3,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 4 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t4,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 5 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t5,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 6 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t6,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 7 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t7,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 8 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t8,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 9 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t9,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 10 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t10,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 11 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t11,"
            'oCommand.CommandText += "(case when month(max(A2.azn01))= 12 then tc_prm04 * (case when tc_prl03 is null then tc_prn03 else tc_prl03 end) else 0 end) as t12 "
            'oCommand.CommandText += "from tc_prm_file left join ima_file on tc_prm01 = ima01 left join (select oeb04,max(A1.azn02) as c1,max(A1.azn05) as c2 from tc_cif_file left join oeb_file on tc_cif_01 = oeb01 and tc_cif_02 = oeb03 "
            'oCommand.CommandText += "left join azn_file A1 on tc_cif_05 = A1.azn01  where year(tc_cif_05) = " & tYear & " and tc_cif_01 not like 'FC%' group by oeb04 ) X1 ON tc_prm01 = X1.oeb04 left join azn_file A2 on tc_prm02 = A2.azn02 and tc_prm03 = A2.azn05 "
            'oCommand.CommandText += "left join tc_prl_file on tc_prm01 = tc_prl01 left join tc_prn_file on tc_prm01 = tc_prn01 "
            'oCommand.CommandText += "where tc_prm02 = " & tYear & " And (X1.c2 Is null Or tc_prm03 > X1.C2)   and (tc_prl02 = (select min(tc_prl02) from tc_prl_file where tc_prm01 = tc_prl01 and tc_prl02 >= A2.azn01) or tc_prl02 is null) "
            'If Not String.IsNullOrEmpty(l_oeb04) Then
            '    oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
            'End If
            'oCommand.CommandText += "having max(A2.azn01) < tc_prl02 or tc_prl02 is null  group by tc_prm01,ima02,ima021,ima25,tc_prm03,tc_prm04 ,tc_prl02,tc_prl06,tc_prl03,tc_prn01,tc_prn02,tc_prn06,tc_prn03  "
            'oCommand.CommandText += ") group by oeb04,ima02,ima021,ima25,tc_prl06 order by oeb04"

            'oReader = oCommand.ExecuteReader()
            'If oReader.HasRows() Then
            '    While oReader.Read()
            '        Ws.Cells(LineZ, 2) = oReader.Item(0)
            '        Ws.Cells(LineZ, 3) = oReader.Item(1)
            '        Ws.Cells(LineZ, 4) = oReader.Item(2)
            '        Ws.Cells(LineZ, 5) = oReader.Item(3)
            '        Ws.Cells(LineZ, 8) = oReader.Item(4) '交貨日期
            '        ' 1月
            '        Ws.Cells(LineZ, 9) = oReader.Item(5)
            '        Ws.Cells(LineZ, 10) = oReader.Item(6)
            '        Ws.Cells(LineZ, 11) = oReader.Item(7)
            '        Ws.Cells(LineZ, 12) = oReader.Item(8)
            '        Ws.Cells(LineZ, 13) = oReader.Item(9)
            '        Ws.Cells(LineZ, 14) = oReader.Item(10)
            '        Ws.Cells(LineZ, 15) = oReader.Item(11)
            '        Ws.Cells(LineZ, 16) = oReader.Item(12)
            '        Ws.Cells(LineZ, 17) = oReader.Item(13)
            '        Ws.Cells(LineZ, 18) = oReader.Item(14)
            '        Ws.Cells(LineZ, 19) = oReader.Item(15)
            '        Ws.Cells(LineZ, 20) = oReader.Item(16)
            '        ' 合計
            '        Ws.Cells(LineZ, 21) = "=SUM(I" & LineZ & ":T" & LineZ & ")"
            '        LineZ += 1

            '    End While
            'End If
            'oReader.Close()
            '' 加總
            'Ws.Cells(LineZ, 8) = "Total"
            'Ws.Cells(LineZ, 9) = "=SUM(I5:I" & LineZ - 1 & ")"
            'oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 9))
            'oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ, 21)), Type:=xlFillDefault)

            '' 劃線
            'oRng = Ws.Range("B3", Ws.Cells(LineZ, 21))
            'oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            'oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            'oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            'oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            'oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            'oRng.Borders(xlInsideVertical).LineStyle = xlContinuous


            ' 第四頁  20191031
        Ws = xWorkBook.Sheets.Add(After:=xWorkBook.Sheets(xWorkBook.Sheets.Count))
        Ws = xWorkBook.Sheets(4)
        Ws.Name = "Amount  USD"
        Ws.Activate()
        AdjustExcelFormat4()

        'oCommand.CommandText = "Select pn,ima02,ima021,ima25,tqa02,Round(xy.t1 * (xz.er/xzz.er),3) as t1,Round(xy.t2 * (xz.er/xzz.er),3) as t2,Round(xy.t3 * (xz.er/xzz.er),3) as t3,Round(xy.t4 * (xz.er/xzz.er),3) as t4,"
        'oCommand.CommandText += "Round(xy.t5 * (xz.er/xzz.er),3) as t5,Round(xy.t6 * (xz.er/xzz.er),3) as t6,Round(xy.t7 * (xz.er/xzz.er),3) as t7,Round(xy.t8 * (xz.er/xzz.er),3) as t8,Round(xy.t9 * (xz.er/xzz.er),3) as t9,"
        'oCommand.CommandText += "Round(xy.t10 * (xz.er/xzz.er),3) as t10,Round(xy.t11 * (xz.er/xzz.er),3) as t11,Round(xy.t12 * (xz.er/xzz.er),3) as t12 from ( "
        'oCommand.CommandText += "Select pn,ima02,ima021,ima25,tqa02,(case when XX1.tc_prl02 is null then tc_prn06 else tc_prl06 end) as C3,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t1 * nvl(tc_prn03,0) else t1 * nvl(tc_prl03,0) end) as t1,sum(case when tc_prl02 is null then t2 * nvl(tc_prn03,0) else t2 * nvl(tc_prl03,0) end) as t2,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t3 * nvl(tc_prn03,0) else t3 * nvl(tc_prl03,0) end) as t3,sum(case when tc_prl02 is null then t4 * nvl(tc_prn03,0) else t4 * nvl(tc_prl03,0) end) as t4,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t5 * nvl(tc_prn03,0) else t5 * nvl(tc_prl03,0) end) as t5,sum(case when tc_prl02 is null then t6 * nvl(tc_prn03,0) else t6 * nvl(tc_prl03,0) end) as t6,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t7 * nvl(tc_prn03,0) else t7 * nvl(tc_prl03,0) end) as t7,sum(case when tc_prl02 is null then t8 * nvl(tc_prn03,0) else t8 * nvl(tc_prl03,0) end) as t8,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t9 * nvl(tc_prn03,0) else t9 * nvl(tc_prl03,0) end) as t9,sum(case when tc_prl02 is null then t10 * nvl(tc_prn03,0) else t10 * nvl(tc_prl03,0) end) as t10,"
        'oCommand.CommandText += "sum(case when tc_prl02 is null then t11 * nvl(tc_prn03,0) else t11 * nvl(tc_prl03,0) end) as t11,sum(case when tc_prl02 is null then t12 * nvl(tc_prn03,0) else t12 * nvl(tc_prl03,0) end) as t12 from ( "
        'oCommand.CommandText += "select pn,ima02,ima021,ima25,tqa02,year1,week1,max(azn01) as c1,(case when month(max(azn01)) = 1 then quantity else 0 end) as t1,(case when month(max(azn01)) = 2 then quantity else 0 end) as t2, "
        'oCommand.CommandText += "(case when month(max(azn01)) = 3 then quantity else 0 end) as t3,(case when month(max(azn01)) = 4 then quantity else 0 end) as t4,(case when month(max(azn01)) = 5 then quantity else 0 end) as t5,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 6 then quantity else 0 end) as t6,(case when month(max(azn01)) = 7 then quantity else 0 end) as t7,(case when month(max(azn01)) = 8 then quantity else 0 end) as t8,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 9 then quantity else 0 end) as t9,(case when month(max(azn01)) = 10 then quantity else 0 end) as t10,(case when month(max(azn01)) = 11 then quantity else 0 end) as t11,"
        'oCommand.CommandText += "(case when month(max(azn01)) = 12 then quantity else 0 end) as t12 from budget2020 left join ima_file on pn = ima01 left join tqa_file on tqa03 = 2 and ima1005 = tqa01 left join azn_file on year1 = azn02 and week1 = azn05 group by pn,ima02,ima021,ima25,tqa02,year1,week1,quantity order by pn,max(azn01) "
        'oCommand.CommandText += ") AB left join tc_prl_file XX1 on AB.pn = XX1.tc_prl01 AND AB.C1 <= XX1.tc_prl02  left join tc_prn_file XX3 on AB.pn = XX3.TC_PRn01 WHERE (XX1.tc_prl02 = (Select min(XX2.tc_prl02) from tc_prl_file XX2 where AB.PN = XX2.tc_prl01 and AB.C1 <= XX2.tc_prl02) or tc_prl02 is null) "
        'If Not String.IsNullOrEmpty(l_oeb04) Then
        '    oCommand.CommandText += " AND pn like '%" & l_oeb04 & "%' "
        '    End If
        'oCommand.CommandText += "group by pn,ima02,ima021,ima25,tqa02, XX1.tc_prl02, XX3.TC_PRN06, XX1.tc_prl06 order by pn "
        'oCommand.CommandText += ") XY  left join exchangeratebyyear xz on xz.year1 = " & tYear & " and xy.c3 = currency left join exchangeratebyyear xzz on xzz.year1 = " & tYear & " and xzz.currency = 'USD'"

        oCommand.CommandText = "Select tc_prm01,ima02,ima021,ima25,tqa02,Round(xy.t1 * (xz.er/xzz.er),3) as t1,Round(xy.t2 * (xz.er/xzz.er),3) as t2,Round(xy.t3 * (xz.er/xzz.er),3) as t3,Round(xy.t4 * (xz.er/xzz.er),3) as t4,"
        oCommand.CommandText += "Round(xy.t5 * (xz.er/xzz.er),3) as t5,Round(xy.t6 * (xz.er/xzz.er),3) as t6,Round(xy.t7 * (xz.er/xzz.er),3) as t7,Round(xy.t8 * (xz.er/xzz.er),3) as t8,Round(xy.t9 * (xz.er/xzz.er),3) as t9,"
        oCommand.CommandText += "Round(xy.t10 * (xz.er/xzz.er),3) as t10,Round(xy.t11 * (xz.er/xzz.er),3) as t11,Round(xy.t12 * (xz.er/xzz.er),3) as t12 from ( "
        oCommand.CommandText += "Select tc_prm01,ima02,ima021,ima25,tqa02,(case when XX1.tc_prl02 is null then tc_prn06 else tc_prl06 end) as C3,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t1 * nvl(tc_prn03,0) else t1 * nvl(tc_prl03,0) end) as t1,sum(case when tc_prl02 is null then t2 * nvl(tc_prn03,0) else t2 * nvl(tc_prl03,0) end) as t2,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t3 * nvl(tc_prn03,0) else t3 * nvl(tc_prl03,0) end) as t3,sum(case when tc_prl02 is null then t4 * nvl(tc_prn03,0) else t4 * nvl(tc_prl03,0) end) as t4,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t5 * nvl(tc_prn03,0) else t5 * nvl(tc_prl03,0) end) as t5,sum(case when tc_prl02 is null then t6 * nvl(tc_prn03,0) else t6 * nvl(tc_prl03,0) end) as t6,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t7 * nvl(tc_prn03,0) else t7 * nvl(tc_prl03,0) end) as t7,sum(case when tc_prl02 is null then t8 * nvl(tc_prn03,0) else t8 * nvl(tc_prl03,0) end) as t8,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t9 * nvl(tc_prn03,0) else t9 * nvl(tc_prl03,0) end) as t9,sum(case when tc_prl02 is null then t10 * nvl(tc_prn03,0) else t10 * nvl(tc_prl03,0) end) as t10,"
        oCommand.CommandText += "sum(case when tc_prl02 is null then t11 * nvl(tc_prn03,0) else t11 * nvl(tc_prl03,0) end) as t11,sum(case when tc_prl02 is null then t12 * nvl(tc_prn03,0) else t12 * nvl(tc_prl03,0) end) as t12 from ( "
        oCommand.CommandText += "select tc_prm01,ima02,ima021,ima25,tqa02,tc_prm02,tc_prm03,max(azn01) as c1,(case when month(max(azn01)) = 1 then tc_prm04 else 0 end) as t1,(case when month(max(azn01)) = 2 then tc_prm04 else 0 end) as t2, "
        oCommand.CommandText += "(case when month(max(azn01)) = 3 then tc_prm04 else 0 end) as t3,(case when month(max(azn01)) = 4 then tc_prm04 else 0 end) as t4,(case when month(max(azn01)) = 5 then tc_prm04 else 0 end) as t5,"
        oCommand.CommandText += "(case when month(max(azn01)) = 6 then tc_prm04 else 0 end) as t6,(case when month(max(azn01)) = 7 then tc_prm04 else 0 end) as t7,(case when month(max(azn01)) = 8 then tc_prm04 else 0 end) as t8,"
        oCommand.CommandText += "(case when month(max(azn01)) = 9 then tc_prm04 else 0 end) as t9,(case when month(max(azn01)) = 10 then tc_prm04 else 0 end) as t10,(case when month(max(azn01)) = 11 then tc_prm04 else 0 end) as t11,"
        oCommand.CommandText += "(case when month(max(azn01)) = 12 then tc_prm04 else 0 end) as t12 from tc_prm_file left join ima_file on tc_prm01 = ima01 left join tqa_file on tqa03 = 2 and ima1005 = tqa01 left join azn_file on tc_prm02 = azn02 and tc_prm03 = azn05 "
        oCommand.CommandText += "WHERE tc_prmlegal = 'ACTIONTEST' and tc_prm02 = " & tYear
        If Not String.IsNullOrEmpty(l_oeb04) Then
            oCommand.CommandText += " AND tc_prm01 like '%" & l_oeb04 & "%' "
        End If
        oCommand.CommandText += "group by tc_prm01,ima02,ima021,ima25,tqa02,tc_prm02,tc_prm03,tc_prm04 order by tc_prm01,max(azn01) "
        oCommand.CommandText += " ) AB left join tc_prl_file XX1 on AB.tc_prm01 = XX1.tc_prl01 AND AB.C1 <= XX1.tc_prl02  left join tc_prn_file XX3 on AB.tc_prm01 = XX3.TC_PRn01 WHERE (XX1.tc_prl02 = (Select min(XX2.tc_prl02) from tc_prl_file XX2 where AB.tc_prm01 = XX2.tc_prl01 and AB.C1 <= XX2.tc_prl02) or tc_prl02 is null) "
        oCommand.CommandText += " group by tc_prm01,ima02,ima021,ima25,tqa02, XX1.tc_prl02, XX3.TC_PRN06, XX1.tc_prl06 order by tc_prm01 "
        oCommand.CommandText += " ) XY  left join exchangeratebyyear xz on xz.year1 = " & tYear & " and xy.c3 = currency left join exchangeratebyyear xzz on xzz.year1 = " & tYear & " and xzz.currency = 'USD'"

        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                        '        For i As Int16 = 0 To oReader.FieldCount - 1 Step 1
                        '            Ws.Cells(LineZ, 2 + i) = oReader.Item(i)
                        '        Next
                        '        Ws.Cells(LineZ, 18) = "=SUM(F" & LineZ & ":Q" & LineZ & ")"
                Ws.Cells(LineZ, 2) = oReader.Item(0)
                Ws.Cells(LineZ, 3) = oReader.Item(1)
                Ws.Cells(LineZ, 4) = oReader.Item(2)
                Ws.Cells(LineZ, 5) = oReader.Item(3)
                Ws.Cells(LineZ, 7) = oReader.Item(4)
                        '各月份
                Ws.Cells(LineZ, 8) = oReader.Item(5)
                Ws.Cells(LineZ, 9) = oReader.Item(6)
                Ws.Cells(LineZ, 10) = oReader.Item(7)
                Ws.Cells(LineZ, 11) = oReader.Item(8)
                Ws.Cells(LineZ, 12) = oReader.Item(9)
                Ws.Cells(LineZ, 13) = oReader.Item(10)
                Ws.Cells(LineZ, 14) = oReader.Item(11)
                Ws.Cells(LineZ, 15) = oReader.Item(12)
                Ws.Cells(LineZ, 16) = oReader.Item(13)
                Ws.Cells(LineZ, 17) = oReader.Item(14)
                Ws.Cells(LineZ, 18) = oReader.Item(15)
                Ws.Cells(LineZ, 19) = oReader.Item(16)
                        ' 合計
                Ws.Cells(LineZ, 20) = "=SUM(H" & LineZ & ":S" & LineZ & ")"
                LineZ += 1
                    End While
                End If
        oReader.Close()
                ' 加總
                'Ws.Cells(LineZ, 8) = "Total"
        Ws.Cells(LineZ, 8) = "=SUM(H5:H" & LineZ - 1 & ")"
        oRng = Ws.Range(Ws.Cells(LineZ, 8), Ws.Cells(LineZ, 8))
        oRng.AutoFill(Destination:=Ws.Range(Ws.Cells(LineZ, 8), Ws.Cells(LineZ, 20)), Type:=xlFillDefault)

                ' 劃線
        oRng = Ws.Range("B3", Ws.Cells(LineZ, 20))
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
        Ws.Cells(3, 2) = "ERP料号"
        Ws.Cells(3, 3) = "品名"
        Ws.Cells(3, 4) = "规格"
        Ws.Cells(3, 5) = "单位"
        Ws.Cells(3, 6) = "年"
        Ws.Cells(3, 7) = "周"
        Ws.Cells(3, 8) = "交货日期"
        Ws.Cells(4, 2) = "Part No."
        Ws.Cells(4, 3) = "Part Name"
        Ws.Cells(4, 4) = "Spec."
        Ws.Cells(4, 5) = "Unit"
        Ws.Cells(4, 6) = "Year"
        Ws.Cells(4, 7) = "Week"
        Ws.Cells(4, 8) = "Delivery Date"
        oRng = Ws.Range("I1", "U1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "
        oRng = Ws.Range("I3", "T3")
        oRng.NumberFormat = "[$-en-US]mmm/yy;@"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(3, 8 + i) = tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(3, 21) = "合计"
        Ws.Cells(4, 21) = "Total"
        LineZ = 5
    End Sub

    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "ERP料号"
        Ws.Cells(3, 3) = "品名"
        Ws.Cells(3, 4) = "规格"
        Ws.Cells(3, 5) = "单位"
        Ws.Cells(3, 6) = "销售区域"
        Ws.Cells(3, 7) = "销售客户"
        Ws.Cells(3, 8) = "项次"
        Ws.Cells(3, 9) = "交货日期"
        Ws.Cells(3, 10) = "备注"
        Ws.Cells(3, 11) = "截止日期"
        Ws.Cells(3, 12) = "销售币别"
        Ws.Cells(3, 13) = "销售售价"
        Ws.Cells(4, 2) = "Part No."
        Ws.Cells(4, 3) = "Part Name"
        Ws.Cells(4, 4) = "Spec."
        Ws.Cells(4, 5) = "Unit"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(4, 7) = "Customer"
        Ws.Cells(4, 8) = "Positi"
        Ws.Cells(4, 9) = "Delivery Date"
        Ws.Cells(4, 10) = "Remark"
        Ws.Cells(4, 11) = "Closing Date"
        Ws.Cells(4, 12) = "Currency"
        Ws.Cells(4, 13) = "HAC Sales"

        oRng = Ws.Range("N1", "Z1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "
        oRng = Ws.Range("N3", "Y3")
        oRng.NumberFormat = "[$-en-US]mmm/yy;@"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(3, 13 + i) = tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(3, 26) = "合计"
        Ws.Cells(4, 26) = "Total"
        LineZ = 5

    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "ERP料号"
        Ws.Cells(3, 3) = "品名"
        Ws.Cells(3, 4) = "规格"
        Ws.Cells(3, 5) = "单位"
        Ws.Cells(3, 6) = "销售区域"
        Ws.Cells(3, 7) = "销售客户"
        Ws.Cells(3, 8) = "销售币别"
        Ws.Cells(4, 2) = "Part No."
        Ws.Cells(4, 3) = "Part Name"
        Ws.Cells(4, 4) = "Spec."
        Ws.Cells(4, 5) = "Unit"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(4, 7) = "Customer"
        Ws.Cells(4, 8) = "Currency"
        oRng = Ws.Range("I1", "U1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "
        oRng = Ws.Range("I3", "T3")
        oRng.NumberFormat = "[$-en-US]mmm/yy;@"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(3, 8 + i) = tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(3, 21) = "合计"
        Ws.Cells(4, 21) = "Total"
        LineZ = 5
    End Sub
    Private Sub AdjustExcelFormat4()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "ERP料号"
        Ws.Cells(3, 3) = "品名"
        Ws.Cells(3, 4) = "规格"
        Ws.Cells(3, 5) = "单位"
        Ws.Cells(3, 6) = "销售区域"
        Ws.Cells(3, 7) = "销售客户"
        Ws.Cells(4, 2) = "Part No."
        Ws.Cells(4, 3) = "Part Name"
        Ws.Cells(4, 4) = "Spec."
        Ws.Cells(4, 5) = "Unit"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(4, 7) = "Customer"
        oRng = Ws.Range("H1", "T1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "

        oRng = Ws.Range("H3", "S3")
        oRng.NumberFormat = "[$-en-US]mmm/yy;@"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(3, 7 + i) = tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(3, 20) = "合计"
        Ws.Cells(4, 20) = "Total"

        oRng = Ws.Range("H2", "T2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng.Interior.Color = Color.Orange
        Ws.Cells(2, 8) = "此表数据以HAC 售价，如果DAC的售价金额，以此表x 98%"

        LineZ = 5
    End Sub
    Private Sub AdjustExcelFormat5()
        xExcel.ActiveWindow.DisplayGridlines = False
        xExcel.ActiveWindow.Zoom = 75
        Ws.Cells(2, 2) = "Company Name：DAC"
        Ws.Cells(3, 2) = "ERP料号"
        Ws.Cells(3, 3) = "品名"
        Ws.Cells(3, 4) = "规格"
        Ws.Cells(3, 5) = "单位"
        Ws.Cells(3, 6) = "销售区域"
        Ws.Cells(3, 7) = "销售客户"
        Ws.Cells(3, 8) = "销售币别"

        Ws.Cells(4, 2) = "Part No."
        Ws.Cells(4, 3) = "Part Name"
        Ws.Cells(4, 4) = "Spec."
        Ws.Cells(4, 5) = "Unit"
        Ws.Cells(4, 6) = "Area"
        Ws.Cells(4, 7) = "Customer"
        Ws.Cells(4, 8) = "Currency"

        oRng = Ws.Range("N1", "Z1")
        oRng.EntireColumn.NumberFormat = "#,##0.00_ "

        oRng = Ws.Range("I3", "T3")
        oRng.NumberFormat = "[$-en-US]mmm/yy;@"
        For i As Int16 = 1 To 12 Step 1
            Ws.Cells(3, 8 + i) = tDate1.AddMonths(i - 1).ToString("yyyy/MM/dd")
        Next
        Ws.Cells(3, 21) = "合计"
        Ws.Cells(4, 21) = "Total"
        LineZ = 5
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
            If Not String.IsNullOrEmpty(l_oeb04) Then
                oCommand.CommandText += " AND pn like '%" & l_oeb04 & "%' "
            End If
            oCommand.CommandText += " group by pn"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        Next

    End Sub
End Class