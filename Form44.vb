Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form44
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim ptime As String = String.Empty
    Dim MaxDetailCount As Int16 = 0
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim HaveReport As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form44_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1).AddSeconds(-1)
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        ' 條件
        'mSQLS1.CommandText = "select model,defect,desc_th,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6 from ( "
        'mSQLS1.CommandText += "select lot.model,cf01,scrap.defect,defect.desc_th,(case when right(cf01,2) in ('31','32','33','35') then 1 else 0 end) as t1,"
        'mSQLS1.CommandText += "(case when right(cf01,2) in ('36') then 1 else 0 end) as t2,(case when right(cf01,2) in ('61') then 1 else 0 end) as t3,"
        'mSQLS1.CommandText += "(case when right(cf01,2) in ('64') then 1 else 0 end) as t4,(case when right(cf01,2) in ('65','66') then 1 else 0 end) as t5,"
        'mSQLS1.CommandText += "(case when right(cf01,2) IS NULL or right(cf01,2) not in ('31','32','33','35','36','61','64','65','66') then 1 else 0 end) as t6 from scrap "
        'mSQLS1.CommandText += "left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot "
        'mSQLS1.CommandText += "left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' "
        'mSQLS1.CommandText += "and model_station_paravalue.station = scrap_sn.updatedstation and model_station_paravalue.model = lot.model "
        'mSQLS1.CommandText += "left join defect on scrap.defect = defect.defect where scrap.datetime between '"
        'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        'mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "') AS CC group by model,defect,desc_th"
        'mSQLS1.CommandText = "select cf01,model,defect,desc_en,desc_th,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t6a) as t6a,sum(t6b) as t6b,"
        'mSQLS1.CommandText += "sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11 from ( "
        'mSQLS1.CommandText += "select lot.model,cf01,scrap.defect,defect.desc_en,defect.desc_th,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '31' then 1 else 0 end) as t1,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '32' or right(cf01,3) = '32A' then 1 else 0 end) as t2,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '35' or right(cf01,3) = '35A' then 1 else 0 end) as t3,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '36' or right(cf01,3) = '36A' then 1 else 0 end) as t4,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '61' or right(cf01,3) = '61A' then 1 else 0 end) as t5,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '64'  then 1 else 0 end) as t6,"
        'mSQLS1.CommandText += "(case when right(cf01,3) = '64A' then 1 else 0 end) as t6a,"
        'mSQLS1.CommandText += "(case when right(cf01,3) = '64B' then 1 else 0 end) as t6b,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '63' or right(cf01,3) = '63A' then 1 else 0 end) as t7,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '65' or right(cf01,3) = '65A' then 1 else 0 end) as t8,"
        'mSQLS1.CommandText += "(case when right(cf01,2) = '66' or right(cf01,3) = '66A' then 1 else 0 end) as t9,"
        'mSQLS1.CommandText += "(case when scrap_sn.updatedstation = '9999' then 1 else 0 end) as t10,"
        'mSQLS1.CommandText += "(case when right(cf01,2) IS NULL and scrap_sn.updatedstation <> '9999' then 1 else 0 end) as t11 "
        mSQLS1.CommandText = "select cf01,model,defect,desc_en,desc_th,cf01A,StationA,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t6a) as t6a,sum(t6b) as t6b,"
        mSQLS1.CommandText += "sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,sum(t11) as t11,sum(t12) as t12 from ( "
        mSQLS1.CommandText += "select lot.model,cf01,scrap.defect,defect.desc_en,defect.desc_th,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0170','0479','0480','0650') then cf01 else '' end) as cf01A,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0170','0479','0480','0650') then scrap_sn.updatedstation else '' end) as StationA,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '31' then 1 else 0 end) as t1,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '32' or right(cf01,3) = '32A' then 1 else 0 end) as t2,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '35' or right(cf01,3) = '35A' then 1 else 0 end) as t3,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '36' or right(cf01,3) = '36A' then 1 else 0 end) as t4,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '61' or right(cf01,3) = '61A' then 1 else 0 end) as t5,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '64'  then 1 else 0 end) as t6,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,3) = '64A' then 1 else 0 end) as t6a,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,3) = '64B' then 1 else 0 end) as t6b,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '63' or right(cf01,3) = '63A' then 1 else 0 end) as t7,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '65' or right(cf01,3) = '65A' then 1 else 0 end) as t8,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation not in ('0170','0479','0480','0650') and right(cf01,2) = '66' or right(cf01,3) = '66A' then 1 else 0 end) as t9,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation = '9999' then 1 else 0 end) as t10, (case when right(cf01,2) IS NULL and scrap_sn.updatedstation <> '9999' then 1 else 0 end) as t11 ,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0170','0479','0480','0650') then 1 else 0 end ) as t12 "
        mSQLS1.CommandText += "from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue on model_station_paravalue.profilename = 'ERP' and model_station_paravalue.station = scrap_sn.updatedstation and model_station_paravalue.model = lot.model "
            'mSQLS1.CommandText += "left join defect on scrap.defect = defect.defect where scrap.defect not in ('DL02','DL01','DL03','DL04','DL12','DL13') and scrap.datetime between '"
        mSQLS1.CommandText += "left join defect on scrap.defect = defect.defect where scrap.defect not in ('DJ01','DJ02','DL01','DL02','DL03','DL04','DL05', 'DL07','DL08', 'DL09', 'DL12', 'DL13') and scrap.datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "') AS CC group by model,defect,desc_en,desc_th,cf01,cf01A, StationA"
        mSQLReader = mSQLS1.ExecuteReader(CommandBehavior.CloseConnection)
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("defect") & " " & mSQLReader.Item("desc_en") & " " & mSQLReader.Item("desc_th")
                If mSQLReader.Item("t1") > 0 Then
                    Ws.Cells(LineZ, 3) = mSQLReader.Item("t1")
                    End If
                If mSQLReader.Item("t2") > 0 Then
                    Ws.Cells(LineZ, 4) = mSQLReader.Item("t2")
                    End If
                If mSQLReader.Item("t3") > 0 Then
                    Ws.Cells(LineZ, 5) = mSQLReader.Item("t3")
                    Ws.Cells(LineZ, 6) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t4") > 0 Then
                    Ws.Cells(LineZ, 7) = mSQLReader.Item("t4")
                    Ws.Cells(LineZ, 8) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t5") > 0 Then
                    Ws.Cells(LineZ, 9) = mSQLReader.Item("t5")
                    Ws.Cells(LineZ, 10) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t6") > 0 Then
                    Ws.Cells(LineZ, 11) = mSQLReader.Item("t6")
                    Ws.Cells(LineZ, 12) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t6a") > 0 Then
                    Ws.Cells(LineZ, 13) = mSQLReader.Item("t6a")
                    Ws.Cells(LineZ, 14) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t6b") > 0 Then
                    Ws.Cells(LineZ, 15) = mSQLReader.Item("t6b")
                    Ws.Cells(LineZ, 16) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t7") > 0 Then
                    Ws.Cells(LineZ, 17) = mSQLReader.Item("t7")
                    Ws.Cells(LineZ, 18) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t8") > 0 Then
                    Ws.Cells(LineZ, 19) = mSQLReader.Item("t8")
                    Ws.Cells(LineZ, 20) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t9") > 0 Then
                    Ws.Cells(LineZ, 21) = mSQLReader.Item("t9")
                    Ws.Cells(LineZ, 22) = mSQLReader.Item("cf01")
                    End If
                If mSQLReader.Item("t10") > 0 Then
                    Ws.Cells(LineZ, 23) = mSQLReader.Item("t10")
                    End If
                If mSQLReader.Item("t11") > 0 Then
                    Ws.Cells(LineZ, 24) = mSQLReader.Item("t11")
                    End If
                Ws.Cells(LineZ, 25) = "=SUM(C" & LineZ & ":X" & LineZ & ")"
                Dim G1 As Integer = Get0330Data(mSQLReader.Item("model"), "0330")
                Ws.Cells(LineZ, 26) = G1
                Dim G2 As Integer = Get0330Data(mSQLReader.Item("model"), "0331")
                Ws.Cells(LineZ, 27) = G2
                    'If G1 > 0 Then
                    'Ws.Cells(LineZ, 19) = "=N" & LineZ & "/" & G1
                    'ElseIf G2 > 0 Then
                    'Ws.Cells(LineZ, 19) = "=N" & LineZ & "/" & G2
                    'Else

                'End If
                If mSQLReader.Item("t12") > 0 Then
                    Ws.Cells(LineZ, 28) = mSQLReader.Item("t12")
                End If
                Ws.Cells(LineZ, 29) = mSQLReader.Item("cf01A")
                Ws.Cells(LineZ, 30) = mSQLReader.Item("StationA")
                LineZ += 1
                End While
            End If
        mSQLReader.Close()
        Ws.Cells(LineZ, 1) = "TOTAL"
        Ws.Cells(LineZ, 3) = "=SUM(C3:C" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 4) = "=SUM(D3:D" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 5) = "=SUM(E3:E" & LineZ - 1 & ")"
            'Ws.Cells(LineZ, 6) = "=SUM(F3:F" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 7) = "=SUM(G3:G" & LineZ - 1 & ")"
            'Ws.Cells(LineZ, 8) = "=SUM(H3:H" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 9) = "=SUM(I3:I" & LineZ - 1 & ")"
            'Ws.Cells(LineZ, 10) = "=SUM(J3:J" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 11) = "=SUM(K3:K" & LineZ - 1 & ")"
            'Ws.Cells(LineZ, 12) = "=SUM(L3:L" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 13) = "=SUM(M3:M" & LineZ - 1 & ")"
            'Ws.Cells(LineZ, 14) = "=SUM(N3:N" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 15) = "=SUM(O3:O" & LineZ - 1 & ")"
            'Ws.Cells(LineZ, 16) = "=SUM(P3:P" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 17) = "=SUM(Q3:Q" & LineZ - 1 & ")"
            'Ws.Cells(LineZ, 18) = "=SUM(R3:R" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 19) = "=SUM(S3:S" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 21) = "=SUM(U3:U" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 23) = "=SUM(W3:W" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 24) = "=SUM(X3:X" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 25) = "=SUM(Y3:Y" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 26) = "=SUM(Z3:Z" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 27) = "=SUM(AA3:AA" & LineZ - 1 & ")"
        Ws.Cells(LineZ, 28) = "=SUM(AB3:AB" & LineZ - 1 & ")"
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'If HaveReport > 0 Then
        SaveExcel()
        'End If
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Custom_Daily_Scrap_Information"
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
        If mConnection.State = ConnectionState.Open Then
            Try
                mConnection.Close()
                Module1.KillExcelProcess(OldExcel)
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 15
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.WrapText = True
        oRng = Ws.Range("A1", "Q1")
        oRng.Merge()
        oRng = Ws.Range("A1", "Q2")
        oRng.EntireRow.RowHeight = 42
        oRng = Ws.Range("C2", "M2")
        oRng.EntireColumn.ColumnWidth = 17.25
        oRng = Ws.Range("A2", "B2")
        oRng.EntireColumn.ColumnWidth = 23.28
        Ws.Cells(1, 1) = "报废日报表 Daliy Scrap Report"
        oRng = Ws.Range("A2", "A3")
        oRng.Merge()
        Ws.Cells(2, 1) = "型号Model"
        oRng = Ws.Range("B2", "B3")
        oRng.Merge()
        Ws.Cells(2, 2) = "缺陷原因" & Chr(10) & "Defect Code"
        oRng = Ws.Range("C2", "C3")
        oRng.Merge()
        Ws.Cells(2, 3) = "Prepreg 裁纱"
        oRng = Ws.Range("D2", "D3")
        oRng.Merge()
        Ws.Cells(2, 4) = "LayUp 预型"
        oRng = Ws.Range("E2", "E3")
        oRng.Merge()
        Ws.Cells(2, 5) = "Mold 成型"
        oRng = Ws.Range("F2", "F3")
        oRng.Merge()
        Ws.Cells(2, 6) = "Mold Scrap ERP P/N 成型报废ERP料号"
        oRng = Ws.Range("G2", "G3")
        oRng.Merge()
        Ws.Cells(2, 7) = "CNC"
        oRng = Ws.Range("H2", "H3")
        oRng.Merge()
        Ws.Cells(2, 8) = "CNC Scrap ERP P/N" & Chr(10) & "CNC报废ERP料号"
        oRng = Ws.Range("I2", "I3")
        oRng.Merge()
        Ws.Cells(2, 9) = "Sanding 补土"
        oRng = Ws.Range("J2", "J3")
        oRng.Merge()
        Ws.Cells(2, 10) = "Sanding Scrap ERP P/N 补土报废ERP料号"

        oRng = Ws.Range("K2", "P2")
        oRng.Merge()
        Ws.Cells(2, 11) = "Gluing 胶合"
        Ws.Cells(3, 11) = "Gluing 胶合 1"
        Ws.Cells(3, 12) = "Gluing 1 Scrap ERP P/N 胶合1报废ERP料号"
        Ws.Cells(3, 13) = "Gluing 胶合 2"
        Ws.Cells(3, 14) = "Gluing 2 Scrap ERP P/N 胶合1报废ERP料号"
        Ws.Cells(3, 15) = "Gluing 胶合 3"
        Ws.Cells(3, 16) = "Gluing 3 Scrap ERP P/N 胶合1报废ERP料号"
        oRng = Ws.Range("Q2", "Q3")
        oRng.Merge()
        Ws.Cells(2, 17) = "Painting 涂装"
        oRng = Ws.Range("R2", "R3")
        oRng.Merge()
        Ws.Cells(2, 18) = "Painting Scrap ERP P/N 涂装报废ERP料号"
        oRng = Ws.Range("S2", "S3")
        oRng.Merge()
        Ws.Cells(2, 19) = "Polishing 抛光"
        oRng = Ws.Range("T2", "T3")
        oRng.Merge()
        Ws.Cells(2, 20) = "Polishing Scrap ERP P/N 抛光报废ERP料号"
        oRng = Ws.Range("U2", "U3")
        oRng.Merge()
        Ws.Cells(2, 21) = "Packing 包装"
        oRng = Ws.Range("V2", "V3")
        oRng.Merge()
        Ws.Cells(2, 22) = "Packing Scrap ERP P/N 包装报废ERP料号"
        oRng = Ws.Range("W2", "W3")
        oRng.Merge()
        Ws.Cells(2, 23) = "'9999"
        oRng = Ws.Range("X2", "X3")
        oRng.Merge()
        Ws.Cells(2, 24) = "Others 其他"
        oRng = Ws.Range("Y2", "Y3")
        oRng.Merge()
        Ws.Cells(2, 25) = "不良总数" & Chr(10) & "Totally Defect"
        oRng = Ws.Range("Z2", "Z3")
        oRng.Merge()
        Ws.Cells(2, 26) = "0330检验数"
        oRng = Ws.Range("AA2", "AA3")
        oRng.Merge()
        Ws.Cells(2, 27) = "0331检验数"
        oRng = Ws.Range("AB2", "AB3")
        oRng.Merge()
        Ws.Cells(2, 28) = "配件报废数量"
        oRng = Ws.Range("AC2", "AC3")
        oRng.Merge()
        Ws.Cells(2, 29) = "配件 报废ERP料号"
        oRng = Ws.Range("AD2", "AD3")
        oRng.Merge()
        Ws.Cells(2, 30) = "配件报废工站"
        'oRng = Ws.Range("AB2", "AB3")
        'oRng.Merge()
        'Ws.Cells(2, 28) = "不良率" & Chr(10) & "Defect Rate"
        LineZ = 4
    End Sub
    Private Function Get0330Data(ByVal l_model As String, ByVal l_station As String)
        Dim mSqlS2 As New SqlClient.SqlCommand
        mSqlS2.Connection = mConnection
        mSqlS2.CommandType = CommandType.Text
        mSqlS2.CommandText = "select sum(t1) as t1 from ( "
        mSqlS2.CommandText += "select count(sn) as t1 from tracking left join lot on tracking.lot = lot.lot where timeout between '"
        mSqlS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSqlS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '" & l_station & "' and lot.model = '"
        mSqlS2.CommandText += l_model & "' union all select count(sn) as t1 from tracking_dup left join lot on tracking_dup.lot = lot.lot where timeout between '"
        mSqlS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSqlS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station = '" & l_station & "' and lot.model = '"
        mSqlS2.CommandText += l_model & "') as CA "
        Dim RV As Integer = mSqlS2.ExecuteScalar()
        If IsDBNull(RV) Then
            RV = 0
        End If
        Return RV
    End Function
End Class