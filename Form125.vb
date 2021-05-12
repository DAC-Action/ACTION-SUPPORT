Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Public Class Form125
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
    Dim pYear As Int16 = 0
    Dim tMonth As Int16 = 0
    Dim pMonth As Int16 = 0
    Dim lYear As Int16 = 0
    Dim tCurrency As String = String.Empty
    Dim ExchangeRate As Decimal = 0
    Dim ExchangeRate1 As Decimal = 0
    Dim gDatabase As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form125_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        If Today.Month < 10 Then
            TextBox1.Text = Today.Year & "0" & Today.Month
        Else
            TextBox1.Text = Today.Year & Today.Month
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        'Dim xPath As String = "C:\temp\IS.xlsx"
        ' If Not My.Computer.FileSystem.FileExists(xPath) Then
        'MsgBox("NO SAMPLE FILE")
        'Return
        'End If

        If TextBox1.Text.Length < 6 Then
            MsgBox("ERROR")
            Return
        End If
        gDatabase = Me.ComboBox2.SelectedItem.ToString()
        If String.IsNullOrEmpty(gDatabase) Then
            MsgBox("Database Error")
            Return
        End If
        Select Case gDatabase
            Case "DAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
            Case "HAC"
                oConnection.ConnectionString = Module1.OpenOracleConnection("hkacttest")
            Case "BVI"
                oConnection.ConnectionString = Module1.OpenOracleConnection("action_bvi")
        End Select
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

        tYear = Strings.Left(Me.TextBox1.Text, 4)
        pYear = tYear - 1
        tMonth = Strings.Right(Me.TextBox1.Text, 2)
        pMonth = tMonth - 1
        If pMonth = 0 Then
            pMonth = 12
            lYear = tYear - 1
        Else
            lYear = tYear
        End If
        tCurrency = Me.ComboBox1.SelectedItem.ToString()
        If String.IsNullOrEmpty(tCurrency) Then
            MsgBox("Currency Error")
            Return
        End If
        ' 確認 ExchangeRate
        If tCurrency = "USD" And gDatabase = "DAC" Then
            Dim CS As String = String.Empty
            If tMonth < 10 Then
                CS = tYear & "0" & tMonth
            Else
                CS = tYear & tMonth
            End If
            oCommand.CommandText = "SELECT nvl(AZJ041,0) FROM AZJ_FILE WHERE AZJ01  = 'USD' AND AZJ02 = '" & CS & "'"
            ExchangeRate = oCommand.ExecuteScalar()
            If ExchangeRate = 0 Then
                ExchangeRate = 1
            End If
            ExchangeRate1 = 6.3
        Else
            ExchangeRate = 1
            ExchangeRate1 = 1
        End If

        BackgroundWorker1.RunWorkerAsync()
        'ExportToExcel()
        'SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        AdjustExcelFormat1()
        DoInputData("100101", "101206", 0)
        LineZ += 1
        DoInputData("112201", "112202", 0)
        LineZ += 1
        DoInputData("1123", "1123", 0)
        LineZ += 1
        DoInputData("'140301','1471','1405','1409','1412','1413','1410','500101','500102','500103'", 0)
        LineZ += 1
        DoInputData("1124", "1124", 0)
        LineZ += 1
        'DoInputData("122101", "122102", 0)
        DoInputData("122101", "122103", 0)
        LineZ += 5
        DoInputData("151101", "1512", 0)
        LineZ += 1
        DoInputData("1524", "1524", 0)
        LineZ += 2
        DoInputData("160101", "160105", 0)
        LineZ += 1
        DoInputData("160201", "160204", 1)
        LineZ += 1
        DoInputData("160401", "160401", 0)
        LineZ += 1
        'DoInputData("1603", "1603", 1)
        DoInputDataMinus("1603", "1606", 1)
        LineZ += 1
        DoInputData("160901", "160901", 1)
        LineZ += 4
        'DoInputData("'170101','180101','180102','180103','180104','180105','180106'", 0)
        'LineZ += 1
        DoInputData("'170101','180207'", 0)
        LineZ += 1
        DoInputData("'180102','180103','180104','180106','180202','180203','180204','180206'", 0)
        LineZ += 6
        DoInputData("2001", "2001", 1)
        LineZ += 1
        'DoInputData("220201", "220204", 1)
        DoInputData("220201", "220206", 1)
        LineZ += 1
        'DoInputData("2203", "2206", 1)
        DoInputData("2203", "2208", 1)
        LineZ += 1
        DoInputData("2211", "2211", 1)
        LineZ += 1
        DoInputData("22210101", "222108", 1)
        LineZ += 1
        DoInputData("224101", "224103", 1)
        LineZ += 5
        DoInputData("2501", "2901", 1)
        LineZ += 6
        DoInputData("400102", "400104", 1)
        LineZ += 1
        DoInputData("400201", "400207", 1)
        LineZ += 1
        DoInputData("4105", "4105", 1)
        LineZ += 1
        DoInputData("4103", "4103", 1)
        LineZ += 1
        DoInputData("410401", "410412", 1)
        LineZ += 1
    End Sub
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.Font.Bold = True

        Ws.Name = "BS " & tYear
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.ColumnWidth = 33.67
        oRng.EntireColumn.Font.Name = "Arial Black"
        oRng = Ws.Range("B1", "Y1")
        oRng.EntireColumn.ColumnWidth = 20
        oRng = Ws.Range("C1", "C1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("G1", "G1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("I1", "I1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("K1", "K1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("M1", "M1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("O1", "O1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("Q1", "Q1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("S1", "S1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("U1", "U1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("W1", "W1")
        oRng.EntireColumn.ColumnWidth = 10
        oRng = Ws.Range("Y1", "Y1")
        oRng.EntireColumn.ColumnWidth = 10



        oRng = Ws.Range("A2", "Y2")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        oRng.Font.Size = 14
        oRng = Ws.Range("A3", "Y3")
        oRng.Merge()
        Ws.Cells(2, 1) = "Banlance Sheet"
        Select Case gDatabase
            Case "DAC"
                Ws.Cells(3, 1) = "Company Name：Dongguan Action Composite LTD. Co"
            Case "HAC"
                Ws.Cells(3, 1) = "Company Name：ACTION COMPOSITE TECHNOLOGY LIMITED"
            Case "BVI"
                Ws.Cells(3, 1) = "Company Name：ACTION COMPOSITES INTERNATIONAL LIMITED"
        End Select
        Ws.Cells(4, 1) = "Currency:"
        Ws.Cells(4, 2) = tCurrency
        oRng = Ws.Range("A5", "A5")
        oRng.EntireRow.Font.Color = Color.Azure
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(5, 1) = "Balance Sheet: Assets"
        oRng = Ws.Range("B5", "Y5")
        oRng.Merge()
        Ws.Cells(5, 2) = "Y" & tYear
        oRng = Ws.Range("A6", "A6")
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(6, 1) = "Year"
        For i As Int16 = 1 To 12 Step 1
            If i < 10 Then
                Ws.Cells(6, i * 2) = tYear & "/0" & i
            Else
                Ws.Cells(6, i * 2) = tYear & "/" & i
            End If
            Ws.Cells(6, i * 2 + 1) = "%"
        Next
        Ws.Cells(7, 1) = "Assets"
        Ws.Cells(8, 1) = "Current Assets"
        Ws.Cells(9, 1) = "Cash and cash equivalents"
        Ws.Cells(10, 1) = "Account Receivables"
        Ws.Cells(11, 1) = "Prepaid Payment"
        Ws.Cells(12, 1) = "Inventories"
        Ws.Cells(13, 1) = "Refundable Deposits"
        Ws.Cells(14, 1) = "Other Receivables"
        Ws.Cells(15, 1) = "Other Current Assets"
        Ws.Cells(16, 1) = "Total Current Assets"
        Ws.Cells(9, 3) = "=B9/B$30"
        oRng = Ws.Range("C9", "C9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("C9", "C16"), Type:=xlFillDefault)
        Ws.Cells(9, 5) = "=D9/D$30"
        oRng = Ws.Range("E9", "E9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("E9", "E16"), Type:=xlFillDefault)
        Ws.Cells(9, 7) = "=F9/F$30"
        oRng = Ws.Range("G9", "G9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("G9", "G16"), Type:=xlFillDefault)
        Ws.Cells(9, 9) = "=H9/H$30"
        oRng = Ws.Range("I9", "I9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("I9", "I16"), Type:=xlFillDefault)
        Ws.Cells(9, 11) = "=J9/J$30"
        oRng = Ws.Range("K9", "K9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("K9", "K16"), Type:=xlFillDefault)
        Ws.Cells(9, 13) = "=L9/L$30"
        oRng = Ws.Range("M9", "M9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("M9", "M16"), Type:=xlFillDefault)
        Ws.Cells(9, 15) = "=N9/N$30"
        oRng = Ws.Range("O9", "O9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("O9", "O16"), Type:=xlFillDefault)
        Ws.Cells(9, 17) = "=P9/P$30"
        oRng = Ws.Range("Q9", "Q9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Q9", "Q16"), Type:=xlFillDefault)

        Ws.Cells(9, 19) = "=R9/R$30"
        oRng = Ws.Range("S9", "S9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("S9", "S16"), Type:=xlFillDefault)
        Ws.Cells(9, 21) = "=T9/T$30"
        oRng = Ws.Range("U9", "U9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("U9", "U16"), Type:=xlFillDefault)
        Ws.Cells(9, 23) = "=V9/V$30"
        oRng = Ws.Range("W9", "W9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("W9", "W16"), Type:=xlFillDefault)
        Ws.Cells(9, 25) = "=X9/X$30"
        oRng = Ws.Range("Y9", "Y9")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Y9", "Y16"), Type:=xlFillDefault)

        Ws.Cells(16, 2) = "=SUM(B9:B15)"
        Ws.Cells(16, 4) = "=SUM(D9:D15)"
        Ws.Cells(16, 6) = "=SUM(F9:F15)"
        Ws.Cells(16, 8) = "=SUM(H9:H15)"
        Ws.Cells(16, 10) = "=SUM(J9:J15)"
        Ws.Cells(16, 12) = "=SUM(L9:L15)"
        Ws.Cells(16, 14) = "=SUM(N9:N15)"
        Ws.Cells(16, 16) = "=SUM(P9:P15)"
        Ws.Cells(16, 18) = "=SUM(R9:R15)"
        Ws.Cells(16, 20) = "=SUM(T9:T15)"
        Ws.Cells(16, 22) = "=SUM(V9:V15)"
        Ws.Cells(16, 24) = "=SUM(X9:X15)"

        oRng = Ws.Range("B9", "B31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("D9", "D31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("F9", "F31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("H9", "H31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("J9", "J31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("L9", "L31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("N9", "N31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("P9", "P31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("R9", "R31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("T9", "T31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("V9", "V31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("X9", "X31")
        oRng.NumberFormat = "#,##0;[Red]#,##0"

        oRng = Ws.Range("A5", "Y16")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 資產第二塊
        Ws.Cells(19, 1) = "Long-Term Investments"
        Ws.Cells(20, 1) = "Lease Assets_Rental"
        Ws.Cells(21, 1) = "Fixed Assets-Net value"
        Ws.Cells(21, 2) = "=B22-B23-B25"
        oRng = Ws.Range("B21", "B21")
        oRng.AutoFill(Destination:=Ws.Range("B21", "Y21"), Type:=xlFillDefault)
        Ws.Cells(22, 1) = "Fixed Assets"
        Ws.Cells(23, 1) = "Accumulated Depreciation"
        Ws.Cells(24, 1) = "Construction in Progress"
        Ws.Cells(25, 1) = "Fixed assets depreciation reserves"
        Ws.Cells(26, 1) = "WIP Accumulated Depreciation"
        Ws.Cells(27, 1) = "Total Fixed Assets"
        Ws.Cells(27, 2) = "=B21+B24-B26"
        oRng = Ws.Range("B27", "B27")
        oRng.AutoFill(Destination:=Ws.Range("B27", "Y27"), Type:=xlFillDefault)

        Ws.Cells(19, 3) = "=B19/B$32"
        oRng = Ws.Range("C19", "C19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("C19", "C27"), Type:=xlFillDefault)
        Ws.Cells(19, 5) = "=D19/D$32"
        oRng = Ws.Range("E19", "E19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("E19", "E27"), Type:=xlFillDefault)
        Ws.Cells(19, 7) = "=F19/F$32"
        oRng = Ws.Range("G19", "G19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("G19", "G27"), Type:=xlFillDefault)
        Ws.Cells(19, 9) = "=H19/H$32"
        oRng = Ws.Range("I19", "I19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("I19", "I27"), Type:=xlFillDefault)
        Ws.Cells(19, 11) = "=J19/J$32"
        oRng = Ws.Range("K19", "K19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("K19", "K27"), Type:=xlFillDefault)
        Ws.Cells(19, 13) = "=L19/L$32"
        oRng = Ws.Range("M19", "M19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("M19", "M27"), Type:=xlFillDefault)
        Ws.Cells(19, 15) = "=N19/N$32"
        oRng = Ws.Range("O19", "O19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("O19", "O27"), Type:=xlFillDefault)
        Ws.Cells(19, 17) = "=P19/P$32"
        oRng = Ws.Range("Q19", "Q19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Q19", "Q27"), Type:=xlFillDefault)

        Ws.Cells(19, 19) = "=R19/R$32"
        oRng = Ws.Range("S19", "S19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("S19", "S27"), Type:=xlFillDefault)
        Ws.Cells(19, 21) = "=T19/T$32"
        oRng = Ws.Range("U19", "U19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("U19", "U27"), Type:=xlFillDefault)
        Ws.Cells(19, 23) = "=V19/V$32"
        oRng = Ws.Range("W19", "W19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("W19", "W27"), Type:=xlFillDefault)
        Ws.Cells(19, 25) = "=X19/X$32"
        oRng = Ws.Range("Y19", "Y19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Y19", "Y27"), Type:=xlFillDefault)

        oRng = Ws.Range("A19", "Y27")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 資產第3塊

        Ws.Cells(29, 1) = "Intangible Assets & Other Assets"
        Ws.Cells(29, 2) = "=B30+B31"
        oRng = Ws.Range("B29", "B29")
        oRng.AutoFill(Destination:=Ws.Range("B29", "Y29"), Type:=xlFillDefault)
        Ws.Cells(30, 1) = "Intangible Assets"
        Ws.Cells(31, 1) = "Deferred Expenses"
        'Ws.Cells(29, 2) = "=B27-B28"
        Ws.Cells(32, 1) = "Total Assets"
        'oRng = Ws.Range("B29", "B29")
        'oRng.AutoFill(Destination:=Ws.Range("B29", "Y29"), Type:=xlFillDefault)
        Ws.Cells(32, 2) = "=B16+B19+B27+B29+B20"
        oRng = Ws.Range("B32", "B32")
        oRng.AutoFill(Destination:=Ws.Range("B32", "Y32"), Type:=xlFillDefault)

        Ws.Cells(29, 3) = "=B29/B$32"
        oRng = Ws.Range("C29", "C29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("C29", "C32"), Type:=xlFillDefault)
        Ws.Cells(29, 5) = "=D29/D$32"
        oRng = Ws.Range("E29", "E29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("E29", "E32"), Type:=xlFillDefault)
        Ws.Cells(29, 7) = "=F29/F$32"
        oRng = Ws.Range("G29", "G29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("G29", "G32"), Type:=xlFillDefault)
        Ws.Cells(29, 9) = "=H29/H$32"
        oRng = Ws.Range("I29", "I29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("I29", "I32"), Type:=xlFillDefault)
        Ws.Cells(29, 11) = "=J29/J$32"
        oRng = Ws.Range("K29", "K29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("K29", "K32"), Type:=xlFillDefault)
        Ws.Cells(29, 13) = "=L29/L$32"
        oRng = Ws.Range("M29", "M29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("M29", "M32"), Type:=xlFillDefault)
        Ws.Cells(29, 15) = "=N29/N$32"
        oRng = Ws.Range("O29", "O29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("O29", "O32"), Type:=xlFillDefault)
        Ws.Cells(29, 17) = "=P29/P$32"
        oRng = Ws.Range("Q29", "Q29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Q29", "Q32"), Type:=xlFillDefault)

        Ws.Cells(29, 19) = "=R29/R$32"
        oRng = Ws.Range("S29", "S29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("S29", "S32"), Type:=xlFillDefault)
        Ws.Cells(29, 21) = "=T29/T$32"
        oRng = Ws.Range("U29", "U29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("U29", "U32"), Type:=xlFillDefault)
        Ws.Cells(29, 23) = "=V29/V$32"
        oRng = Ws.Range("W29", "W29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("W29", "W32"), Type:=xlFillDefault)
        Ws.Cells(29, 25) = "=X29/X$32"
        oRng = Ws.Range("Y29", "Y29")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Y29", "Y32"), Type:=xlFillDefault)

        oRng = Ws.Range("A29", "Y32")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 負債和股東權益

        oRng = Ws.Range("A33", "A33")
        oRng.EntireRow.Font.Color = Color.Azure
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(5, 1) = "Balance Sheet: Assets"
        oRng = Ws.Range("B33", "Y33")
        oRng.Merge()
        Ws.Cells(33, 2) = "Y" & tYear
        oRng = Ws.Range("A34", "A34")
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(34, 1) = "Year"
        For i As Int16 = 1 To 12 Step 1
            If i < 10 Then
                Ws.Cells(34, i * 2) = tYear & "/0" & i
            Else
                Ws.Cells(34, i * 2) = tYear & "/" & i
            End If
            Ws.Cells(34, i * 2 + 1) = "%"
        Next
        Ws.Cells(35, 1) = "Total Liability & Equity"
        Ws.Cells(36, 1) = "Current Liability"
        Ws.Cells(37, 1) = "Short-Term loans"
        Ws.Cells(38, 1) = "Account Payable"
        Ws.Cells(39, 1) = "Advances from customers"
        Ws.Cells(40, 1) = "Accrued Payroll Exp"
        Ws.Cells(41, 1) = "Tax Payable"
        Ws.Cells(42, 1) = "Other Payables"
        Ws.Cells(43, 1) = "Accrued Expenses"
        Ws.Cells(44, 1) = "Total Current Liability"

        Ws.Cells(46, 1) = "Long-Term Liability"
        Ws.Cells(47, 1) = "Other Liability"
        Ws.Cells(49, 1) = "Total Liability"
        Ws.Cells(51, 1) = "Shareholders' Equity"
        Ws.Cells(52, 1) = "Capital"
        Ws.Cells(53, 1) = "Capital by Investees"
        Ws.Cells(54, 1) = "Capital Premiun"
        Ws.Cells(55, 1) = "Exchange Reserve"
        Ws.Cells(56, 1) = "Profit and Loss-Current year"
        Ws.Cells(57, 1) = "Profit or loss for prior year"
        Ws.Cells(59, 1) = "Retained Earnings"
        Ws.Cells(60, 1) = "Total Shareholders' Equity"
        Ws.Cells(62, 1) = "Total Liability & Equity"

        Ws.Cells(44, 2) = "=SUM(B36:B42)"
        oRng = Ws.Range("B44", "B44")
        oRng.AutoFill(Destination:=Ws.Range("B44", "X44"), Type:=xlFillDefault)
        Ws.Cells(49, 2) = "=SUM(B44:B47)"
        oRng = Ws.Range("B49", "B49")
        oRng.AutoFill(Destination:=Ws.Range("B49", "X49"), Type:=xlFillDefault)
        'Ws.Cells(55, 2) = "=B57-B54"
        'oRng = Ws.Range("B55", "B55")
        'oRng.AutoFill(Destination:=Ws.Range("B55", "X55"), Type:=xlFillDefault)
        Ws.Cells(59, 2) = "=B56+B57"
        oRng = Ws.Range("B59", "B59")
        oRng.AutoFill(Destination:=Ws.Range("B59", "X59"), Type:=xlFillDefault)
        Ws.Cells(60, 2) = "=SUM(B52:B57)"
        oRng = Ws.Range("B60", "B60")
        oRng.AutoFill(Destination:=Ws.Range("B60", "X60"), Type:=xlFillDefault)
        Ws.Cells(62, 2) = "=B49+B60"
        oRng = Ws.Range("B62", "B62")
        oRng.AutoFill(Destination:=Ws.Range("B62", "X62"), Type:=xlFillDefault)

        Ws.Cells(37, 3) = "=B37/B$62"
        oRng = Ws.Range("C37", "C37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("C37", "C62"), Type:=xlFillDefault)
        Ws.Cells(37, 5) = "=D37/D$62"
        oRng = Ws.Range("E37", "E37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("E37", "E62"), Type:=xlFillDefault)
        Ws.Cells(37, 7) = "=F37/F$62"
        oRng = Ws.Range("G37", "G37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("G37", "G62"), Type:=xlFillDefault)
        Ws.Cells(37, 9) = "=H37/H$62"
        oRng = Ws.Range("I37", "I37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("I37", "I62"), Type:=xlFillDefault)
        Ws.Cells(37, 11) = "=J37/J$62"
        oRng = Ws.Range("K37", "K37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("K37", "K62"), Type:=xlFillDefault)
        Ws.Cells(37, 13) = "=L37/L$62"
        oRng = Ws.Range("M37", "M37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("M37", "M62"), Type:=xlFillDefault)
        Ws.Cells(37, 15) = "=N37/N$62"
        oRng = Ws.Range("O37", "O37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("O37", "O62"), Type:=xlFillDefault)
        Ws.Cells(37, 17) = "=P37/P$62"
        oRng = Ws.Range("Q37", "Q37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Q37", "Q62"), Type:=xlFillDefault)

        Ws.Cells(37, 19) = "=R37/R$62"
        oRng = Ws.Range("S37", "S37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("S37", "S62"), Type:=xlFillDefault)
        Ws.Cells(37, 21) = "=T37/T$62"
        oRng = Ws.Range("U37", "U37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("U37", "U62"), Type:=xlFillDefault)
        Ws.Cells(37, 23) = "=V37/V$62"
        oRng = Ws.Range("W37", "W37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("W37", "W62"), Type:=xlFillDefault)
        Ws.Cells(37, 25) = "=X37/X$62"
        oRng = Ws.Range("Y37", "Y37")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Y37", "Y62"), Type:=xlFillDefault)


        ' 清除
        oRng = Ws.Range("C45", "Y45")
        oRng.ClearContents()
        oRng = Ws.Range("C48", "Y48")
        oRng.ClearContents()
        oRng = Ws.Range("C50", "Y51")
        oRng.ClearContents()
        oRng = Ws.Range("C58", "Y58")
        oRng.ClearContents()
        oRng = Ws.Range("C61", "Y61")
        oRng.ClearContents()

        oRng = Ws.Range("B37", "B62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("D37", "D62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("F37", "F62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("H37", "H62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("J37", "J62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("L37", "L62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("N37", "N62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("P37", "P62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("R37", "R62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("T37", "T62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("V37", "V62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("X37", "X62")
        oRng.NumberFormat = "#,##0;[Red]#,##0"

        oRng = Ws.Range("A33", "Y44")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A46", "Y47")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A49", "Y49")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A51", "Y57")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A59", "Y60")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A62", "Y62")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        LineZ = 9
    End Sub

    Private Sub DoInputData(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As Int16)

        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
        Next
        oCommand.CommandText += "nvl(sum(s1),0) as s1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then (aah05 - aah04) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04),3) else 0 end ) as s1 from aah_file,aag_file where aah00 =aag00 and aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' )"


        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    ' 設定期初
                    Dim Scount As Decimal = 0
                    Scount += oReader.Item("s1")
                    For j As Int16 = 1 To i Step 1
                        Scount += oReader.Item(j - 1)
                    Next
                    ' Scount = 本幣額
                    ' 計算匯率
                    If tCurrency = "USD" And gDatabase = "DAC" Then
                        oCommand2.CommandText = "select nvl(azj041,1) from azj_file where azj01 = 'USD' and azj02 = '" & tYear
                        If i < 10 Then
                            oCommand2.CommandText += "0" & i
                        Else
                            oCommand2.CommandText += i.ToString()
                        End If
                        oCommand2.CommandText += "'"
                        ExchangeRate = oCommand2.ExecuteScalar
                    Else
                        ExchangeRate = 1
                    End If
                    oCommand2.CommandText = ""
                    If ACC3 = 0 Then
                        Ws.Cells(LineZ, 2 * i) = Decimal.Round(Scount * Decimal.MinusOne / ExchangeRate, 2)
                    Else
                        Ws.Cells(LineZ, 2 * i) = Decimal.Round(Scount / ExchangeRate, 2)
                    End If

                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub DoInputData(ByVal ACC1 As String, ByVal ACC3 As Int16)
        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
        Next
        oCommand.CommandText += "nvl(sum(s1),0) as s1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then (aah05 - aah04) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04),3) else 0 end ) as s1 from aah_file,aag_file where aah00 =aag00  and aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 in (" & ACC1 & " ) )"

        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    ' 設定期初
                    Dim Scount As Decimal = 0
                    Scount += oReader.Item("s1")
                    For j As Int16 = 1 To i Step 1
                        Scount += oReader.Item(j - 1)
                    Next
                    ' Scount = 本幣額
                    ' 計算匯率
                    If tCurrency = "USD" And gDatabase = "DAC" Then
                        oCommand2.CommandText = "select nvl(azj041,1) from azj_file where azj01 = 'USD' and azj02 = '" & tYear
                        If i < 10 Then
                            oCommand2.CommandText += "0" & i
                        Else
                            oCommand2.CommandText += i.ToString()
                        End If
                        oCommand2.CommandText += "'"
                        ExchangeRate = oCommand2.ExecuteScalar
                    Else
                        ExchangeRate = 1
                    End If
                    oCommand2.CommandText = ""
                    If ACC3 = 0 Then
                        Ws.Cells(LineZ, 2 * i) = Decimal.Round(Scount * Decimal.MinusOne / ExchangeRate, 2)
                    Else
                        Ws.Cells(LineZ, 2 * i) = Decimal.Round(Scount / ExchangeRate, 2)
                    End If

                Next
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Balance_Sheet_" & gDatabase
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
    Private Sub DoInputDataMinus(ByVal ACC1 As String, ByVal ACC2 As String, ByVal ACC3 As String)
        oCommand.CommandText = "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "nvl(sum(t" & i & "),0) as t" & i & ","
        Next
        oCommand.CommandText += "nvl(sum(s1),0) as s1 from ( select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then (aah05 - aah04) else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04),3) else 0 end ) as s1 from aah_file,aag_file where aah00 =aag00  and aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 = '" & ACC1 & "' "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then (aah05 - aah04)  else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04) ,3) else 0 end ) as s1 from aah_file,aag_file where aah00 =aag00 and aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 = '" & ACC2 & "') "


        oReader = oCommand.ExecuteReader()
        If oReader.HasRows Then
            While oReader.Read()
                For i As Int16 = 1 To tMonth Step 1
                    ' 設定期初
                    Dim Scount As Decimal = 0
                    Scount += oReader.Item("s1")
                    For j As Int16 = 1 To i Step 1
                        Scount += oReader.Item(j - 1)
                    Next
                    ' Scount = 本幣額
                    ' 計算匯率
                    If tCurrency = "USD" And gDatabase = "DAC" Then
                        oCommand2.CommandText = "select nvl(azj041,1) from azj_file where azj01 = 'USD' and azj02 = '" & tYear
                        If i < 10 Then
                            oCommand2.CommandText += "0" & i
                        Else
                            oCommand2.CommandText += i.ToString()
                        End If
                        oCommand2.CommandText += "'"
                        ExchangeRate = oCommand2.ExecuteScalar
                    Else
                        ExchangeRate = 1
                    End If
                    oCommand2.CommandText = ""
                    If ACC3 = 0 Then
                        Ws.Cells(LineZ, 2 * i) = Decimal.Round(Scount * Decimal.MinusOne / ExchangeRate, 2)
                    Else
                        Ws.Cells(LineZ, 2 * i) = Decimal.Round(Scount / ExchangeRate, 2)
                    End If

                Next
            End While
        End If
        oReader.Close()
    End Sub
End Class