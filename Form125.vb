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
        DoInputData("122101", "122102", 0)
        LineZ += 5
        DoInputData("151101", "1512", 0)
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
        DoInputData("220201", "220204", 1)
        LineZ += 1
        DoInputData("2203", "2206", 1)
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
        Ws.Cells(20, 1) = "Fixed Assets-Net value"
        Ws.Cells(20, 2) = "=B21-B22-B24"
        oRng = Ws.Range("B20", "B20")
        oRng.AutoFill(Destination:=Ws.Range("B20", "Y20"), Type:=xlFillDefault)
        Ws.Cells(21, 1) = "Fixed Assets"
        Ws.Cells(22, 1) = "Accumulated Depreciation"
        Ws.Cells(23, 1) = "Construction in Progress"
        Ws.Cells(24, 1) = "Fixed assets depreciation reserves"
        Ws.Cells(25, 1) = "WIP Accumulated Depreciation"
        Ws.Cells(26, 1) = "Total Fixed Assets"
        Ws.Cells(26, 2) = "=B20+B23-B25"
        oRng = Ws.Range("B26", "B26")
        oRng.AutoFill(Destination:=Ws.Range("B26", "Y26"), Type:=xlFillDefault)

        Ws.Cells(19, 3) = "=B19/B$31"
        oRng = Ws.Range("C19", "C19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("C19", "C26"), Type:=xlFillDefault)
        Ws.Cells(19, 5) = "=D19/D$31"
        oRng = Ws.Range("E19", "E19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("E19", "E26"), Type:=xlFillDefault)
        Ws.Cells(19, 7) = "=F19/F$31"
        oRng = Ws.Range("G19", "G19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("G19", "G26"), Type:=xlFillDefault)
        Ws.Cells(19, 9) = "=H19/H$31"
        oRng = Ws.Range("I19", "I19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("I19", "I26"), Type:=xlFillDefault)
        Ws.Cells(19, 11) = "=J19/J$31"
        oRng = Ws.Range("K19", "K19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("K19", "K26"), Type:=xlFillDefault)
        Ws.Cells(19, 13) = "=L19/L$31"
        oRng = Ws.Range("M19", "M19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("M19", "M26"), Type:=xlFillDefault)
        Ws.Cells(19, 15) = "=N19/N$31"
        oRng = Ws.Range("O19", "O19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("O19", "O26"), Type:=xlFillDefault)
        Ws.Cells(19, 17) = "=P19/P$31"
        oRng = Ws.Range("Q19", "Q19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Q19", "Q26"), Type:=xlFillDefault)

        Ws.Cells(19, 19) = "=R19/R$31"
        oRng = Ws.Range("S19", "S19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("S19", "S26"), Type:=xlFillDefault)
        Ws.Cells(19, 21) = "=T19/T$31"
        oRng = Ws.Range("U19", "U19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("U19", "U26"), Type:=xlFillDefault)
        Ws.Cells(19, 23) = "=V19/V$31"
        oRng = Ws.Range("W19", "W19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("W19", "W26"), Type:=xlFillDefault)
        Ws.Cells(19, 25) = "=X19/X$31"
        oRng = Ws.Range("Y19", "Y19")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Y19", "Y26"), Type:=xlFillDefault)

        oRng = Ws.Range("A19", "Y26")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 資產第3塊

        Ws.Cells(28, 1) = "Intangible Assets & Other Assets"
        Ws.Cells(28, 2) = "=B29+B30"
        oRng = Ws.Range("B28", "B28")
        oRng.AutoFill(Destination:=Ws.Range("B28", "Y28"), Type:=xlFillDefault)
        Ws.Cells(29, 1) = "Intangible Assets"
        Ws.Cells(30, 1) = "Deferred Expenses"
        'Ws.Cells(29, 2) = "=B27-B28"
        Ws.Cells(31, 1) = "Total Assets"
        'oRng = Ws.Range("B29", "B29")
        'oRng.AutoFill(Destination:=Ws.Range("B29", "Y29"), Type:=xlFillDefault)
        Ws.Cells(31, 2) = "=B16+B19+B26+B28"
        oRng = Ws.Range("B31", "B31")
        oRng.AutoFill(Destination:=Ws.Range("B31", "Y31"), Type:=xlFillDefault)

        Ws.Cells(28, 3) = "=B28/B$31"
        oRng = Ws.Range("C28", "C28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("C28", "C31"), Type:=xlFillDefault)
        Ws.Cells(28, 5) = "=D28/D$31"
        oRng = Ws.Range("E28", "E28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("E28", "E31"), Type:=xlFillDefault)
        Ws.Cells(28, 7) = "=F28/F$31"
        oRng = Ws.Range("G28", "G28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("G28", "G31"), Type:=xlFillDefault)
        Ws.Cells(28, 9) = "=H28/H$31"
        oRng = Ws.Range("I28", "I28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("I28", "I31"), Type:=xlFillDefault)
        Ws.Cells(28, 11) = "=J28/J$31"
        oRng = Ws.Range("K28", "K28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("K28", "K31"), Type:=xlFillDefault)
        Ws.Cells(28, 13) = "=L28/L$31"
        oRng = Ws.Range("M28", "M28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("M28", "M31"), Type:=xlFillDefault)
        Ws.Cells(28, 15) = "=N28/N$31"
        oRng = Ws.Range("O28", "O28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("O28", "O31"), Type:=xlFillDefault)
        Ws.Cells(28, 17) = "=P28/P$31"
        oRng = Ws.Range("Q28", "Q28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Q28", "Q31"), Type:=xlFillDefault)

        Ws.Cells(28, 19) = "=R28/R$31"
        oRng = Ws.Range("S28", "S28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("S28", "S31"), Type:=xlFillDefault)
        Ws.Cells(28, 21) = "=T28/T$31"
        oRng = Ws.Range("U28", "U28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("U28", "U31"), Type:=xlFillDefault)
        Ws.Cells(28, 23) = "=V28/V$31"
        oRng = Ws.Range("W28", "W28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("W28", "W31"), Type:=xlFillDefault)
        Ws.Cells(28, 25) = "=X28/X$31"
        oRng = Ws.Range("Y28", "Y28")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Y28", "Y31"), Type:=xlFillDefault)

        oRng = Ws.Range("A28", "Y31")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        ' 負債和股東權益

        oRng = Ws.Range("A32", "A32")
        oRng.EntireRow.Font.Color = Color.Azure
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(5, 1) = "Balance Sheet: Assets"
        oRng = Ws.Range("B32", "Y32")
        oRng.Merge()
        Ws.Cells(32, 2) = "Y" & tYear
        oRng = Ws.Range("A33", "A33")
        oRng.EntireRow.HorizontalAlignment = xlCenter
        Ws.Cells(33, 1) = "Year"
        For i As Int16 = 1 To 12 Step 1
            If i < 10 Then
                Ws.Cells(33, i * 2) = tYear & "/0" & i
            Else
                Ws.Cells(33, i * 2) = tYear & "/" & i
            End If
            Ws.Cells(33, i * 2 + 1) = "%"
        Next
        Ws.Cells(34, 1) = "Total Liability & Equity"
        Ws.Cells(35, 1) = "Current Liability"
        Ws.Cells(36, 1) = "Short-Term loans"
        Ws.Cells(37, 1) = "Account Payable"
        Ws.Cells(38, 1) = "Advances from customers"
        Ws.Cells(39, 1) = "Accrued Payroll Exp"
        Ws.Cells(40, 1) = "Tax Payable"
        Ws.Cells(41, 1) = "Other Payables"
        Ws.Cells(42, 1) = "Accrued Expenses"
        Ws.Cells(43, 1) = "Total Current Liability"

        Ws.Cells(45, 1) = "Long-Term Liability"
        Ws.Cells(46, 1) = "Other Liability"
        Ws.Cells(48, 1) = "Total Liability"
        Ws.Cells(50, 1) = "Shareholders' Equity"
        Ws.Cells(51, 1) = "Capital"
        Ws.Cells(52, 1) = "Capital by Investees"
        Ws.Cells(53, 1) = "Capital Premiun"
        Ws.Cells(54, 1) = "Exchange Reserve"
        Ws.Cells(55, 1) = "Profit and Loss-Current year"
        Ws.Cells(56, 1) = "Profit or loss for prior year"
        Ws.Cells(58, 1) = "Retained Earnings"
        Ws.Cells(59, 1) = "Total Shareholders' Equity"
        Ws.Cells(61, 1) = "Total Liability & Equity"

        Ws.Cells(43, 2) = "=SUM(B36:B42)"
        oRng = Ws.Range("B43", "B43")
        oRng.AutoFill(Destination:=Ws.Range("B43", "X43"), Type:=xlFillDefault)
        Ws.Cells(48, 2) = "=SUM(B43:B46)"
        oRng = Ws.Range("B48", "B48")
        oRng.AutoFill(Destination:=Ws.Range("B48", "X48"), Type:=xlFillDefault)
        'Ws.Cells(55, 2) = "=B57-B54"
        'oRng = Ws.Range("B55", "B55")
        'oRng.AutoFill(Destination:=Ws.Range("B55", "X55"), Type:=xlFillDefault)
        Ws.Cells(58, 2) = "=B55+B56"
        oRng = Ws.Range("B58", "B58")
        oRng.AutoFill(Destination:=Ws.Range("B58", "X58"), Type:=xlFillDefault)
        Ws.Cells(59, 2) = "=SUM(B51:B56)"
        oRng = Ws.Range("B59", "B59")
        oRng.AutoFill(Destination:=Ws.Range("B59", "X59"), Type:=xlFillDefault)
        Ws.Cells(61, 2) = "=B48+B59"
        oRng = Ws.Range("B61", "B61")
        oRng.AutoFill(Destination:=Ws.Range("B61", "X61"), Type:=xlFillDefault)

        Ws.Cells(36, 3) = "=B36/B$61"
        oRng = Ws.Range("C36", "C36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("C36", "C61"), Type:=xlFillDefault)
        Ws.Cells(36, 5) = "=D36/D$61"
        oRng = Ws.Range("E36", "E36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("E36", "E61"), Type:=xlFillDefault)
        Ws.Cells(36, 7) = "=F36/F$61"
        oRng = Ws.Range("G36", "G36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("G36", "G61"), Type:=xlFillDefault)
        Ws.Cells(36, 9) = "=H36/H$61"
        oRng = Ws.Range("I36", "I36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("I36", "I61"), Type:=xlFillDefault)
        Ws.Cells(36, 11) = "=J36/J$61"
        oRng = Ws.Range("K36", "K36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("K36", "K61"), Type:=xlFillDefault)
        Ws.Cells(36, 13) = "=L36/L$61"
        oRng = Ws.Range("M36", "M36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("M36", "M61"), Type:=xlFillDefault)
        Ws.Cells(36, 15) = "=N36/N$61"
        oRng = Ws.Range("O36", "O36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("O36", "O61"), Type:=xlFillDefault)
        Ws.Cells(36, 17) = "=P36/P$61"
        oRng = Ws.Range("Q36", "Q36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Q36", "Q61"), Type:=xlFillDefault)

        Ws.Cells(36, 19) = "=R36/R$61"
        oRng = Ws.Range("S36", "S36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("S36", "S61"), Type:=xlFillDefault)
        Ws.Cells(36, 21) = "=T36/T$61"
        oRng = Ws.Range("U36", "U36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("U36", "U61"), Type:=xlFillDefault)
        Ws.Cells(36, 23) = "=V36/V$61"
        oRng = Ws.Range("W36", "W36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("W36", "W61"), Type:=xlFillDefault)
        Ws.Cells(36, 25) = "=X36/X$61"
        oRng = Ws.Range("Y36", "Y36")
        oRng.NumberFormat = "0.00%"
        oRng.AutoFill(Destination:=Ws.Range("Y36", "Y61"), Type:=xlFillDefault)


        ' 清除
        oRng = Ws.Range("C44", "Y44")
        oRng.ClearContents()
        oRng = Ws.Range("C47", "Y47")
        oRng.ClearContents()
        oRng = Ws.Range("C49", "Y50")
        oRng.ClearContents()
        oRng = Ws.Range("C57", "Y57")
        oRng.ClearContents()
        oRng = Ws.Range("C60", "Y60")
        oRng.ClearContents()

        oRng = Ws.Range("B36", "B61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("D36", "D61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("F36", "F61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("H36", "H61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("J36", "J61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("L36", "L61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("N36", "N61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("P36", "P61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("R36", "R61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("T36", "T61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("V36", "V61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"
        oRng = Ws.Range("X36", "X61")
        oRng.NumberFormat = "#,##0;[Red]#,##0"

        oRng = Ws.Range("A32", "Y43")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A45", "Y46")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A48", "Y48")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A50", "Y56")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A58", "Y59")
        oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        oRng.Borders(xlInsideVertical).LineStyle = xlContinuous

        oRng = Ws.Range("A61", "Y61")
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
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04),3) else 0 end ) as s1 from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 between '" & ACC1 & "' and '" & ACC2 & "' )"


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
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04),3) else 0 end ) as s1 from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 in (" & ACC1 & " ) )"

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
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04),3) else 0 end ) as s1 from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 = '" & ACC1 & "' "
        oCommand.CommandText += "union all "
        oCommand.CommandText += "select "
        For i As Int16 = 1 To tMonth Step 1
            oCommand.CommandText += "(case when aah03 = " & i & " then (aah05 - aah04)  else 0 end ) as t" & i & ","
        Next
        oCommand.CommandText += "(case when aah03 = 0 then round((aah05 - aah04) ,3) else 0 end ) as s1 from aah_file,aag_file where aah01 = aag01 and aag07 in ('2','3') and aah02 = " & tYear & " and aah03 <= " & tMonth & " and aah01 = '" & ACC2 & "') "


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