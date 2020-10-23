Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form197
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
    Dim l_tc_bmu01 As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form197_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        If String.IsNullOrEmpty(TextBox1.Text) Then
            MsgBox("请输入RFQ单号")
            Return
        End If
        l_tc_bmu01 = TextBox1.Text
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
        oCommand.CommandText = "Select count(*) from tc_bmu_file where tc_bmu01 = '" & l_tc_bmu01 & "'"
        Dim HasRows As Int16 = oCommand.ExecuteScalar()
        If HasRows = 0 Then
            MsgBox("无此RFQ单号")
            Return
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
        SaveFileDialog1.FileName = "RFQ"
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
        Ws.Name = "RFQ"
        Ws.Activate()
        AdjustExcelFormat()

        oCommand.CommandText = "Select * from tc_bmu_file where tc_bmu01 = '" & l_tc_bmu01 & "'"
        oReader = oCommand.ExecuteReader
        If oReader.HasRows() Then
            While oReader.Read()
                oRng = Ws.Range(Ws.Cells(LineZ, 1), Ws.Cells(LineZ + 5, 1))
                oRng.Merge()
                Ws.Cells(LineZ, 1) = oReader.Item("tc_bmu02")
                oRng = Ws.Range(Ws.Cells(LineZ, 2), Ws.Cells(LineZ + 5, 2))
                oRng.Merge()
                Ws.Cells(LineZ, 2) = oReader.Item("tc_bmuud01")
                oRng = Ws.Range(Ws.Cells(LineZ, 3), Ws.Cells(LineZ + 5, 3))
                oRng.Merge()
                Ws.Cells(LineZ, 3) = oReader.Item("tc_bmu03")
                oRng = Ws.Range(Ws.Cells(LineZ, 4), Ws.Cells(LineZ + 5, 4))
                oRng.Merge()
                Ws.Cells(LineZ, 4) = "方案" & oReader.Item("tc_bmu04")
                oRng = Ws.Range(Ws.Cells(LineZ, 5), Ws.Cells(LineZ + 5, 5))
                oRng.Merge()
                Ws.Cells(LineZ, 5) = oReader.Item("tc_bmu05")
                oRng = Ws.Range(Ws.Cells(LineZ, 6), Ws.Cells(LineZ + 5, 6))
                oRng.Merge()
                Ws.Cells(LineZ, 6) = oReader.Item("tc_bmu06")
                oRng = Ws.Range(Ws.Cells(LineZ, 7), Ws.Cells(LineZ + 5, 7))
                oRng.Merge()
                Ws.Cells(LineZ, 7) = oReader.Item("tc_bmu07")
                oRng = Ws.Range(Ws.Cells(LineZ, 8), Ws.Cells(LineZ + 2, 8))
                oRng.Merge()
                Ws.Cells(LineZ, 8) = 1
                oRng = Ws.Range(Ws.Cells(LineZ + 3, 8), Ws.Cells(LineZ + 5, 8))
                oRng.Merge()
                Ws.Cells(LineZ + 3, 8) = oReader.Item("tc_bmu09")

                oRng = Ws.Range(Ws.Cells(LineZ, 9), Ws.Cells(LineZ + 2, 9))
                oRng.Merge()
                Ws.Cells(LineZ, 9) = oReader.Item("tc_bmu10")
                oRng = Ws.Range(Ws.Cells(LineZ + 3, 9), Ws.Cells(LineZ + 5, 9))
                oRng.Merge()
                Ws.Cells(LineZ + 3, 9) = oReader.Item("tc_bmu14")
                oRng = Ws.Range(Ws.Cells(LineZ, 10), Ws.Cells(LineZ + 2, 10))
                oRng.Merge()
                Ws.Cells(LineZ, 10) = oReader.Item("tc_bmu11")
                oRng = Ws.Range(Ws.Cells(LineZ + 3, 10), Ws.Cells(LineZ + 5, 10))
                oRng.Merge()
                Ws.Cells(LineZ + 3, 10) = oReader.Item("tc_bmu15")
                oRng = Ws.Range(Ws.Cells(LineZ, 11), Ws.Cells(LineZ + 2, 11))
                oRng.Merge()
                Ws.Cells(LineZ, 11) = oReader.Item("tc_bmu12")
                oRng = Ws.Range(Ws.Cells(LineZ + 3, 11), Ws.Cells(LineZ + 5, 11))
                oRng.Merge()
                Ws.Cells(LineZ + 3, 11) = oReader.Item("tc_bmu16")
                oRng = Ws.Range(Ws.Cells(LineZ, 12), Ws.Cells(LineZ + 2, 12))
                oRng.Merge()
                Ws.Cells(LineZ, 12) = oReader.Item("tc_bmu13")
                oRng = Ws.Range(Ws.Cells(LineZ + 3, 12), Ws.Cells(LineZ + 5, 12))
                oRng.Merge()
                Ws.Cells(LineZ + 3, 12) = oReader.Item("tc_bmu17")

                Ws.Cells(LineZ, 13) = oReader.Item("tc_bmu18")
                Ws.Cells(LineZ + 1, 13) = oReader.Item("tc_bmu23")
                Ws.Cells(LineZ + 2, 13) = oReader.Item("tc_bmu28")
                Ws.Cells(LineZ + 3, 13) = oReader.Item("tc_bmu33")
                Ws.Cells(LineZ + 4, 13) = oReader.Item("tc_bmu38")
                Ws.Cells(LineZ + 5, 13) = oReader.Item("tc_bmu43")

                Ws.Cells(LineZ, 14) = oReader.Item("tc_bmu19")
                Ws.Cells(LineZ + 1, 14) = oReader.Item("tc_bmu24")
                Ws.Cells(LineZ + 2, 14) = oReader.Item("tc_bmu29")
                Ws.Cells(LineZ + 3, 14) = oReader.Item("tc_bmu34")
                Ws.Cells(LineZ + 4, 14) = oReader.Item("tc_bmu39")
                Ws.Cells(LineZ + 5, 14) = oReader.Item("tc_bmu44")

                Ws.Cells(LineZ, 15) = oReader.Item("tc_bmu20")
                Ws.Cells(LineZ + 1, 15) = oReader.Item("tc_bmu25")
                Ws.Cells(LineZ + 2, 15) = oReader.Item("tc_bmu30")
                Ws.Cells(LineZ + 3, 15) = oReader.Item("tc_bmu35")
                Ws.Cells(LineZ + 4, 15) = oReader.Item("tc_bmu40")
                Ws.Cells(LineZ + 5, 15) = oReader.Item("tc_bmu45")

                Ws.Cells(LineZ, 16) = oReader.Item("tc_bmu21")
                Ws.Cells(LineZ + 1, 16) = oReader.Item("tc_bmu26")
                Ws.Cells(LineZ + 2, 16) = oReader.Item("tc_bmu31")
                Ws.Cells(LineZ + 3, 16) = oReader.Item("tc_bmu36")
                Ws.Cells(LineZ + 4, 16) = oReader.Item("tc_bmu41")
                Ws.Cells(LineZ + 5, 16) = oReader.Item("tc_bmu46")

                If Not IsDBNull(oReader.Item("tc_bmu22")) Then
                    Select Case oReader.Item("tc_bmu22")
                        Case "A"
                            Ws.Cells(LineZ, 17) = "Day"
                        Case "B"
                            Ws.Cells(LineZ, 17) = "Week"
                        Case "C"
                            Ws.Cells(LineZ, 17) = "Month"
                    End Select
                End If
                If Not IsDBNull(oReader.Item("tc_bmu27")) Then
                    Select Case oReader.Item("tc_bmu27")
                        Case "A"
                            Ws.Cells(LineZ + 1, 17) = "Day"
                        Case "B"
                            Ws.Cells(LineZ + 1, 17) = "Week"
                        Case "C"
                            Ws.Cells(LineZ + 1, 17) = "Month"
                    End Select
                End If
                If Not IsDBNull(oReader.Item("tc_bmu32")) Then
                    Select Case oReader.Item("tc_bmu32")
                        Case "A"
                            Ws.Cells(LineZ + 2, 17) = "Day"
                        Case "B"
                            Ws.Cells(LineZ + 2, 17) = "Week"
                        Case "C"
                            Ws.Cells(LineZ + 2, 17) = "Month"
                    End Select
                End If
                If Not IsDBNull(oReader.Item("tc_bmu37")) Then
                    Select Case oReader.Item("tc_bmu37")
                        Case "A"
                            Ws.Cells(LineZ + 3, 17) = "Day"
                        Case "B"
                            Ws.Cells(LineZ + 3, 17) = "Week"
                        Case "C"
                            Ws.Cells(LineZ + 3, 17) = "Month"
                    End Select
                End If
                If Not IsDBNull(oReader.Item("tc_bmu42")) Then
                    Select Case oReader.Item("tc_bmu42")
                        Case "A"
                            Ws.Cells(LineZ + 4, 17) = "Day"
                        Case "B"
                            Ws.Cells(LineZ + 4, 17) = "Week"
                        Case "C"
                            Ws.Cells(LineZ + 4, 17) = "Month"
                    End Select
                End If
                If Not IsDBNull(oReader.Item("tc_bmu47")) Then
                    Select Case oReader.Item("tc_bmu47")
                        Case "A"
                            Ws.Cells(LineZ + 5, 17) = "Day"
                        Case "B"
                            Ws.Cells(LineZ + 5, 17) = "Week"
                        Case "C"
                            Ws.Cells(LineZ + 5, 17) = "Month"
                    End Select
                End If
                LineZ += 6
            End While
        End If
        oReader.Close()
    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 16
        oRng = Ws.Range("B1", "D1")
        oRng.Merge()
        oRng.HorizontalAlignment = xlCenter
        Ws.Cells(1, 1) = "单号："
        Ws.Cells(1, 2) = l_tc_bmu01
        Ws.Cells(2, 1) = "No"
        Ws.Cells(2, 2) = "项目专案号"
        Ws.Cells(2, 3) = "Support Type 任务类型(备注)"
        Ws.Cells(2, 4) = "方案类型(备注)"
        Ws.Cells(2, 5) = "Description 需求信息"
        Ws.Cells(2, 6) = "用于产品型号"
        Ws.Cells(2, 7) = "Details 细节要求"
        Ws.Cells(2, 8) = "需求数量"
        Ws.Cells(2, 9) = "类似模具料号"
        Ws.Cells(2, 10) = "类似模具的价格"
        Ws.Cells(2, 11) = "币别"
        Ws.Cells(2, 12) = "类似模具的供应商"
        Ws.Cells(2, 13) = "供应商名称"
        Ws.Cells(2, 14) = "单价"
        Ws.Cells(2, 15) = "币别"
        Ws.Cells(2, 16) = "采购周期"
        Ws.Cells(2, 17) = "单位"

        LineZ = 3
    End Sub
End Class