Public Class Form207
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim Value1 As String = String.Empty
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form207_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
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
        Value1 = TextBox1.Text
        If String.IsNullOrEmpty(Value1) Then
            MsgBox("请输入模具号")
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
        mSQLS1.CommandText = "Select sum(c1) from ( Select count(*) as c1 from paravalue left join sn on paravalue.sn = sn.sn where value like '"
        mSQLS1.CommandText += Value1 & "%' and parameter = 'MOLD_ID' "
        mSQLS1.CommandText += "Union all "
        mSQLS1.CommandText += "Select count(*) from scrap_paravalue left join scrap_sn on scrap_paravalue.sn = scrap_sn.sn where value like '"
        mSQLS1.CommandText += Value1 & "%' and parameter = 'MOLD_ID' ) as ab"
        Dim HasR As Decimal = mSQLS1.ExecuteScalar()

        If HasR > 0 Then
            BackgroundWorker1.RunWorkerAsync()
        Else
            MsgBox("无资料")
            Return
        End If

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Try
            Ws.Name = Value1
        Catch ex As Exception

        End Try

        AdjustExcelFormat()

        mSQLS1.CommandText = "Select paravalue.sn, (case when sn.topreworkstation = '' or sn.topreworkstation is null then updatedstation  else topreworkstation end) as c1 from paravalue left join sn on paravalue.sn = sn.sn where value LIKE '"
        mSQLS1.CommandText += Value1 & "%' and parameter = 'MOLD_ID' "
        mSQLS1.CommandText += "Union all "
        mSQLS1.CommandText += "Select scrap_paravalue.sn, (case when scrap_sn.topreworkstation = '' or scrap_sn.topreworkstation is null then updatedstation  else topreworkstation end) as c1 from scrap_paravalue left join scrap_sn on scrap_paravalue.sn = scrap_sn.sn where value LIKE '"
        mSQLS1.CommandText += Value1 & "%' and parameter = 'MOLD_ID' "

        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim SN1 As Int16 = 1
            While mSQLReader.Read()
                For i As Int16 = 0 To mSQLReader.FieldCount - 1 Step 1
                    Ws.Cells(LineZ, i + 1) = mSQLReader.Item(i)
                Next
                LineZ += 1
                End While
            End If
        mSQLReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 25

        Ws.Cells(1, 1) = "SN"
        Ws.Cells(1, 2) = "所在工站"

        oRng = Ws.Range("A1", "B1")
        oRng.EntireColumn.NumberFormat = "@"
        LineZ = 2
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "模具使用记录"
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
End Class