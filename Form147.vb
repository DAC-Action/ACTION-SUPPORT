Public Class Form147
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim tModel_type = String.Empty
    Dim tModel As String = String.Empty
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form147_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindModel_Type()
        BindModel(tModel)
    End Sub
    Private Sub BindModel_Type()
        Me.ComboBox1.Items.Clear()
        mSQLS1.CommandText = "SELECT * FROM model_type WHERE model_type <> 'Action'"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox1.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub BindModel(ByVal Models1 As String)
        Me.ComboBox2.Items.Clear()
        mSQLS1.CommandText = "select distinct lot.model,model.modelname  from lot,model " _
                          & " where lot.model = model.model and model.model_type <> 'Action'"
        If Not String.IsNullOrEmpty(Models1) Then
            mSQLS1.CommandText += " AND model.model_type = '" & Models1 & "'"
        End If
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString())
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim model_type As String = ComboBox1.Items(ComboBox1.SelectedIndex).ToString()
        BindModel(model_type)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\MES涂装称重报告.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
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

        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value

        tModel = String.Empty
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
        End If
        'BackgroundWorker1.RunWorkerAsync()
        ExportToExcel()
        SaveExcel()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "MES涂装称重报告"
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

    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        Dim xPath As String = "C:\temp\MES涂装称重报告.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        AdjustmentExcelFormat()
        LineZ = 4
        mSQLS1.CommandText = "select model,sn,sum(t1) as t1,sum(t2) as t2,u1,u2 from ( "
        mSQLS1.CommandText += "select lot.model,paravalue.sn,convert(decimal(10,3),isnull((case when station = '0418' then paravalue.value end),0)) as t1 "
        mSQLS1.CommandText += ",convert(decimal(10,4),isnull((case when station = '0587' and paravalue.parameter = '0587_CW_AFTER' then paravalue.value end),0)) as t2,s1.value  as u1,s2.value as u2 "
        mSQLS1.CommandText += "from paravalue left join lot on paravalue.lot = lot.lot left join model_paravalue s1 on lot.model = s1.model and s1.parameter = '0587_CW_MIN' "
        mSQLS1.CommandText += "left join model_paravalue s2 on lot.model = s2.model and s2.parameter = '0587_CW_MAX' "
        mSQLS1.CommandText += "left join (select sn,max(timeout) as tt1 from ( select sn,timeout from tracking where station = '0587' and timeout is not null "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select sn,timeout from tracking_dup where station = '0587' and timeout is not null "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select sn,timeout from scrap_tracking where station = '0587' and timeout is not null ) as ab group by sn ) AC on paravalue.sn = AC.sn "
        mSQLS1.CommandText += "where station in ('0418','0587') and AC.tt1 between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " and lot.model = '" & tModel & "' "
        End If
        mSQLS1.CommandText += ") AS AD group by model,sn,u1,u2 order by sn "
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim Tno As Integer = 1
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = Tno
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t1")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("t2")
                Ws.Cells(LineZ, 6) = "=E" & LineZ & "-D" & LineZ
                Ws.Cells(LineZ, 7) = mSQLReader.Item("u1")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("u2")
                Ws.Cells(LineZ, 9) = "=IF(F" & LineZ & "<G" & LineZ & ",F" & LineZ & "-G" & LineZ & ",IF(F" & LineZ & ">H" & LineZ & ",F" & LineZ & "-H" & LineZ & ",""OK""))"
                Tno += 1
                LineZ += 1
            End While
        End If
        mSQLReader.Close()
    End Sub
    Private Sub AdjustmentExcelFormat()
        Ws.Cells(2, 1) = "MES取数日期/时间:" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "-" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss")
    End Sub
End Class