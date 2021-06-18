Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form206
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLS3 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim mSQLReader2 As SqlClient.SqlDataReader

    Dim hConnection As New SqlClient.SqlConnection
    Dim hSQLS1 As New SqlClient.SqlCommand
    Dim hSQLReader As SqlClient.SqlDataReader

    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader

    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form206_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1)
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        hConnection.ConnectionString = Module1.OpenConnectionOfHR()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
                mSQLS2.CommandTimeout = 600
                mSQLS3.Connection = mConnection
                mSQLS3.CommandType = CommandType.Text
                mSQLS3.CommandTimeout = 600
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If hConnection.State <> ConnectionState.Open Then
            Try
                hConnection.Open()
                hSQLS1.Connection = hConnection
                hSQLS1.CommandType = CommandType.Text
            Catch ex As Exception

            End Try
        End If

        If oConnection.State <> ConnectionState.Open Then
            oConnection.Open()
            oCommand.Connection = oConnection
            oCommand.CommandType = CommandType.Text
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
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Name = "Detail"
        AdjustExcelFormat()

        mSQLS1.CommandText = "Select Scrap.datetime , lot.model, TaktTime.Degree,Scrap.sn , Scrap.defect +' '+ defect.desc_th as c1, S1.value as c2 , S2.value as c3 , TaktTime.TaktTimeVal    from scrap "
        mSQLS1.CommandText += "left join lot on scrap.lot = lot.lot left join EFD.dbo.TaktTime on lot.model = TaktTime.ModelId and TaktTime.StationGroupId = 13 Left join defect on Scrap.defect = defect.defect "
        mSQLS1.CommandText += "left join scrap_paravalue S1 on scrap.sn = s1.sn and s1.parameter = 'MOLD_ID' and station in ('0150','0151') Left join scrap_paravalue S2 on Scrap.sn = S2.sn  and S2.parameter = 'EM_ID'  and S2.station <> '0390' "
        mSQLS1.CommandText += "left join scrap_sn S9 on scrap.sn = s9.sn "
        mSQLS1.CommandText += "Where Scrap.datetime between '" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and s9.updatedstation in ('0330','0331') order by sn"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim SN1 As Int16 = 1
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = SN1
                Ws.Cells(LineZ, 2) = mSQLReader.Item("datetime")
                'Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("Degree")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("c1")
                Ws.Cells(LineZ, 9) = mSQLReader.Item("c2")
                Ws.Cells(LineZ, 10) = mSQLReader.Item("c3")
                Ws.Cells(LineZ, 16) = mSQLReader.Item("TaktTimeVal")
                Ws.Cells(LineZ, 14) = "=L" & LineZ & "/M" & LineZ
                GetLayupData(mSQLReader.Item("sn"))
                SN1 += 1
                LineZ += 1
            End While
        End If
        mSQLReader.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Columns.EntireColumn.ColumnWidth = 25
        Ws.Columns.EntireColumn.HorizontalAlignment = xlCenter
        Ws.Columns.EntireColumn.VerticalAlignment = xlCenter
        
        Ws.Cells(1, 2) = "检验日期"
        'Ws.Cells(1, 3) = "产品型号"
        Ws.Cells(1, 3) = "難度係數"
        Ws.Cells(1, 4) = "产品序列号"
        Ws.Cells(1, 5) = "不良现象"
        Ws.Cells(1, 6) = "预型作业员"
        Ws.Cells(1, 7) = "生产日期"
        Ws.Cells(1, 8) = "新/老员工"
        Ws.Cells(1, 9) = "使用模具編號"
        Ws.Cells(1, 10) = "生產設備編號"
        Ws.Cells(1, 11) = "白班/夜班"
        Ws.Cells(1, 12) = "個人此款產品報廢數"
        Ws.Cells(1, 13) = "個人此款產品製作總量"
        Ws.Cells(1, 14) = "個人此款產品總報廢率"
        Ws.Cells(1, 15) = "個人報廢累計總金額"
        Ws.Cells(1, 16) = "標準工時"
        Ws.Cells(1, 17) = "個人累計報廢工時"

        oRng = Ws.Range("B1", "Q1")
        oRng.Interior.Color = Color.LightBlue

        oRng = Ws.Range("N1", "N1")
        oRng.EntireColumn.NumberFormat = "0.00%"

        oRng = Ws.Range("O1", "O1")
        oRng.EntireColumn.NumberFormat = "0.00"

        LineZ = 2
    End Sub
    Private Sub GetLayupData(ByVal sn As String)
        mSQLS2.CommandText = "Select scrap_tracking.users,users.name,   timeout, lot.model  from scrap_tracking left join users on scrap_tracking.users = users.id left join lot on scrap_tracking.lot = lot.lot  where sn = '" & sn & "' and station in ('0150','0151')"
        mSQLReader2 = mSQLS2.ExecuteReader()
        Dim Date1 As Date = Now()
        Dim Model1 As String = String.Empty
        Dim Users1 As String = String.Empty
        If mSQLReader2.HasRows() Then
            While mSQLReader2.Read()
                Ws.Cells(LineZ, 6) = mSQLReader2.Item("users") & " " & mSQLReader2.Item("name")
                Ws.Cells(LineZ, 7) = mSQLReader2.Item("timeout")
                Date1 = mSQLReader2.Item("timeout")
                Model1 = mSQLReader2.Item("model")
                Users1 = mSQLReader2.Item("users")
            End While
        End If
        mSQLReader2.Close()

        GetHRData(Users1, Date1)

        GetScrapData(Users1, Model1)

    End Sub
    Private Sub GetHRData(ByVal empno As String, ByVal Date1 As Date)
        If empno.StartsWith("0") Then
            empno = Strings.Right(empno, 4)
        End If
        hSQLS1.CommandText = "Select HireDate from T_EMP_Employee where empcode = '" & empno & "'"
        Dim HD1 As Date = hSQLS1.ExecuteScalar()
        If HD1.AddDays(30) > Date1 Then
            Ws.Cells(LineZ, 8) = "新员工"
        Else
            Ws.Cells(LineZ, 8) = "老员工"
        End If
        Dim Date2 As Date = Date1.AddHours(-8).Date

        hSQLS1.CommandText = "Select ShiftCode   from T_ATD_AttDaily left join T_EMP_Employee on T_ATD_AttDaily.EmpID = T_EMP_Employee.id "
        hSQLS1.CommandText += " where empcode = '" & empno & "' and AttDate = '" & Date2.ToString("yyyy/MM/dd") & "'"
        Dim SFC As String = hSQLS1.ExecuteScalar()
        Select Case SFC
            Case "01", "02", "03", "06", "13", "18", "18-1", "18-2", "23-1", "23-2", "23-3", "28"
                Ws.Cells(LineZ, 11) = "白班"
            Case "24-1", "24-2", "24-3", "24-4", "25", "25-1", "25-2", "25-3"
                Ws.Cells(LineZ, 11) = "夜班"
        End Select


    End Sub

    Private Sub GetScrapData(ByVal users1 As String, ByVal Model1 As String)
        mSQLS2.CommandText = "Select isnull(count(*),0) from scrap_tracking left join lot on scrap_tracking.lot = lot.lot where scrap_tracking.users = '" & users1 & "' and station in ('0150','0151') and lot.model = '" & Model1 & "'"
        Dim SS1 As Integer = mSQLS2.ExecuteScalar()
        Ws.Cells(LineZ, 12) = SS1

        Ws.Cells(LineZ, 17) = "=" & SS1 & "*P" & LineZ

        mSQLS2.CommandText = "Select sum(t1) from ( Select isnull(count(*),0) as t1 from scrap_tracking left join lot on scrap_tracking.lot = lot.lot "
        mSQLS2.CommandText += "where scrap_tracking.users = '" & users1 & "' and station in ('0150','0151') and lot.model = '" & Model1 & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "Select isnull(count(*),0) from tracking left join lot on tracking.lot = lot.lot "
        mSQLS2.CommandText += "where tracking.users = '" & users1 & "' and station in ('0150','0151') and lot.model = '" & Model1 & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "Select isnull(count(*),0) from tracking_dup left join lot on tracking_dup.lot = lot.lot "
        mSQLS2.CommandText += "where tracking_dup.users = '" & users1 & "' and station in ('0150','0151') and lot.model = '" & Model1 & "' ) as ab"

        Dim SS2 As Integer = mSQLS2.ExecuteScalar()
        Ws.Cells(LineZ, 13) = SS2

        If SS1 > 0 Then
            mSQLS2.CommandText = "Select distinct cf01 from model_station_paravalue where model = '" & Model1 & "' and station in ('0150','0151')"
            Dim ERPPN As String = mSQLS2.ExecuteScalar()

            oCommand.CommandText = "Select nvl(ccc23,0) from ccc_file where ccc01 = '" & ERPPN & "' order by (case when length(ccc03) = 1 then  ccc02||'0'||ccc03 else ccc02 || ccc03 end) desc"
            Dim SS3 As Decimal = 0
            SS3 = oCommand.ExecuteScalar()
            If IsDBNull(SS3) Then
                SS3 = 0
            End If
            Ws.Cells(LineZ, 15) = "=" & SS1 & "*" & SS3
        Else
            Ws.Cells(LineZ, 15) = 0
        End If
        
    End Sub
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "QRQC统计分析报告"
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