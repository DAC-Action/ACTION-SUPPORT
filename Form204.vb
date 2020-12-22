Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.Constants
Public Class Form204
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim tModel_type As String
    Dim tModel As String
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form204_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        BindModel_Type()
        Dim Model_Type As String = String.Empty
        BindModel(Model_Type)
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
                Me.ComboBox2.Items.Add(mSQLReader.Item(0).ToString() & "|" & mSQLReader.Item(1).ToString())
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
        tModel_type = String.Empty
        tModel = String.Empty
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If Not IsNothing(ComboBox1.SelectedItem) Then
            tModel_type = ComboBox1.SelectedItem.ToString()
        End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tModel, "|")
            If stCount > 0 Then
                tModel = Strings.Left(tModel, stCount - 1)
            End If
        End If
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Routing_Data"
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
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets(1)
        AdjustExcelFormat()
        mSQLS1.CommandText = "SELECT model.model,model.modelname,model_station_paravalue.cf01,routing.route,routing.seq,routing.station,station.stationname,model_paravalue.value  FROM MODEL "
        mSQLS1.CommandText += "left join model_paravalue on model.model = model_paravalue.model and model_paravalue.parameter = 'Accessory' "
        mSQLS1.CommandText += "LEFT JOIN ROUTING ON MODEL.default_route = ROUTING.ROUTE left join model_station_paravalue on model_station_paravalue.model = model.model and model_station_paravalue.profilename = 'ERP' "
        mSQLS1.CommandText += "and model_station_paravalue.station = routing.station left join station on station.station = model_station_paravalue.station and station.station = routing.station "
        mSQLS1.CommandText += "Where model.default_route not in ( Select distinct route from routing where seq = 1 and station <> '0080' ) and routing.station is not null and cf01 is not null "
        If Not String.IsNullOrEmpty(tModel_type) Then
            mSQLS1.CommandText += "and model.model_type like '" & tModel_type & "' "
        End If
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += "and model.model like '" & tModel & "' "
        End If
        mSQLS1.CommandText += " order by model.model,routing.seq"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelname")
                If Not IsDBNull(mSQLReader.Item("cf01")) Then
                    Dim ERPPN As String = mSQLReader.Item("cf01")
                    Ws.Cells(LineZ, 3) = ERPPN
                    oCommand.CommandText = "Select nvl(imaud02,' ') from ima_file where ima01 = '" & ERPPN & "' "
                    Dim l_imaud02 As String = oCommand.ExecuteScalar()
                    Ws.Cells(LineZ, 4) = l_imaud02
                    If Not String.IsNullOrWhiteSpace(l_imaud02) Then
                        oCommand.CommandText = "Select gem02 from gem_file where gem01 = '" & l_imaud02 & "' "
                        Dim l_gem02 As String = oCommand.ExecuteScalar()
                        Ws.Cells(LineZ, 5) = l_gem02
                    End If
                End If
                Ws.Cells(LineZ, 6) = mSQLReader.Item("route")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("seq")
                Ws.Cells(LineZ, 8) = mSQLReader.Item("station")
                Ws.Cells(LineZ, 9) = mSQLReader.Item("stationname")
                mSQLS2.CommandText = "Select IETime  from ERPSUPPORT.dbo.FIIETIME where ModelID = '" & mSQLReader.Item("model") & "' and StationCode = '" & mSQLReader.Item("station") & "'"
                Dim IET As Decimal = mSQLS2.ExecuteScalar()
                Ws.Cells(LineZ, 10) = IET
                Ws.Cells(LineZ, 11) = mSQLReader.Item("value")
                LineZ += 1
                Label3.Text = LineZ
                Label3.Refresh()
            End While
        End If
        mSQLReader.Close()
        oConnection.Close()

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Name = "Routing资料"
        oRng = Ws.Range("A1", "H1")
        oRng.EntireColumn.ColumnWidth = 18.2
        Ws.Cells(1, 1) = "model"
        Ws.Cells(1, 2) = "modelname"
        Ws.Cells(1, 3) = "ERP"
        Ws.Cells(1, 4) = "ERP生产部门代码"
        Ws.Cells(1, 5) = "ERP生产部门名称"
        Ws.Cells(1, 6) = "Routing"
        Ws.Cells(1, 7) = "seq"
        Ws.Cells(1, 8) = "station"
        Ws.Cells(1, 9) = "stationname"
        Ws.Cells(1, 10) = "MES作业站标准人工工时"
        Ws.Cells(1, 11) = "Accessory"
        oRng = Ws.Range("H1", "H1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        LineZ = 2
    End Sub
End Class