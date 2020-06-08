Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.Constants
Imports Microsoft.Office.Interop.Excel.XlAutoFillType
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Public Class Form193
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
    Dim LineZ As Integer = 0
    Dim tModel As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")

    Private Sub Form193_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1).AddSeconds(-1)
        
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If
        Dim xPath As String = "C:\temp\Action产品各工段返工及报废率报表.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
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
        tModel = String.Empty
        'If Not IsNothing(ComboBox1.SelectedItem) Then
        'tModel_type = ComboBox1.SelectedItem.ToString()
        'End If
        If Not IsNothing(ComboBox2.SelectedItem) Then
            tModel = ComboBox2.SelectedItem.ToString()
            Dim stCount As Int16 = Strings.InStr(tModel, "|")
            If stCount > 0 Then
                tModel = Strings.Left(tModel, stCount - 1)
            End If
        End If

        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "Action产品各工段返工及报废率报表"
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
                'mConnection.Close()
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
        Dim xPath As String = "C:\temp\Action产品各工段返工及报废率报表.xlsx"
        If Not My.Computer.FileSystem.FileExists(xPath) Then
            MsgBox("NO SAMPLE FILE")
            Return
        End If
        xWorkBook = xExcel.Workbooks.Open(xPath)
        Ws = xWorkBook.Sheets(1)
        LineZ = 5
        Ws.Cells(2, 1) = "取数日期/时间：" & TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "-" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss")


        mSQLS1.CommandText = "Select model,value,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,"
        mSQLS1.CommandText += "sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,"
        mSQLS1.CommandText += "sum(t21) as t21,sum(t22) as t22,sum(t23) as t23,sum(t24) as t24 "
        mSQLS1.CommandText += "from ( select model,value,(case when station = '0330' then count(sn) else 0 end ) as t1,(case when station = '0331' then count(sn) else 0 end ) as t2,"
        mSQLS1.CommandText += "(case when station in ('0380','0530') then count(sn) else 0 end ) as t3,(case when station in ('0490','0620','0627') then count(sn) else 0 end ) as t4,"
        mSQLS1.CommandText += "(case when station in ('0430','0475') then count(sn) else 0 end ) as t5,(case when station in ('0590','0595') then count(sn) else 0 end ) as t6,"
        mSQLS1.CommandText += "(case when station in ('0640','0645') then count(sn) else 0 end ) as t7,(case when station in ('0647','0670') then count(sn) else 0 end ) as t8, 0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16,"
        mSQLS1.CommandText += "0 as t17,0 as t18,0 as t19,0 as t20,0 as t21,0 as t22,0 as t23,0 as t24 "
        mSQLS1.CommandText += "from ( select distinct sn, lot.model,station,value from ( select tracking.lot,tracking.sn,station from tracking  where tracking.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0330','0331','0380','0530','0490','0620','0627','0430','0475','0590','0595','0640','0645','0670') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select tracking_dup.lot,tracking_dup.sn,station from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0330','0331','0380','0530','0490','0620','0627','0430','0475','0590','0595','0640','0645','0670') "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "select scrap_tracking.lot,scrap_tracking.sn,station from scrap_tracking where scrap_tracking.timein between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0330','0331','0380','0530','0490','0620','0627','0430','0475','0590','0595','0640','0645','0670') "
        mSQLS1.CommandText += ") as AB left join lot on ab.lot = lot.lot left join model_paravalue on parameter = 'ERP PN' AND lot.model = model_paravalue.model ) as Ac group by ac.model,ac.station,value "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "Select lot.model,value,0,0,0,0,0,0,0,0,(case when scrap_sn.updatedstation = '0330' then count(scrap.sn) else 0 end) as t9,(case when scrap_sn.updatedstation = '0331' then count(scrap.sn) else 0 end) as t10,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0380','0530') then count(scrap.sn) else 0 end) as t11,(case when scrap_sn.updatedstation in ('0490','0620','0627') then count(scrap.sn) else 0 end) as t12,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0430','0475') then count(scrap.sn) else 0 end) as t13,(case when scrap_sn.updatedstation in ('0590','0595') then count(scrap.sn) else 0 end) as t14,"
        mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0640','0645') then count(scrap.sn) else 0 end) as t15,(case when scrap_sn.updatedstation in ('0670') then count(scrap.sn) else 0 end) as t16,0,0,0,0,0,0,0,0 "
        mSQLS1.CommandText += "from scrap_sn left join lot on scrap_sn.lot = lot.lot left join scrap on scrap_sn.sn = scrap.sn left join model_paravalue on parameter = 'ERP PN' AND lot.model = model_paravalue.model where scrap.datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_sn.updatedstation in ('0330','0331','0380','0530','0490','0620','0627','0430','0475','0590','0595','0640','0645','0670')  "
        'mSQLS1.CommandText += "and scrap.defect not in ('DJ01','DJ02','DL02','DL03','DL04','DL07','DL08', 'DL12', 'DL13','DK05') "
        mSQLS1.CommandText += "and scrap.defect not in ('DJ01','DJ02','DL01','DL02','DL03','DL04','DL05','DL07','DL08','DL09', 'DL12', 'DL13') "
        mSQLS1.CommandText += "group by lot.model,updatedstation,value "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "Select lot.model,value,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,(case when failstation = '0330' then count(sn) else 0 end) as t17,(case when failstation = '0331' then count(sn) else 0 end) as t18,"
        mSQLS1.CommandText += "(case when failstation in ('0380','0530') then count(sn) else 0 end) as t19,(case when failstation in ('0490','0620','0627') then count(sn) else 0 end) as t20,"
        mSQLS1.CommandText += "(case when failstation in ('0430','0475') then count(sn) else 0 end) as t21,(case when failstation in ('0590','0595') then count(sn) else 0 end) as t22,"
        mSQLS1.CommandText += "(case when failstation in ('0640','0645') then count(sn) else 0 end) as t23,(case when failstation in ('0670') then count(sn) else 0 end) as t24 "
        mSQLS1.CommandText += "from failure left join lot on failure.lot = lot.lot left join model_paravalue on parameter = 'ERP PN' AND lot.model = model_paravalue.model where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and failstation in ('0330','0331','0380','0530','0490','0620','0627','0430','0475','0590','0595','0640','0645','0670') "
        mSQLS1.CommandText += "group by lot.model,failstation,value "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "Select lot.model,value,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,(case when failstation = '0330' then count(sn) else 0 end) as t17,(case when failstation = '0331' then count(sn) else 0 end) as t18,"
        mSQLS1.CommandText += "(case when failstation in ('0380','0530') then count(sn) else 0 end) as t19,(case when failstation in ('0490','0620','0627') then count(sn) else 0 end) as t20,"
        mSQLS1.CommandText += "(case when failstation in ('0430','0475') then count(sn) else 0 end) as t21,(case when failstation in ('0590','0595') then count(sn) else 0 end) as t22,"
        mSQLS1.CommandText += "(case when failstation in ('0640','0645') then count(sn) else 0 end) as t23,(case when failstation in ('0670') then count(sn) else 0 end) as t24 "
        mSQLS1.CommandText += "from scrap_failure left join lot on scrap_failure.lot = lot.lot  left join model_paravalue on parameter = 'ERP PN' AND lot.model = model_paravalue.model where failtime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and failstation in ('0330','0331','0380','0530','0490','0620','0627','0430','0475','0590','0595','0640','0645','0670') "
        mSQLS1.CommandText += "group by lot.model,failstation,value ) as AE "
        If Not String.IsNullOrEmpty(tModel) Then
            mSQLS1.CommandText += " WHERE model = '" & tModel & "' "
        End If
        mSQLS1.CommandText += " group by model,value order by model"

                            'mSQLS1.CommandText = "Select model,sum(t1) as t1,sum(t2) as t2,sum(t3) as t3,sum(t4) as t4,sum(t5) as t5,sum(t6) as t6,sum(t7) as t7,sum(t8) as t8,sum(t9) as t9,sum(t10) as t10,"
                            'mSQLS1.CommandText += "sum(t11) as t11,sum(t12) as t12,sum(t13) as t13,sum(t14) as t14,sum(t15) as t15,sum(t16) as t16,sum(t17) as t17,sum(t18) as t18,sum(t19) as t19,sum(t20) as t20,"
                            'mSQLS1.CommandText += "sum(t21) as t21,sum(t22) as t22,sum(t23) as t23,sum(t24) as t24 from ( "
                            'mSQLS1.CommandText += "select model,(case when station = '0330' then count(sn) else 0 end ) as t1,(case when station = '0331' then count(sn) else 0 end ) as t2,"
                            'mSQLS1.CommandText += "(case when station in ('0380','0385','0530') then count(sn) else 0 end ) as t3,(case when station in ('0490','0491','0620','0627') then count(sn) else 0 end ) as t4,"
                            'mSQLS1.CommandText += "(case when station in ('0430','0475') then count(sn) else 0 end ) as t5,(case when station in ('0590','0595') then count(sn) else 0 end ) as t6,"
                            'mSQLS1.CommandText += "(case when station in ('0640','0645') then count(sn) else 0 end ) as t7,(case when station in ('0647','0670') then count(sn) else 0 end ) as t8, 0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16,"
                            'mSQLS1.CommandText += "0 as t17,0 as t18,0 as t19,0 as t20,0 as t21,0 as t22,0 as t23,0 as t24  from ( "
                            'mSQLS1.CommandText += "select distinct sn, lot.model,station from ( select tracking.lot,tracking.sn,station from tracking where tracking.timeout between '"
                            'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0330','0331','0380','0385','0530','0490','0491','0620','0627','0430','0475','0590','0595','0640','0645','0674','0670') "
                            'mSQLS1.CommandText += "union all "
                            'mSQLS1.CommandText += "select tracking_dup.lot,tracking_dup.sn,station from tracking_dup left join lot on tracking_dup.lot = lot.lot where tracking_dup.timeout between '"
                            'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0330','0331','0380','0385','0530','0490','0491','0620','0627','0430','0475','0590','0595','0640','0645','0674','0670') "
                            'mSQLS1.CommandText += "union all "
                            'mSQLS1.CommandText += "select scrap_tracking.lot,scrap_tracking.sn,station from scrap_tracking where scrap_tracking.timein between '"
                            'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and station in ('0330','0331','0380','0385','0530','0490','0491','0620','0627','0430','0475','0590','0595','0640','0645','0674','0670') "
                            'mSQLS1.CommandText += ") as AB left join lot on ab.lot = lot.lot ) as Ac group by ac.model,ac.station "
                            'mSQLS1.CommandText += "union all "
                            'mSQLS1.CommandText += "Select model,0,0,0,0,0,0,0,0,(case when scrap_sn.updatedstation = '0330' then count(scrap.sn) else 0 end) as t9,(case when scrap_sn.updatedstation = '0331' then count(scrap.sn) else 0 end) as t10,"
                            'mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0380','0385','0530') then count(scrap.sn) else 0 end) as t11,(case when scrap_sn.updatedstation in ('0490','0491','0620','0627') then count(scrap.sn) else 0 end) as t12,"
                            'mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0430','0475') then count(scrap.sn) else 0 end) as t13,(case when scrap_sn.updatedstation in ('0590','0595') then count(scrap.sn) else 0 end) as t14,"
                            'mSQLS1.CommandText += "(case when scrap_sn.updatedstation in ('0640','0645') then count(scrap.sn) else 0 end) as t15,(case when scrap_sn.updatedstation in ('0674','0670') then count(scrap.sn) else 0 end) as t16,0,0,0,0,0,0,0,0 "
                            'mSQLS1.CommandText += "from scrap_sn left join lot on scrap_sn.lot = lot.lot left join scrap on scrap_sn.sn = scrap.sn where scrap.datetime between '"
                            'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap_sn.updatedstation in ('0330','0331','0380','0385','0530','0490','0491','0620','0627','0430','0475','0590','0595','0640','0645','0674','0670') "
                            'mSQLS1.CommandText += "group by model,updatedstation "
                            'mSQLS1.CommandText += "union all "
                            'mSQLS1.CommandText += "Select model,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,(case when failstation = '0330' then count(sn) else 0 end) as t17,(case when failstation = '0331' then count(sn) else 0 end) as t18,"
                            'mSQLS1.CommandText += "(case when failstation in ('0380','0385','0530') then count(sn) else 0 end) as t19,(case when failstation in ('0490','0491','0620','0627') then count(sn) else 0 end) as t20,"
                            'mSQLS1.CommandText += "(case when failstation in ('0430','0475') then count(sn) else 0 end) as t21,(case when failstation in ('0500','0595') then count(sn) else 0 end) as t22,"
                            'mSQLS1.CommandText += "(case when failstation in ('0640','0645') then count(sn) else 0 end) as t23,(case when failstation in ('0674','0670') then count(sn) else 0 end) as t24 "
                            'mSQLS1.CommandText += "from failure left join lot on failure.lot = lot.lot where failtime between '"
                            'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and failstation in ('0330','0331','0380','0385','0530','0490','0491','0620','0627','0430','0475','0590','0595','0640','0645','0674','0670') "
                            'mSQLS1.CommandText += "group by model,failstation "
                            'mSQLS1.CommandText += "union all "
                            'mSQLS1.CommandText += "Select model,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,(case when failstation = '0330' then count(sn) else 0 end) as t17,(case when failstation = '0331' then count(sn) else 0 end) as t18,"
                            'mSQLS1.CommandText += "(case when failstation in ('0380','0385','0530') then count(sn) else 0 end) as t19,(case when failstation in ('0490','0491','0620','0627') then count(sn) else 0 end) as t20,"
                            'mSQLS1.CommandText += "(case when failstation in ('0430','0475') then count(sn) else 0 end) as t21,(case when failstation in ('0500','0595') then count(sn) else 0 end) as t22,"
                            'mSQLS1.CommandText += "(case when failstation in ('0640','0645') then count(sn) else 0 end) as t23,(case when failstation in ('0674','0670') then count(sn) else 0 end) as t24 "
                            'mSQLS1.CommandText += "from scrap_failure left join lot on scrap_failure.lot = lot.lot where failtime between '"
                            'mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' and '" & TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and rework <> 'SCRP' and failstation in ('0330','0331','0380','0385','0530','0490','0491','0620','0627','0430','0475','0590','0595','0640','0645','0674','0670') "
                            'mSQLS1.CommandText += "group by model,failstation ) as AE group by model order by model"

        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            Dim DS As Int16 = 1
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = DS
                Ws.Cells(LineZ, 2) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("value")

                For j As Int16 = 1 To 8 Step 1
                    Ws.Cells(LineZ, 10 + (j - 1) * 5) = mSQLReader.Item(1 + j)
                    Ws.Cells(LineZ, 11 + (j - 1) * 5) = mSQLReader.Item(9 + j)
                    Ws.Cells(LineZ, 12 + (j - 1) * 5) = mSQLReader.Item(17 + j)
                    If mSQLReader.Item(1 + j) = 0 Then
                        If mSQLReader.Item(9 + j) = 0 Then
                            Ws.Cells(LineZ, 13 + (j - 1) * 5) = 0
                        Else
                            Ws.Cells(LineZ, 13 + (j - 1) * 5) = 1
                        End If
                        If mSQLReader.Item(17 + j) = 0 Then
                            Ws.Cells(LineZ, 14 + (j - 1) * 5) = 0
                        Else
                            Ws.Cells(LineZ, 14 + (j - 1) * 5) = 1
                        End If
                    Else
                        Ws.Cells(LineZ, 13 + (j - 1) * 5) = Decimal.Divide(mSQLReader.Item(9 + j), mSQLReader.Item(1 + j))
                        Ws.Cells(LineZ, 14 + (j - 1) * 5) = Decimal.Divide(mSQLReader.Item(17 + j), mSQLReader.Item(1 + j))
                    End If

                Next
                Ws.Cells(LineZ, 4) = mSQLReader.Item(2) + mSQLReader.Item(3)
                Ws.Cells(LineZ, 5) = mSQLReader.Item(2) + mSQLReader.Item(3) + mSQLReader.Item(4) + mSQLReader.Item(5) + mSQLReader.Item(6) + mSQLReader.Item(7) + mSQLReader.Item(8) + mSQLReader.Item(9)
                Ws.Cells(LineZ, 6) = mSQLReader.Item(10) + mSQLReader.Item(11) + mSQLReader.Item(12) + mSQLReader.Item(13) + mSQLReader.Item(14) + mSQLReader.Item(15) + mSQLReader.Item(16) + mSQLReader.Item(17)
                Ws.Cells(LineZ, 7) = mSQLReader.Item(18) + mSQLReader.Item(19) + mSQLReader.Item(20) + mSQLReader.Item(21) + mSQLReader.Item(22) + mSQLReader.Item(23) + mSQLReader.Item(24) + mSQLReader.Item(25)
                If mSQLReader.Item(2) + mSQLReader.Item(3) = 0 Then
                    If mSQLReader.Item(10) + mSQLReader.Item(11) + mSQLReader.Item(12) + mSQLReader.Item(13) + mSQLReader.Item(14) + mSQLReader.Item(15) + mSQLReader.Item(16) + mSQLReader.Item(17) > 0 Then
                        Ws.Cells(LineZ, 8) = 1
                    Else
                        Ws.Cells(LineZ, 8) = 0
                    End If
                Else
                    Ws.Cells(LineZ, 8) = Decimal.Divide((mSQLReader.Item(10) + mSQLReader.Item(11) + mSQLReader.Item(12) + mSQLReader.Item(13) + mSQLReader.Item(14) + mSQLReader.Item(15) + mSQLReader.Item(16) + mSQLReader.Item(17)), (mSQLReader.Item(2) + mSQLReader.Item(3)))
                End If
                If (mSQLReader.Item(2) + mSQLReader.Item(3) + mSQLReader.Item(4) + mSQLReader.Item(5) + mSQLReader.Item(6) + mSQLReader.Item(7) + mSQLReader.Item(8) + mSQLReader.Item(9) = 0) Then
                    If (mSQLReader.Item(18) + mSQLReader.Item(19) + mSQLReader.Item(20) + mSQLReader.Item(21) + mSQLReader.Item(22) + mSQLReader.Item(23) + mSQLReader.Item(24) + mSQLReader.Item(25)) = 0 Then
                        Ws.Cells(LineZ, 9) = 0
                    Else
                        Ws.Cells(LineZ, 9) = 1
                    End If
                Else
                    Ws.Cells(LineZ, 9) = Decimal.Divide((mSQLReader.Item(18) + mSQLReader.Item(19) + mSQLReader.Item(20) + mSQLReader.Item(21) + mSQLReader.Item(22) + mSQLReader.Item(23) + mSQLReader.Item(24) + mSQLReader.Item(25)), (mSQLReader.Item(2) + mSQLReader.Item(3) + mSQLReader.Item(4) + mSQLReader.Item(5) + mSQLReader.Item(6) + mSQLReader.Item(7) + mSQLReader.Item(8) + mSQLReader.Item(9)))
                End If

                LineZ += 1
                DS += 1
            End While



            oRng = Ws.Range("B6", Ws.Cells(LineZ, 49))
            oRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
            oRng.Borders(xlEdgeTop).LineStyle = xlContinuous
            oRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
            oRng.Borders(xlEdgeRight).LineStyle = xlContinuous
            oRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            oRng.Borders(xlInsideVertical).LineStyle = xlContinuous
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
        Dim gDatabase As String = Me.ComboBox1.SelectedItem
        If mConnection.State = ConnectionState.Open Then
            mConnection.Close()
        End If
        If gDatabase = "Production" Then
            mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        Else
            mConnection.ConnectionString = Module1.OpenConnectionOfRDMes()
        End If
        'mConnection.ConnectionString = Module1.OpenConnectionOfMes()
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
        BindModel(tModel)
    End Sub
End Class