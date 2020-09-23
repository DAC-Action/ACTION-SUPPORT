Public Class Form198
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    'Dim tStation1 As String
    Dim ptime As String = String.Empty
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim TimeS3 As DateTime  ' 出勤日期用
    Dim TimeS4 As DateTime  ' 出勤日期用
    Dim TS As Decimal = 0
    Dim xExcel As Microsoft.Office.Interop.Excel.Application
    Dim xWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim Ws As Microsoft.Office.Interop.Excel.Worksheet
    Dim oRng As Microsoft.Office.Interop.Excel.Range
    Dim LineZ As Integer = 0
    Dim Sector As Int16 = 0
    Dim TempDB As String = String.Empty
    'Dim LastStation As String = String.Empty
    'Dim ERPPN As String = String.Empty
    'Dim tModel As String = String.Empty
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.BackgroundWorker1.IsBusy() Then
            MsgBox("处理中，请等待")
            Return
        End If

        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.ConnectionString = Module1.OpenConnectionOfMes()
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
                mSQLS1.CommandTimeout = 600
                mSQLS2.Connection = mConnection
                mSQLS2.CommandType = CommandType.Text
                mSQLS2.CommandTimeout = 600
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        TimeS1 = DateTimePicker1.Value
        TimeS2 = DateTimePicker2.Value
        TS = DateDiff(DateInterval.Day, TimeS1, TimeS2)
        TimeS3 = TimeS1
        If TS > 0 Then
            TimeS4 = TimeS3.AddDays(TS - 1)
        End If
        Sector = Me.ComboBox1.SelectedIndex
        If Sector = -1 Then
            MsgBox("请选择工段")
            Return
        End If

        TempDB = "F196" & Now.ToString("yyyyMMddHHmmssff")
        mSQLS2.CommandText = "CREATE TABLE ERPSUPPORT.dbo." & TempDB & " (WorkID nvarchar(5) Not null, WorkName nvarchar(50) Not null, ModelID nvarchar(50) Not null , ERPPN nvarchar(50), WorkDept nvarchar(50), T1 numeric(18, 2), T2 numeric(18, 2), T3 numeric(18, 2)) "
        'mSQLS2.CommandText = "DELETE ERPSUPPORT.dbo.Form198DB "
        Try
            mSQLS2.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Form198_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
        Me.DateTimePicker2.Value = Convert.ToDateTime(ptime).AddDays(1)
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExportToExcel()
    End Sub
    Private Sub ExportToExcel()
        xExcel = New Microsoft.Office.Interop.Excel.Application
        xWorkBook = xExcel.Workbooks.Add()
        Ws = xWorkBook.Sheets.Add()
        Ws = xWorkBook.Sheets(1)
        Ws.Activate()
        Ws.Name = "Detail"
        AdjustExcelFormat()
        LineZ = 2
        mSQLS1.CommandText = "Select lot.model,model.modelname ,lot.lot, tracking.sn, tracking.station, station.stationname_cn , cf01,timein, timeout, tracking.users , s2.UserID ,"
        mSQLS1.CommandText += "(Select count(*) from MultipleUserRecord s3 where tracking.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) as IETime , s5.name , fresh "
        mSQLS1.CommandText += "from tracking left join lot on tracking.lot = lot.lot left join model on lot.model = model.model left join station on tracking.station = station.station "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and tracking.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join MultipleUserRecord s2 on tracking.Id = s2.TrackingID left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and tracking.station = s4.StationCode "
        mSQLS1.CommandText += "left join users s5 on (case when s2.UserID is null then tracking.users else s2.UserID  end ) = s5.id where tracking.station in ("
        Select Case Sector
            Case 0
                mSQLS1.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS1.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS1.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS1.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS1.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS1.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS1.CommandText += ") and timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "Select lot.model,model.modelname ,lot.lot, tracking_dup.sn, tracking_dup.station, station.stationname_cn , cf01,timein, timeout, tracking_dup.users , s2.UserID ,"
        mSQLS1.CommandText += "(Select count(*) from MultipleUserRecord s3 where tracking_dup.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) , s5.name , fresh "
        mSQLS1.CommandText += "from tracking_dup left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model left join station on tracking_dup.station = station.station "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and tracking_dup.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join MultipleUserRecord s2 on tracking_dup.Id = s2.TrackingID  left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and tracking_dup.station = s4.StationCode "
        mSQLS1.CommandText += "left join users s5 on (case when s2.UserID is null then tracking_dup.users else s2.UserID  end ) = s5.id where tracking_dup.station in ("
        Select Case Sector
            Case 0
                mSQLS1.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS1.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS1.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS1.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS1.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS1.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS1.CommandText += ") and timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "Select lot.model,model.modelname ,lot.lot, scrap_tracking.sn, scrap_tracking.station, station.stationname_cn , cf01,timein, timeout, scrap_tracking.users , s2.UserID ,"
        mSQLS1.CommandText += "(Select count(*) from MultipleUserRecord s3 where scrap_tracking.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) , s5.name , fresh "
        mSQLS1.CommandText += "from scrap_tracking left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model left join station on scrap_tracking.station = station.station "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and scrap_tracking.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join MultipleUserRecord s2 on scrap_tracking.Id = s2.TrackingID left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and scrap_tracking.station = s4.StationCode "
        mSQLS1.CommandText += "left join users s5 on (case when s2.UserID is null then scrap_tracking.users else s2.UserID  end ) = s5.id where scrap_tracking.station in ("
        Select Case Sector
            Case 0
                mSQLS1.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS1.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS1.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS1.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS1.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS1.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS1.CommandText += ") and timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "'"

        mSQLReader = mSQLS1.ExecuteReader
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("modelname")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("lot")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("sn")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("station")
                Ws.Cells(LineZ, 6) = mSQLReader.Item("stationname_cn")
                Ws.Cells(LineZ, 7) = mSQLReader.Item("cf01")
                Ws.Cells(LineZ, 8) = GetDept(mSQLReader.Item("cf01"))
                If mSQLReader.Item("count1") = 0 Then
                    Ws.Cells(LineZ, 9) = mSQLReader.Item("IETime")
                Else
                    Ws.Cells(LineZ, 9) = Decimal.Round(Decimal.Divide(mSQLReader.Item("IETime"), mSQLReader.Item("count1")), 2)
                End If
                Ws.Cells(LineZ, 10) = mSQLReader.Item("timein")
                Ws.Cells(LineZ, 11) = mSQLReader.Item("timeout")
                If IsDBNull(mSQLReader.Item("UserID")) Then
                    Ws.Cells(LineZ, 12) = mSQLReader.Item("users")
                    Ws.Cells(LineZ, 14) = GetUserDept(mSQLReader.Item("users"))
                    Ws.Cells(LineZ, 15) = GetUserLevel(mSQLReader.Item("users"))
                Else
                    Ws.Cells(LineZ, 12) = mSQLReader.Item("UserID")
                    Ws.Cells(LineZ, 14) = GetUserDept(mSQLReader.Item("UserID"))
                    Ws.Cells(LineZ, 15) = GetUserLevel(mSQLReader.Item("UserID"))
                End If
                Ws.Cells(LineZ, 13) = mSQLReader.Item("name")
                Ws.Cells(LineZ, 16) = mSQLReader.Item("fresh")
                If mSQLReader.Item("count1") = 0 Then
                    Ws.Cells(LineZ, 17) = 1
                Else
                    Ws.Cells(LineZ, 17) = mSQLReader.Item("count1")
                End If

                LineZ += 1
                Label3.Text = "1-" & LineZ
                Label3.Refresh()
            End While
        End If
        mSQLReader.Close()

        oRng = Ws.Range("A1", "Q1")
        oRng.EntireColumn.AutoFit()

        Ws = xWorkBook.Sheets(2)
        Ws.Activate()
        Ws.Name = "Summary"
        AdjustExcelFormat1()
        LineZ = 3

        mSQLS1.CommandText = "Select (Case when UserID is null then users else UserID end) as c1,name, model, cf01, isnull(sum(case when count1 = 0 then Round(IETime,2) else round(ietime/count1, 2) end),0) as t1 , "
        Select Case Sector
            Case 0
                mSQLS1.CommandText += "isnull(sum( case when station = '0151' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t2,isnull(sum( case when station = '0150' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t3,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0160' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t4,isnull(sum( case when station = '0170' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t5 "
            Case 1
                mSQLS1.CommandText += "isnull(sum( case when station = '0370' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t2,isnull(sum( case when station = '0360' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t3,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0520' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t4,isnull(sum( case when station = '0335' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t5,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0350' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t6,isnull(sum( case when station = '0340' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t7,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0500' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t8,isnull(sum( case when station = '0510' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t9 "
            Case 2
                mSQLS1.CommandText += "isnull(sum( case when station = '0400' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t2,isnull(sum( case when station = '0479' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t3,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0480' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t4,isnull(sum( case when station = '0485' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t5,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0492' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t6,isnull(sum( case when station = '0610' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t7,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0623' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t8 "
            Case 3
                mSQLS1.CommandText += "isnull(sum( case when station = '0418' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t2,isnull(sum( case when station = '0410' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t3,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0415' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t4,isnull(sum( case when station = '0416' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t5,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0417' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t6,isnull(sum( case when station = '0420' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t7,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0440' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t8,isnull(sum( case when station = '0445' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t9,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0450' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t10,isnull(sum( case when station = '0460' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t11,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0465' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t12,isnull(sum( case when station = '0470' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t13,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0540' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t14,isnull(sum( case when station = '0545' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t15,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0550' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t16,isnull(sum( case when station = '0570' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t17,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0575' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t18,isnull(sum( case when station = '0580' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t19,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0583' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t20,isnull(sum( case when station = '0584' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t21,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0585' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t22,isnull(sum( case when station = '0560' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t23,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0408' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t24,isnull(sum( case when station = '0413' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t25,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0422' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t26 "
            Case 4
                mSQLS1.CommandText += "isnull(sum( case when station = '0625' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t2,isnull(sum( case when station = '0630' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t3,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0635' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t4,isnull(sum( case when station = '0642' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t5,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0650' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t6,isnull(sum( case when station = '0658' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t7,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0665' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t8,isnull(sum( case when station = '0666' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t9,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0669' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t10,isnull(sum( case when station = '0671' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t11,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0675' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t12,isnull(sum( case when station = '0680' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t13 "
            Case 5
                mSQLS1.CommandText += "isnull(sum( case when station = '0142' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t2,isnull(sum( case when station = '0145' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t3,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0148' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t4,isnull(sum( case when station = '0173' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t5,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0177' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t6,isnull(sum( case when station = '0175' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t7,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0190' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t8,isnull(sum( case when station = '0195' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t9,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0200' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t10,isnull(sum( case when station = '0215' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t11,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0223' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t12,isnull(sum( case when station = '0225' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t13,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0230' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t14,isnull(sum( case when station = '0231' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t15,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0240' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t16,isnull(sum( case when station = '0250' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t17,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0255' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t18,isnull(sum( case when station = '0260' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t19,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0280' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t20,isnull(sum( case when station = '0300' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t21,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0315' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t22,isnull(sum( case when station = '0320' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t23,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0321' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t24,isnull(sum( case when station = '0325' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t25,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0326' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t26,isnull(sum( case when station = '0333' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t27,"
                mSQLS1.CommandText += "isnull(sum( case when station = '0390' then (case when count1 = 0 then round(ietime, 2) else round(ietime/count1, 2) end ) else 0 end),0) as t28 "
        End Select
        mSQLS1.CommandText += "from ( "
        mSQLS1.CommandText += "Select lot.model,model.modelname ,lot.lot, tracking.sn, tracking.station, station.stationname_cn , cf01,timein, timeout, tracking.users , s2.UserID ,"
        mSQLS1.CommandText += "(Select count(*) from MultipleUserRecord s3 where tracking.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) as IETime , s5.name , fresh "
        mSQLS1.CommandText += "from tracking left join lot on tracking.lot = lot.lot left join model on lot.model = model.model left join station on tracking.station = station.station "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and tracking.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join MultipleUserRecord s2 on tracking.Id = s2.TrackingID left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and tracking.station = s4.StationCode "
        mSQLS1.CommandText += "left join users s5 on (case when s2.UserID is null then tracking.users else s2.UserID  end ) = s5.id where tracking.station in ("
        Select Case Sector
            Case 0
                mSQLS1.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS1.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS1.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS1.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS1.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS1.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS1.CommandText += ") and timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "Select lot.model,model.modelname ,lot.lot, tracking_dup.sn, tracking_dup.station, station.stationname_cn , cf01,timein, timeout, tracking_dup.users , s2.UserID ,"
        mSQLS1.CommandText += "(Select count(*) from MultipleUserRecord s3 where tracking_dup.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) , s5.name , fresh "
        mSQLS1.CommandText += "from tracking_dup left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model left join station on tracking_dup.station = station.station "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and tracking_dup.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join MultipleUserRecord s2 on tracking_dup.Id = s2.TrackingID  left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and tracking_dup.station = s4.StationCode "
        mSQLS1.CommandText += "left join users s5 on (case when s2.UserID is null then tracking_dup.users else s2.UserID  end ) = s5.id where tracking_dup.station in ("
        Select Case Sector
            Case 0
                mSQLS1.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS1.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS1.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS1.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS1.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS1.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS1.CommandText += ") and timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS1.CommandText += "union all "
        mSQLS1.CommandText += "Select lot.model,model.modelname ,lot.lot, scrap_tracking.sn, scrap_tracking.station, station.stationname_cn , cf01,timein, timeout, scrap_tracking.users , s2.UserID ,"
        mSQLS1.CommandText += "(Select count(*) from MultipleUserRecord s3 where scrap_tracking.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) , s5.name , fresh "
        mSQLS1.CommandText += "from scrap_tracking left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model left join station on scrap_tracking.station = station.station "
        mSQLS1.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and scrap_tracking.station = model_station_paravalue.station "
        mSQLS1.CommandText += "left join MultipleUserRecord s2 on scrap_tracking.Id = s2.TrackingID left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and scrap_tracking.station = s4.StationCode "
        mSQLS1.CommandText += "left join users s5 on (case when s2.UserID is null then scrap_tracking.users else s2.UserID  end ) = s5.id where scrap_tracking.station in ("
        Select Case Sector
            Case 0
                mSQLS1.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS1.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS1.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS1.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS1.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS1.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS1.CommandText += ") and timeout between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as XX Group by UserID, users  , name, model, cf01 order by c1"

        mSQLReader = mSQLS1.ExecuteReader
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("c1")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("name")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("model")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("cf01")
                Dim l_Dept As String = GetDept(mSQLReader.Item("cf01"))
                Ws.Cells(LineZ, 5) = l_Dept
                Dim THour As Decimal = 0
                If TS = 0 Then
                    THour = 0
                Else
                    THour = GetGZHour(mSQLReader.Item("c1")) * 60
                End If

                Ws.Cells(LineZ, 8) = THour
                Dim TTS As Decimal = Calculate1(mSQLReader.Item("c1"))
                Dim l_TTS As Decimal = 0
                If TTS <> 0 Then
                    l_TTS = Decimal.Round(mSQLReader.Item("t1") / TTS * THour, 2)
                    Ws.Cells(LineZ, 9) = l_TTS
                End If
                Ws.Cells(LineZ, 10) = mSQLReader.Item("t1")
                Select Case Sector
                    Case 0
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("t5")
                    Case 1
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("t5")
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("t6")
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("t7")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("t8")
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("t9")
                    Case 2
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("t5")
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("t6")
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("t7")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("t8")
                    Case 3
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("t5")
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("t6")
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("t7")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("t8")
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("t9")
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("t10")
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("t11")
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("t12")
                        Ws.Cells(LineZ, 22) = mSQLReader.Item("t13")
                        Ws.Cells(LineZ, 23) = mSQLReader.Item("t14")
                        Ws.Cells(LineZ, 24) = mSQLReader.Item("t15")
                        Ws.Cells(LineZ, 25) = mSQLReader.Item("t16")
                        Ws.Cells(LineZ, 26) = mSQLReader.Item("t17")
                        Ws.Cells(LineZ, 27) = mSQLReader.Item("t18")
                        Ws.Cells(LineZ, 28) = mSQLReader.Item("t19")
                        Ws.Cells(LineZ, 29) = mSQLReader.Item("t20")
                        Ws.Cells(LineZ, 30) = mSQLReader.Item("t21")
                        Ws.Cells(LineZ, 31) = mSQLReader.Item("t22")
                        Ws.Cells(LineZ, 32) = mSQLReader.Item("t23")
                        Ws.Cells(LineZ, 33) = mSQLReader.Item("t24")
                        Ws.Cells(LineZ, 34) = mSQLReader.Item("t25")
                        Ws.Cells(LineZ, 35) = mSQLReader.Item("t26")
                    Case 4
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("t5")
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("t6")
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("t7")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("t8")
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("t9")
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("t10")
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("t11")
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("t12")
                        Ws.Cells(LineZ, 22) = mSQLReader.Item("t13")
                    Case 5
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("t5")
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("t6")
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("t7")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("t8")
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("t9")
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("t10")
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("t11")
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("t12")
                        Ws.Cells(LineZ, 22) = mSQLReader.Item("t13")
                        Ws.Cells(LineZ, 11) = mSQLReader.Item("t2")
                        Ws.Cells(LineZ, 12) = mSQLReader.Item("t3")
                        Ws.Cells(LineZ, 13) = mSQLReader.Item("t4")
                        Ws.Cells(LineZ, 14) = mSQLReader.Item("t5")
                        Ws.Cells(LineZ, 15) = mSQLReader.Item("t6")
                        Ws.Cells(LineZ, 16) = mSQLReader.Item("t7")
                        Ws.Cells(LineZ, 17) = mSQLReader.Item("t8")
                        Ws.Cells(LineZ, 18) = mSQLReader.Item("t9")
                        Ws.Cells(LineZ, 19) = mSQLReader.Item("t10")
                        Ws.Cells(LineZ, 20) = mSQLReader.Item("t11")
                        Ws.Cells(LineZ, 21) = mSQLReader.Item("t12")
                        Ws.Cells(LineZ, 22) = mSQLReader.Item("t13")
                        Ws.Cells(LineZ, 23) = mSQLReader.Item("t14")
                        Ws.Cells(LineZ, 24) = mSQLReader.Item("t15")
                        Ws.Cells(LineZ, 25) = mSQLReader.Item("t16")
                        Ws.Cells(LineZ, 26) = mSQLReader.Item("t17")
                        Ws.Cells(LineZ, 27) = mSQLReader.Item("t18")
                        Ws.Cells(LineZ, 28) = mSQLReader.Item("t19")
                        Ws.Cells(LineZ, 29) = mSQLReader.Item("t20")
                        Ws.Cells(LineZ, 30) = mSQLReader.Item("t21")
                        Ws.Cells(LineZ, 31) = mSQLReader.Item("t22")
                        Ws.Cells(LineZ, 32) = mSQLReader.Item("t23")
                        Ws.Cells(LineZ, 33) = mSQLReader.Item("t24")
                        Ws.Cells(LineZ, 34) = mSQLReader.Item("t25")
                        Ws.Cells(LineZ, 35) = mSQLReader.Item("t26")
                        Ws.Cells(LineZ, 36) = mSQLReader.Item("t27")
                        Ws.Cells(LineZ, 37) = mSQLReader.Item("t28")
                End Select
                
                'mSQLS2.CommandText = "INSERT INTO ERPSUPPORT.dbo.Form198DB VALUES ('" & mSQLReader.Item("c1") & "','" & mSQLReader.Item("name") & "','" & mSQLReader.Item("model") & "','" & mSQLReader.Item("cf01") & "','" & l_Dept & "'," & THour & "," & l_TTS & "," & mSQLReader.Item("t1") & ")"
                mSQLS2.CommandText = "INSERT INTO ERPSUPPORT.dbo." & TempDB & " VALUES ('" & mSQLReader.Item("c1") & "','" & mSQLReader.Item("name") & "','" & mSQLReader.Item("model") & "','" & mSQLReader.Item("cf01") & "','" & l_Dept & "'," & THour & "," & l_TTS & "," & mSQLReader.Item("t1") & ")"
                Try
                    mSQLS2.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message())
                End Try
                LineZ += 1
                Label3.Text = "2-" & LineZ
                Label3.Refresh()
            End While
        End If
        mSQLReader.Close()
        Select Case Sector
            Case 0
                oRng = Ws.Range("A1", "N1")
            Case 1
                oRng = Ws.Range("A1", "R1")
            Case 2
                oRng = Ws.Range("A1", "Q1")
            Case 3
                oRng = Ws.Range("A1", "AI1")
            Case 4
                oRng = Ws.Range("A1", "V1")
            Case 5
                oRng = Ws.Range("A1", "AK1")
        End Select

        oRng.EntireColumn.AutoFit()

        ' 20200727
        Ws = xWorkBook.Sheets(3)
        Ws.Activate()
        Ws.Name = "员工效率情况"
        AdjustExcelFormat2()
        LineZ = 2

        'mSQLS1.CommandText = "Select WorkID, WorkName, SUM(T3) as t3, sum(t2) as t2 from ERPSUPPORT.dbo.Form198DB group by Workid, WorkName Order by WorkID "
        mSQLS1.CommandText = "Select WorkID, WorkName, SUM(T3) as t3, sum(t2) as t2 from ERPSUPPORT.dbo." & TempDB & " group by Workid, WorkName Order by WorkID "
        mSQLReader = mSQLS1.ExecuteReader
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("WorkID")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("WorkName")
                Ws.Cells(LineZ, 3) = GetDeptbyHR(mSQLReader.Item("WorkID"))
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t3")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("t2")
                Ws.Cells(LineZ, 6) = "=IFERROR(D" & LineZ & "/E" & LineZ & ",)"
                LineZ += 1
                Label3.Text = "3-" & LineZ
                Label3.Refresh()
            End While
            Ws.Cells(LineZ, 3) = "合计"
            Ws.Cells(LineZ, 4) = "=SUM(D2:D" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 5) = "=SUM(E2:E" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 6) = "=IFERROR(D" & LineZ & "/E" & LineZ & ",)"
        End If
        mSQLReader.Close()

        oRng = Ws.Range("A1", "F1")
        oRng.EntireColumn.AutoFit()

        Ws = xWorkBook.Sheets(4)
        Ws.Activate()
        Ws.Name = "产品效率情况"
        AdjustExcelFormat3()
        LineZ = 2

        'mSQLS1.CommandText = "Select ModelID, ERPPN, WorkDept, SUM(T3) as t3, sum(t2) as t2 from ERPSUPPORT.dbo.Form198DB group by ModelID, ERPPN, WorkDept Order by ModelID"
        mSQLS1.CommandText = "Select ModelID, ERPPN, WorkDept, SUM(T3) as t3, sum(t2) as t2 from ERPSUPPORT.dbo." & TempDB & " group by ModelID, ERPPN, WorkDept Order by ModelID"
        mSQLReader = mSQLS1.ExecuteReader
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                Ws.Cells(LineZ, 1) = mSQLReader.Item("ModelID")
                Ws.Cells(LineZ, 2) = mSQLReader.Item("ERPPN")
                Ws.Cells(LineZ, 3) = mSQLReader.Item("WorkDept")
                Ws.Cells(LineZ, 4) = mSQLReader.Item("t3")
                Ws.Cells(LineZ, 5) = mSQLReader.Item("t2")
                Ws.Cells(LineZ, 6) = "=IFERROR(D" & LineZ & "/E" & LineZ & ",)"
                LineZ += 1
                Label3.Text = "4-" & LineZ
                Label3.Refresh()
            End While
            Ws.Cells(LineZ, 3) = "合计"
            Ws.Cells(LineZ, 4) = "=SUM(D2:D" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 5) = "=SUM(E2:E" & LineZ - 1 & ")"
            Ws.Cells(LineZ, 6) = "=IFERROR(D" & LineZ & "/E" & LineZ & ",)"
        End If
        mSQLReader.Close()

        oRng = Ws.Range("A1", "F1")
        oRng.EntireColumn.AutoFit()


        ' DROP TABLE

        mSQLS1.CommandText = "DROP TABLE ERPSUPPORT.dbo." & TempDB
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            'MsgBox(ex.Message())
            'Return
        End Try

    End Sub
    Private Sub AdjustExcelFormat()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.EntireRow.RowHeight = 16
        Ws.Cells(1, 1) = "品号"
        Ws.Cells(1, 2) = "产品描述"
        Ws.Cells(1, 3) = "生产制令"
        Ws.Cells(1, 4) = "序列号"
        Ws.Cells(1, 5) = "工作站代码"
        oRng = Ws.Range("E1", "E1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 6) = "工作站名称"
        Ws.Cells(1, 7) = "ERP料号"
        Ws.Cells(1, 8) = "生产部门"
        Ws.Cells(1, 9) = "单位标准IE工时"
        oRng = Ws.Range("I1", "I1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        Ws.Cells(1, 10) = "开始时间"
        Ws.Cells(1, 11) = "完成时间"
        Ws.Cells(1, 12) = "工号"
        oRng = Ws.Range("L1", "L1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(1, 13) = "作业员姓名"
        Ws.Cells(1, 14) = "三级部门"
        Ws.Cells(1, 15) = "员工级别"
        Ws.Cells(1, 16) = "Fresh"
        Ws.Cells(1, 17) = "作业人数"
        oRng = Ws.Range("E2", "E2")
        oRng.Select()
        'Ws.FreezePanes(2, 5)
        xExcel.ActiveWindow.FreezePanes = True
    End Sub
    Private Function GetDept(ByVal cf01 As String)
        oCommand.CommandText = "Select nvl(imaud02,'') from ima_file where ima01 = '" & cf01 & "'"
        Dim SS As String = oCommand.ExecuteScalar()
        Return SS
    End Function

    Private Function GetUserDept(ByVal empno As String)
        If empno.StartsWith("0") Then
            empno = Strings.Right(empno, 4)
        End If
        mSQLS2.CommandText = "Select _DeptName3 from T8eHR.dbo.T_EMP_Employee where EmpCode = '" & empno & "'"
        Dim SS As String = mSQLS2.ExecuteScalar()
        Return SS
    End Function

    Private Function GetUserLevel(ByVal empno As String)
        If empno.StartsWith("0") Then
            empno = Strings.Right(empno, 4)
        End If
        mSQLS2.CommandText = "Select  _zjjj from T8eHR.dbo.T_EMP_Employee where EmpCode = '" & empno & "'"
        Dim SS As String = mSQLS2.ExecuteScalar()
        Return SS
    End Function
    Private Function GetGZHour(ByVal empno As String)
        If empno.StartsWith("0") Then
            empno = Strings.Right(empno, 4)
        End If
        mSQLS2.CommandText = "Select isnull(sum(s1._gzss ),0) from T8eHR.dbo.T_ATD_AttDaily  s1 left join T8eHR.dbo.T_EMP_Employee  s2 on s1.EmpID = s2.ID where s1.AttDate between '"
        mSQLS2.CommandText += TimeS3.ToString("yyyy/MM/dd") & "' and '" & TimeS4.ToString("yyyy/MM/dd") & "' and s2.EmpCode = '" & empno & "'"
        Dim SS As String = mSQLS2.ExecuteScalar()
        Return SS
    End Function
    Private Function GetDeptbyHR(ByVal empno As String)
        If empno.StartsWith("0") Then
            empno = Strings.Right(empno, 4)
        End If
        mSQLS2.CommandText = "Select isnull(_DeptName3,'NA') from T8eHR.dbo.T_EMP_Employee Where EmpCode = '" & empno & "'"
        Dim SS As String = mSQLS2.ExecuteScalar()
        Return SS
    End Function
    Private Sub BackgroundWorker1_RunWorkCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SaveExcel()
    End Sub
    Private Sub SaveExcel()
        SaveFileDialog1.FileName = "实际工时分摊明细表"
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
    Private Sub AdjustExcelFormat1()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.EntireRow.RowHeight = 16
        Ws.Cells(1, 8) = "已扣除除外工时"
        Ws.Cells(1, 9) = "已扣除除外工时"
        Ws.Cells(1, 10) = "总标准IE工时"
        Select Case Sector
            Case 0
                Ws.Cells(1, 11) = "'0151"
                Ws.Cells(1, 12) = "'0150"
                Ws.Cells(1, 13) = "'0160"
                Ws.Cells(1, 14) = "'0170"
            Case 1
                Ws.Cells(1, 11) = "'0370"
                Ws.Cells(1, 12) = "'0360"
                Ws.Cells(1, 13) = "'0520"
                Ws.Cells(1, 14) = "'0335"
                Ws.Cells(1, 15) = "'0350"
                Ws.Cells(1, 16) = "'0340"
                Ws.Cells(1, 17) = "'0500"
                Ws.Cells(1, 18) = "'0510"
            Case 2
                Ws.Cells(1, 11) = "'0400"
                Ws.Cells(1, 12) = "'0479"
                Ws.Cells(1, 13) = "'0480"
                Ws.Cells(1, 14) = "'0485"
                Ws.Cells(1, 15) = "'0492"
                Ws.Cells(1, 16) = "'0610"
                Ws.Cells(1, 17) = "'0623"
            Case 3
                Ws.Cells(1, 11) = "'0418"
                Ws.Cells(1, 12) = "'0410"
                Ws.Cells(1, 13) = "'0415"
                Ws.Cells(1, 14) = "'0416"
                Ws.Cells(1, 15) = "'0417"
                Ws.Cells(1, 16) = "'0420"
                Ws.Cells(1, 17) = "'0440"
                Ws.Cells(1, 18) = "'0445"
                Ws.Cells(1, 19) = "'0450"
                Ws.Cells(1, 20) = "'0460"
                Ws.Cells(1, 21) = "'0465"
                Ws.Cells(1, 22) = "'0470"
                Ws.Cells(1, 23) = "'0540"
                Ws.Cells(1, 24) = "'0545"
                Ws.Cells(1, 25) = "'0550"
                Ws.Cells(1, 26) = "'0570"
                Ws.Cells(1, 27) = "'0575"
                Ws.Cells(1, 28) = "'0580"
                Ws.Cells(1, 29) = "'0583"
                Ws.Cells(1, 30) = "'0584"
                Ws.Cells(1, 31) = "'0585"
                Ws.Cells(1, 32) = "'0560"
                Ws.Cells(1, 33) = "'0408"
                Ws.Cells(1, 34) = "'0413"
                Ws.Cells(1, 35) = "'0422"
            Case 4
                Ws.Cells(1, 11) = "'0625"
                Ws.Cells(1, 12) = "'0630"
                Ws.Cells(1, 13) = "'0635"
                Ws.Cells(1, 14) = "'0642"
                Ws.Cells(1, 15) = "'0650"
                Ws.Cells(1, 16) = "'0658"
                Ws.Cells(1, 17) = "'0665"
                Ws.Cells(1, 18) = "'0666"
                Ws.Cells(1, 19) = "'0669"
                Ws.Cells(1, 20) = "'0671"
                Ws.Cells(1, 21) = "'0675"
                Ws.Cells(1, 22) = "'0680"
            Case 5
                Ws.Cells(1, 11) = "'0142"
                Ws.Cells(1, 12) = "'0145"
                Ws.Cells(1, 13) = "'0148"
                Ws.Cells(1, 14) = "'0173"
                Ws.Cells(1, 15) = "'0177"
                Ws.Cells(1, 16) = "'0175"
                Ws.Cells(1, 17) = "'0190"
                Ws.Cells(1, 18) = "'0195"
                Ws.Cells(1, 19) = "'0200"
                Ws.Cells(1, 20) = "'0215"
                Ws.Cells(1, 21) = "'0223"
                Ws.Cells(1, 22) = "'0225"
                Ws.Cells(1, 23) = "'0230"
                Ws.Cells(1, 24) = "'0231"
                Ws.Cells(1, 25) = "'0240"
                Ws.Cells(1, 26) = "'0250"
                Ws.Cells(1, 27) = "'0255"
                Ws.Cells(1, 28) = "'0260"
                Ws.Cells(1, 29) = "'0280"
                Ws.Cells(1, 30) = "'0300"
                Ws.Cells(1, 31) = "'0315"
                Ws.Cells(1, 32) = "'0320"
                Ws.Cells(1, 33) = "'0321"
                Ws.Cells(1, 34) = "'0325"
                Ws.Cells(1, 35) = "'0326"
                Ws.Cells(1, 36) = "'0333"
                Ws.Cells(1, 37) = "'0390"
        End Select
        Ws.Cells(2, 1) = "工号"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        Ws.Cells(2, 2) = "作业员姓名"
        Ws.Cells(2, 3) = "品号"
        Ws.Cells(2, 4) = "ERP料号"
        Ws.Cells(2, 5) = "生产部门"
        Ws.Cells(2, 6) = "订单批号"
        Ws.Cells(2, 7) = "工单编号"
        Ws.Cells(2, 8) = "总出勤实际工时(分钟）"
        Ws.Cells(2, 9) = "分摊后总出勤实际工时(分钟）"
        Ws.Cells(2, 10) = "标准IE工时合计"
        Select Case Sector
            Case 0
                Ws.Cells(2, 11) = "PCM-Layup(预型)"
                Ws.Cells(2, 12) = "Layup----(预型)"
                Ws.Cells(2, 13) = "Layup_Laser(镭射定位叠层）"
                Ws.Cells(2, 14) = "Layup----预型（二次）"
            Case 1
                Ws.Cells(2, 11) = "手工切割TRIMMING"
                Ws.Cells(2, 12) = "CNC_3AXIS"
                Ws.Cells(2, 13) = "CNC_3AXIS"
                Ws.Cells(2, 14) = "CNC_4/3"
                Ws.Cells(2, 15) = "CNC_4AXIS"
                Ws.Cells(2, 16) = "CNC_5AXIS"
                Ws.Cells(2, 17) = "CNC_5AXIS"
                Ws.Cells(2, 18) = "CNC_4AXIS"
            Case 2
                Ws.Cells(2, 11) = "Sand Blast(噴砂)"
                Ws.Cells(2, 12) = "Accessory Warehouse"
                Ws.Cells(2, 13) = "Glue / Bond_I (胶合Ⅰ)"
                Ws.Cells(2, 14) = "Gluing Assembling"
                Ws.Cells(2, 15) = "Post Curing after gluing胶合后硬化"
                Ws.Cells(2, 16) = "Glue / Bond_Ⅱ( 胶合2)"
                Ws.Cells(2, 17) = "Glue / Bond_Ⅲ (胶合Ⅲ)"
            Case 3
                Ws.Cells(2, 11) = "Weigh Before Painting(涂装前称重）"
                Ws.Cells(2, 12) = "Sanding&Filling_Ⅰ(研磨&补土Ⅰ)"
                Ws.Cells(2, 13) = "Replenish(点补)"
                Ws.Cells(2, 14) = "Sanding(磨土)"
                Ws.Cells(2, 15) = "Cleaning Ⅰ（清洁Ⅰ）"
                Ws.Cells(2, 16) = "Painting_Ⅰ(喷涂Ⅰ)"
                Ws.Cells(2, 17) = "Sanding&Filling_Ⅱ(研磨&补土2)"
                Ws.Cells(2, 18) = "Cleaning Ⅱ（清洁Ⅱ）"
                Ws.Cells(2, 19) = "Painting_Ⅱ(喷涂2)"
                Ws.Cells(2, 20) = "Sanding_Ⅲ ( 研磨 3)"
                Ws.Cells(2, 21) = "Cleaning Ⅲ(清洁Ⅲ）"
                Ws.Cells(2, 22) = "Painting_Ⅲ (噴塗3)"
                Ws.Cells(2, 23) = "Sanding_Ⅳ (研磨4)"
                Ws.Cells(2, 24) = "Cleaning Ⅳ(清洁 4）"
                Ws.Cells(2, 25) = "Painting_IV (喷涂4)"
                Ws.Cells(2, 26) = "Sanding_Ⅴ (研磨5)"
                Ws.Cells(2, 27) = "Cleaning Ⅴ(清洁5）"
                Ws.Cells(2, 28) = "Painting_Ⅴ (喷涂5)"
                Ws.Cells(2, 29) = "Sanding_ⅤI (研磨6)"
                Ws.Cells(2, 30) = "Cleaning Ⅵ(清洁6）"
                Ws.Cells(2, 31) = "Painting_ⅤI (喷涂6)"
                Ws.Cells(2, 32) = "Apply Decal (貼水標)"
                Ws.Cells(2, 33) = "Painting proof I(防漆I)"
                Ws.Cells(2, 34) = "Filling Cleaning(补土清洁)"
                Ws.Cells(2, 35) = "Painting proof II(防漆II)"
            Case 4
                Ws.Cells(2, 11) = "Polishing Let Stand"
                Ws.Cells(2, 12) = "Polish(拋光)"
                Ws.Cells(2, 13) = "Polish-2(抛光2）"
                Ws.Cells(2, 14) = "Packing Cleaning (包装清洁)"
                Ws.Cells(2, 15) = "Assembling (装配)"
                Ws.Cells(2, 16) = "组装前整理"
                Ws.Cells(2, 17) = "Let stand(静置)"
                Ws.Cells(2, 18) = "Xray"
                Ws.Cells(2, 19) = "Repair"
                Ws.Cells(2, 20) = "Document upload (文件上传)"
                Ws.Cells(2, 21) = "PizzaPack (包裝)"
                Ws.Cells(2, 22) = "Outer Pack(包裝-裝箱)"
            Case 5
                Ws.Cells(2, 11) = "Latex Core Production"
                Ws.Cells(2, 12) = "Latex Core Inspection"
                Ws.Cells(2, 13) = "Latex Core Warehouse"
                Ws.Cells(2, 14) = "Dissolve Epscore(溶芯材)"
                Ws.Cells(2, 15) = "Preforming Assembling(预型组装)"
                Ws.Cells(2, 16) = "Preheating  (预热)"
                Ws.Cells(2, 17) = "Autoclave (真空压力釜)"
                Ws.Cells(2, 18) = "Autoclave (二次真空压力釜)"
                Ws.Cells(2, 19) = "Fast_S----成型_S"
                Ws.Cells(2, 20) = "Fast Autoclave_B/M（大/中）"
                Ws.Cells(2, 21) = "Fast Autoclave_B/M/S/OVEN（大/中/小）/成型烤箱"
                Ws.Cells(2, 22) = "Fast Autoclave_B/M/S（大/中/小）"
                Ws.Cells(2, 23) = "Cooling I----冷却 I"
                Ws.Cells(2, 24) = "PCM_Cooling I(冷却)"
                Ws.Cells(2, 25) = "Hotpress_250t"
                Ws.Cells(2, 26) = "Cooling_50t"
                Ws.Cells(2, 27) = "315T Forming315T成型"
                Ws.Cells(2, 28) = "Press_500T"
                Ws.Cells(2, 29) = "Press 1000t"
                Ws.Cells(2, 30) = "Cooling_50t"
                Ws.Cells(2, 31) = "Cooling II----冷却II"
                Ws.Cells(2, 32) = "De-Mold----脱模"
                Ws.Cells(2, 33) = "PCM-De-Mold(脱模)"
                Ws.Cells(2, 34) = "Edge Sanding----打毛边"
                Ws.Cells(2, 35) = "PCM Edge Sanding"
                Ws.Cells(2, 36) = "Take_Out_The_Tube抽取风管"
                Ws.Cells(2, 37) = "Post Curing(後加溫(固化))"
        End Select
        
        oRng = Ws.Range("H1", "N1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("H3", "H3")
        oRng.Select()
        'Ws.FreezePanes(2, 5)
        xExcel.ActiveWindow.FreezePanes = True
    End Sub
    Private Sub AdjustExcelFormat2()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.EntireRow.RowHeight = 16
        Ws.Cells(1, 1) = "工号"
        Ws.Cells(1, 2) = "作业员姓名"
        Ws.Cells(1, 3) = "部门"
        Ws.Cells(1, 4) = "标准IE工时"
        Ws.Cells(1, 5) = "实际投入工时"
        Ws.Cells(1, 6) = "效率"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("D1", "E1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("C2", "C2")
        oRng.Select()
        'Ws.FreezePanes(2, 5)
        xExcel.ActiveWindow.FreezePanes = True
    End Sub
    Private Sub AdjustExcelFormat3()
        xExcel.ActiveWindow.Zoom = 75
        Ws.Rows.EntireRow.RowHeight = 16
        Ws.Cells(1, 1) = "品号"
        Ws.Cells(1, 2) = "ERP料号"
        Ws.Cells(1, 3) = "生产部门"
        Ws.Cells(1, 4) = "标准IE工时"
        Ws.Cells(1, 5) = "实际投入工时"
        Ws.Cells(1, 6) = "效率"
        oRng = Ws.Range("A1", "A1")
        oRng.EntireColumn.NumberFormatLocal = "@"
        oRng = Ws.Range("D1", "E1")
        oRng.EntireColumn.NumberFormatLocal = "#,##0.00_ "
        oRng = Ws.Range("F1", "F1")
        oRng.EntireColumn.NumberFormatLocal = "0.00%"
        oRng = Ws.Range("C2", "C2")
        oRng.Select()
        'Ws.FreezePanes(2, 5)
        xExcel.ActiveWindow.FreezePanes = True
    End Sub
    Private Function Calculate1(ByVal empno As String)
        mSQLS2.CommandText = "Select isnull(sum(case when count1 = 0 then Round(IETime,2) else round(ietime/count1, 2) end),0) as t1 "
        mSQLS2.CommandText += "from ( "
        mSQLS2.CommandText += "Select lot.model,model.modelname ,lot.lot, tracking.sn, tracking.station, station.stationname_cn , cf01,timein, timeout, tracking.users , s2.UserID ,"
        mSQLS2.CommandText += "(Select count(*) from MultipleUserRecord s3 where tracking.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) as IETime , s5.name , fresh "
        mSQLS2.CommandText += "from tracking left join lot on tracking.lot = lot.lot left join model on lot.model = model.model left join station on tracking.station = station.station "
        mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and tracking.station = model_station_paravalue.station "
        mSQLS2.CommandText += "left join MultipleUserRecord s2 on tracking.Id = s2.TrackingID left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and tracking.station = s4.StationCode "
        mSQLS2.CommandText += "left join users s5 on (case when s2.UserID is null then tracking.users else s2.UserID  end ) = s5.id where tracking.station in ("
        Select Case Sector
            Case 0
                mSQLS2.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS2.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS2.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS2.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS2.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS2.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS2.CommandText += ") and timeout between '"

        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "Select lot.model,model.modelname ,lot.lot, tracking_dup.sn, tracking_dup.station, station.stationname_cn , cf01,timein, timeout, tracking_dup.users , s2.UserID ,"
        mSQLS2.CommandText += "(Select count(*) from MultipleUserRecord s3 where tracking_dup.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) , s5.name , fresh "
        mSQLS2.CommandText += "from tracking_dup left join lot on tracking_dup.lot = lot.lot left join model on lot.model = model.model left join station on tracking_dup.station = station.station "
        mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and tracking_dup.station = model_station_paravalue.station "
        mSQLS2.CommandText += "left join MultipleUserRecord s2 on tracking_dup.Id = s2.TrackingID  left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and tracking_dup.station = s4.StationCode "
        mSQLS2.CommandText += "left join users s5 on (case when s2.UserID is null then tracking_dup.users else s2.UserID  end ) = s5.id where tracking_dup.station in ("
        Select Case Sector
            Case 0
                mSQLS2.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS2.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS2.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS2.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS2.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS2.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS2.CommandText += ") and timeout between '"

        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' "
        mSQLS2.CommandText += "union all "
        mSQLS2.CommandText += "Select lot.model,model.modelname ,lot.lot, scrap_tracking.sn, scrap_tracking.station, station.stationname_cn , cf01,timein, timeout, scrap_tracking.users , s2.UserID ,"
        mSQLS2.CommandText += "(Select count(*) from MultipleUserRecord s3 where scrap_tracking.id = s3.TrackingID ) as count1, isnull(s4.IETime,0) , s5.name , fresh "
        mSQLS2.CommandText += "from scrap_tracking left join lot on scrap_tracking.lot = lot.lot left join model on lot.model = model.model left join station on scrap_tracking.station = station.station "
        mSQLS2.CommandText += "left join model_station_paravalue on lot.model = model_station_paravalue.model and model_station_paravalue.profilename = 'ERP' and scrap_tracking.station = model_station_paravalue.station "
        mSQLS2.CommandText += "left join MultipleUserRecord s2 on scrap_tracking.Id = s2.TrackingID left join ERPSUPPORT.dbo.FIIETIME s4 on lot.model = s4.ModelID and scrap_tracking.station = s4.StationCode "
        mSQLS2.CommandText += "left join users s5 on (case when s2.UserID is null then scrap_tracking.users else s2.UserID  end ) = s5.id where scrap_tracking.station in ("
        Select Case Sector
            Case 0
                mSQLS2.CommandText += "'0150','0151','0160','0170'"
            Case 1
                mSQLS2.CommandText += "'0360','0370','0335','0350','0340','0500','0510','0520'"
            Case 2
                mSQLS2.CommandText += "'0400','0479','0480','0485','0492','0610','0623'"
            Case 3
                mSQLS2.CommandText += "'0418','0410','0415','0416','0417','0420','0440','0445', '0450', '0460','0465','0470','0540','0545','0550','0570','0575','0580','0583','0584','0585','0560','0408','0413','0422'"
            Case 4
                mSQLS2.CommandText += "'0625','0630','0635','0642','0650','0658','0665','0666', '0669', '0671','0675','0680'"
            Case 5
                mSQLS2.CommandText += "'0142','0145','0148','0173','0177','0175','0190','0195', '0200', '0215','0223','0225','0230','0231','0240','0250','0255','0260','0280','0300','0315','0320','0321','0325','0326','0333','0390'"
        End Select
        mSQLS2.CommandText += ") and timeout between '"
        mSQLS2.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS2.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' ) as XX Where (case when UserID is null then users else UserID end) = '" & empno & "'"
        Dim SS As String = mSQLS2.ExecuteScalar()
        Return SS

    End Function
End Class