Public Class Form189
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oConnection2 As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oConnection9 As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oSQLS1 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oSQLReader As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oCommander As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander2 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oCommander9 As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader99 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader98 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim oReader97 As Oracle.ManagedDataAccess.Client.OracleDataReader
    Dim StopSign As Boolean = False
    Dim TempHeader As String = String.Empty
    Dim ptime As String = String.Empty
    Dim r_percentage As Decimal = 0
    Dim TimeS1 As DateTime
    Dim TimeS2 As DateTime
    Dim LineZ As Integer = 0
    Dim PaperDate As Date
    Dim OldExcel() As Process = Process.GetProcessesByName("Excel")
    Private Sub Form189_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfMes()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        oConnection2.ConnectionString = Module1.OpenOracleConnection("actiontest")
        oConnection9.ConnectionString = Module1.OpenOracleConnection("actiontest")
        PaperDate = Now.AddDays(-1)
        ptime = Today.AddDays(-1).ToString("yyyy/MM/dd")
        ptime = ptime & " 08:00:00"
        Me.DateTimePicker1.Value = Convert.ToDateTime(ptime)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oSQLS1.Connection = oConnection
                oSQLS1.CommandType = CommandType.Text
                oCommander.Connection = oConnection
                oCommander.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        If oConnection2.State <> ConnectionState.Open Then
            Try
                oConnection2.Open()                
                oCommander2.Connection = oConnection2
                oCommander2.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

        If oConnection9.State <> ConnectionState.Open Then
            oCommander9.Connection = oConnection9
            oCommander9.CommandType = CommandType.Text
            oConnection9.Open()
        End If

        TimeS1 = DateTimePicker1.Value
        TimeS2 = TimeS1.AddDays(1)
        mSQLS1.CommandText = "Select scrap.lot, cf01, count(scrap.sn) as t1  from scrap left join scrap_sn on scrap.sn = scrap_sn.sn left join lot on scrap.lot = lot.lot "
        mSQLS1.CommandText += "left join model_station_paravalue as b on b.profilename = 'ERP' and b.model = lot.model and b.station = scrap_sn.updatedstation where scrap.datetime between '"
        mSQLS1.CommandText += TimeS1.ToString("yyyy/MM/dd HH:mm:ss") & "' AND '"
        mSQLS1.CommandText += TimeS2.ToString("yyyy/MM/dd HH:mm:ss") & "' and scrap.defect not in ('DJ01','DJ02') group by scrap.lot,cf01"
        mSQLReader = mSQLS1.ExecuteReader()
        If mSQLReader.HasRows() Then
            While mSQLReader.Read()
                FindTempHeader(mSQLReader.Item("cf01"))
                StopSign = False
                WorkOrder(TempHeader, mSQLReader.Item("t1"), mSQLReader.Item("cf01"), mSQLReader.Item("lot"))
                ExtendWorkOrder(mSQLReader.Item("cf01"), mSQLReader.Item("t1"), mSQLReader.Item("lot"))
                If oConnection2.State <> ConnectionState.Closed Then
                    oConnection2.Close()
                End If
                'WorkOrder()
            End While
        End If
        mSQLReader.Close()
        MsgBox("Done")
    End Sub
    Private Sub WorkOrder(ByVal header As String, sfb08 As Decimal, ByVal sfb05 As String, ByVal sfbud02 As String)

        Dim sfb01 As String = String.Empty
        'Dim sfb82 As String = "D35"
        Dim sfb82 As String = String.Empty
        oCommander.CommandText = "SELECT nvl(imaud02,'NA') FROM IMA_file where ima01 = '" & sfb05 & "'"
        sfb82 = oCommander.ExecuteScalar()
        sfb01 = Getsfb01(header)
        'If Strings.Right(sfb05, 1) = "A" Then
        'sfb82 = sfb82 & Strings.Right(sfb05, 3)
        'sfb82 = sfb82.Remove(sfb82.Count() - 1)
        'Else
        'sfb82 = sfb82 & Strings.Right(sfb05, 2)
        'End If

        oCommander.CommandText = "INSERT INTO sfb_file VALUES ('" & sfb01 & "',1,NULL,2,'" & sfb05 & "',NULL,NULL,to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        oCommander.CommandText += "," & sfb08 & ",0,0,0,0,0,0,0,NULL,to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd'),'00:00',to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        oCommander.CommandText += ",'00:00',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'Y','N',NULL,to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        oCommander.CommandText += ",NULL,NULL,NULL,NULL,'Y',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,'N',NULL,to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/MM/dd')"
        oCommander.CommandText += ",'" & sfb82 & "',NULL,NULL,'Y',NULL,NULL,NULL,'N','N',' ',0,NULL,NULL,'N',1,NULL,'Y','Automation','D1461',NULL,NULL,NULL,'N',NULL,NULL,NULL,NULL,NULL,NULL,'" & sfbud02 & "',NULL,NULL,NULL,NULL"
        oCommander.CommandText += ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,'Automation','N','ACTIONTEST','ACTIONTEST','Automation','D1461','N',NULL)"
        Try
            Dim ED As Int16 = oCommander.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        AutoGenerateReworkOrderDetail(sfb05, sfb08, sfb01)
        ' 處理旗標
        oCommander.CommandText = "select count(sfa01) from sfa_file where sfa01 = '" & sfb01 & "' and sfa11 <> 'E'"
        Dim ChangeFlag As Decimal = oCommander.ExecuteScalar()
        If ChangeFlag = 0 Then
            oCommander.CommandText = "UPDATE sfb_file SET sfb39 = 2 WHERE sfb01 = '" & sfb01 & "'"
            Try
                oCommander.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If

    End Sub
    Public Function Getsfb01(ByVal header As String)
        Dim AB As String = header & "-" & Now.ToString("yy") & Now.ToString("MM")
        Dim oCommanderNew As New Oracle.ManagedDataAccess.Client.OracleCommand
        If oConnection2.State <> ConnectionState.Open Then
            Try
                oConnection2.Open()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        oCommanderNew.Connection = oConnection2
        oCommanderNew.CommandType = CommandType.Text
        oCommanderNew.CommandText = "select nvl(MAX(SUBSTR(SFB01,11,4)),0) from sfb_file where sfb01 LIKE '" & AB & "%'"

        Dim MaxInt As Integer = oCommanderNew.ExecuteScalar()
        MaxInt += 1
        Select Case Strings.Len(MaxInt.ToString())
            Case 1
                AB = AB & "000" & MaxInt
            Case 2
                AB = AB & "00" & MaxInt
            Case 3
                AB = AB & "0" & MaxInt
            Case 4
                AB = AB & MaxInt
        End Select
        Return AB
    End Function
    Private Sub AutoGenerateReworkOrderDetail(ByVal erp1 As String, ByVal quantity As Decimal, ByVal s1 As String)
        Dim oCommander99 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander99.Connection = oConnection2
        oCommander99.CommandType = CommandType.Text
        Dim oCommander98 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander98.Connection = oConnection2
        oCommander98.CommandType = CommandType.Text
        oCommander98.CommandText = "select Round(sum(bmb06/bmb07),3) as t1,Round(sum(bmb06/bmb07 * (1+ bmb08 /100)),3) as t2,bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16 from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '"
        oCommander98.CommandText += erp1 & "' and bmb29 = ima910 and (bmb05 is NULL or bmb05 > to_date('" & TimeS1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')) AND bmb04 <= to_date('"
        oCommander98.CommandText += TimeS1.ToString("yyyy/MM/dd") & "','yyyy/mm/dd') and ima06 <> '106' group by bmb01,bmb03,ima70,bmb10,ima86,ima64,ima86_fac,bmb16 order by bmb03"
        oReader98 = oCommander98.ExecuteReader()
        If oReader98.HasRows() Then
            While oReader98.Read()
                Dim Usage As Decimal = oReader98.Item("t2") * quantity
                'Dim UnitUsage As Decimal = oReader98.Item("bmb06") / oReader98.Item("bmb07") * (1 + oReader98.Item("bmb08") / 100)
                Dim UnitUsage As Decimal = oReader98.Item("t1")
                Dim UnitUsageR As Decimal = oReader98.Item("t2")
                Dim sfa11 As String = String.Empty
                If oReader98.Item("ima70") = "Y" Then
                    sfa11 = "E"
                Else
                    sfa11 = "N"
                End If
                'If sfa11 = "E" Then
                'Usage = Usage * percentage
                'Usage = Decimal.Round(Usage, 3)
                'UnitUsage = Usage / quantity
                'UnitUsageR = UnitUsage
                'End If
                Dim sfa13 As Decimal = 1
                If oReader98.Item("bmb10").ToString() <> oReader98.Item("ima86").ToString() Then
                    sfa13 = Gsfa13(oReader98.Item("bmb10").ToString(), oReader98.Item("ima86").ToString(), oReader98.Item("bmb03").ToString())
                End If
                If sfa11 = "N" And oReader98.Item("ima64") = 1 Then
                    Usage = Decimal.Ceiling(Usage)
                    UnitUsageR = Usage / quantity
                End If
                oCommander99.CommandText = "INSERT INTO sfa_file VALUES ('" & s1 & "',1,'" & oReader98.Item("bmb03") & "'," & Usage & "," & Usage & ",0,0,0,0,0,0,0,NULL,' ',0,NULL,'" & sfa11 & "','" & oReader98.Item("bmb10") & "'," & sfa13 & ",'" & oReader98.Item("ima86") & "'," & oReader98.Item("ima86_fac") & "," & UnitUsage & "," & UnitUsageR & ",0," & oReader98.Item("bmb16") & ",'"
                oCommander99.CommandText += oReader98.Item("bmb03") & "',1,'" & erp1 & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,'Y','N',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'ACTIONTEST','ACTIONTEST',' ',0)"
                Try
                    oCommander99.ExecuteNonQuery()
                Catch ex As Exception
                    'MsgBox(ex.Message())
                End Try
            End While
        End If
        oReader98.Close()
    End Sub
    Private Function Gsfa13(ByVal v1 As String, ByVal v2 As String, ByVal erppn As String)
        oCommander2.CommandText = "select nvl((smd04/smd06),0) from smd_file where smd01 = '" & erppn & "' and smd03 = '" & v1 & "' and smd02 = '" & v2 & "'"
        Dim sfa13 As Decimal = 1
        Try
            sfa13 = oCommander2.ExecuteScalar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        If IsDBNull(sfa13) Then
            sfa13 = 1
        End If
        Return sfa13
    End Function
    Private Sub ExtendWorkOrder(ByVal erp1 As String, ByVal quantity As Decimal, ByVal sfbud02 As String)
        Dim oCommander97 As New Oracle.ManagedDataAccess.Client.OracleCommand
        oCommander97.Connection = oConnection2
        oCommander97.CommandType = CommandType.Text
        oCommander97.CommandText = "select * from bmb_file full join ima_file on bmb03 = ima01 where bmb01 = '"
        oCommander97.CommandText += erp1 & "' and bmb05 is NULL and bmb29 = ima910 and bmb19 = 2 order by bmb03"
        oReader97 = oCommander97.ExecuteReader()
        If oReader97.HasRows() Then
            While oReader97.Read()
                If IsDBNull(oReader97.Item("ima111")) Then
                    Continue While
                End If
                WorkOrder(oReader97.Item("ima111"), quantity, oReader97.Item("bmb03"), sfbud02)
                'WorkOrder(oReader97.Item("bmb03"), oReader97.Item("bmb01"), quantity)
                If StopSign = False Then
                    Call ExtendWorkOrder(oReader97.Item("bmb03"), quantity, sfbud02)
                End If
            End While
        Else
            StopSign = True
        End If
        'oReader97.Close()

    End Sub
    Private Sub FindTempHeader(ByVal ima01 As String)
        
        oCommander9.CommandText = "select ima111 from ima_file where ima01 = '"
        oCommander9.CommandText += ima01 & "'"
        oReader99 = oCommander9.ExecuteReader
        If oReader99.HasRows() Then
            oReader99.Read()
            TempHeader = oReader99.Item("ima111")
        End If
        oReader99.Close()
    End Sub
End Class