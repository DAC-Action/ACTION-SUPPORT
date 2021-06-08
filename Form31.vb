Public Class Form31
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLS2 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader

    Dim Ds As New DataSet()
    Dim Sda As New SqlClient.SqlDataAdapter
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim oReader As Oracle.ManagedDataAccess.Client.OracleDataReader


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '先檢查有沒有足夠的資料
        If String.IsNullOrEmpty(TextBox3.Text) Or IsDBNull(ComboBox1.SelectedItem) Or IsDBNull(ComboBox2.SelectedItem) Or IsDBNull(TextBox1.Text) Then
            MsgBox("请输入所有的栏位")
            Return
        End If
        ' 檢查有無此專案號

        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        'oCommand.CommandText = "SELECT COUNT(*) FROM pja_file WHERE pja01 = '"
        'oCommand.CommandText += TextBox1.Text & "' AND pjaconf = 'Y'"
        'Dim HC As Decimal = oCommand.ExecuteScalar()
        'If HC <= 0 Then
        'MsgBox("无此专案编号,请确认")
        'Return
        'End If
        ' 檢查有無此人員 (暫時不做)
        ' 檢查是否為0為負的小時數
        If TextBox3.Text <= 0 Then
            MsgBox("时数为0或负数")
            Return
        End If
        If Not IsNumeric(TextBox1.Text) Then
            MsgBox("APQP请输入数字")
            Return
        End If
        If Strings.Len(Me.TextBox4.Text) > 60 Then
            MsgBox("工作内容字数超过60")
            Return
        End If
        If Strings.Len(Me.TextBox5.Text) > 60 Then
            MsgBox("备注字数超过60")
            Return
        End If
        ' 檢查完畢, 插入系統
        Dim CA As String() = Strings.Split(ComboBox1.SelectedItem.ToString(), ",", 2, CompareMethod.Text)
        Dim CB As String() = Strings.Split(ComboBox2.SelectedItem.ToString(), "|", 2, CompareMethod.Text)

        'mSQLS1.CommandText = "INSERT INTO ProjectHR (EProject,EDate,EUser,EHour,eUserName) VALUES ('" & TextBox1.Text & "','"
        mSQLS1.CommandText = "INSERT INTO ProjectHR (EProject,EDate,EUser,EHour,eUserName,EAP,ModelID,WorkDesc,Remark) VALUES ('" & CB(1) & "','"
        'mSQLS1.CommandText += DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','" & ComboBox1.SelectedItem & "'," & TextBox3.Text & ",'"
        mSQLS1.CommandText += DateTimePicker1.Value.ToString("yyyy/MM/dd") & "','" & CA(0) & "'," & TextBox3.Text & ",'"
        'mSQLS1.CommandText += TextBox2.Text & "')"
        mSQLS1.CommandText += CA(1) & "'," & TextBox1.Text & ",'" & TextBox2.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "')"
        Try
            mSQLS1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message())
            Return
        End Try
        ReloadGrid()
    End Sub

    Private Sub Form31_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
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
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "select gen01,gen02 from gen_file where genacti = 'Y' and gen03 = 'D2300' order by gen01"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                ComboBox1.Items.Add(oReader.Item("gen01") & "," & oReader.Item("gen02"))
            End While
        End If
        oReader.Close()
        oCommand.CommandText = "select pja01,pja02 FROM pja_file where pjaacti = 'Y' and pjaclose = 'N' and pjaconf = 'Y' order by pja02"
        oReader = oCommand.ExecuteReader()
        If oReader.HasRows() Then
            While oReader.Read()
                ComboBox2.Items.Add(oReader.Item("pja02") & "|" & oReader.Item("pja01"))
            End While
        End If
        ReloadGrid()
    End Sub
    Private Sub ReloadGrid()
        Ds.Clear()
        Sda = New SqlClient.SqlDataAdapter("select * from ProjectHR order by RecordTime Desc", mConnection)
        Sda.Fill(Ds)
        Me.DataGridView1.DataSource = Ds.Tables(0)
        Me.DataGridView1.Columns(0).Width = 100
        Me.DataGridView1.Columns(0).HeaderText = "专案编号"
        Me.DataGridView1.Columns(1).Width = 100
        Me.DataGridView1.Columns(1).HeaderText = "日期"
        Me.DataGridView1.Columns(2).Width = 100
        Me.DataGridView1.Columns(2).HeaderText = "人员编号"
        Me.DataGridView1.Columns(3).Width = 100
        Me.DataGridView1.Columns(3).HeaderText = "时数"
        Me.DataGridView1.Columns(4).Width = 100
        Me.DataGridView1.Columns(4).HeaderText = "人员名称"
        Me.DataGridView1.Columns(5).Width = 100
        Me.DataGridView1.Columns(5).HeaderText = "登记时间"
        Me.DataGridView1.Columns(6).Width = 80
        Me.DataGridView1.Columns(6).HeaderText = "APQP"
        Me.DataGridView1.Columns(7).Width = 100
        Me.DataGridView1.Columns(7).HeaderText = "型号"
        Me.DataGridView1.Columns(8).Width = 100
        Me.DataGridView1.Columns(8).HeaderText = "工作内容"
        Me.DataGridView1.Columns(9).Width = 100
        Me.DataGridView1.Columns(9).HeaderText = "备注"
        Me.DataGridView1.Columns(10).Width = 100
        Me.DataGridView1.Columns(10).HeaderText = "部门编号"
        'MsgBox(Me.DataGridView1.Columns(1).Width)

        '20170313
        Me.DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke
        Me.DataGridView1.Enabled = True
        Me.DataGridView1.Show()
    End Sub

    'Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
    '    If oConnection.State <> ConnectionState.Open Then
    '        Try
    '            oConnection.Open()
    '            oCommand.Connection = oConnection
    '            oCommand.CommandType = CommandType.Text
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    End If
    '    oCommand.CommandText = "SELECT pja02 FROM pja_file WHERE pjaconf = 'Y' and pja01 = '" & TextBox1.Text & "'"
    '    Try
    '        TextBox4.Text = oCommand.ExecuteScalar()
    '    Catch ex As Exception
    '        MsgBox(ex.Message())
    '    End Try
    'End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'If oConnection.State <> ConnectionState.Open Then
        '    Try
        '        oConnection.Open()
        '        oCommand.Connection = oConnection
        '        oCommand.CommandType = CommandType.Text
        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    End Try
        'End If
        'oCommand.CommandText = "SELECT gen02 FROM gen_file where gen01 = '" & Me.ComboBox1.SelectedItem & "'"
        'Try
        '    TextBox2.Text = oCommand.ExecuteScalar()
        'Catch ex As Exception
        '    MsgBox(ex.Message())
        'End Try
        'Dim CA As String() = Strings.Split(ComboBox1.SelectedItem.ToString(), ",", 2, CompareMethod.Text)
        'ComboBox1.SelectedText = CA(0)
        'MsgBox(CA(0))
        'MsgBox(ComboBox1.SelectedText)
        'MsgBox(ComboBox1.SelectedValue)

    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Dim i As Integer = ComboBox2.FindString(ComboBox2.Text)
        ComboBox2.Select()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.DataGridView1.SelectedCells.Count > 1 Or Me.DataGridView1.SelectedCells.Count <= 0 Then
            MsgBox("请选定修改格")
            Return
        End If
        If Me.DataGridView1.SelectedCells.Item(0).ColumnIndex = 1 Or Me.DataGridView1.SelectedCells.Item(0).ColumnIndex = 3 Then
            'Me.DataGridView1.
            Me.DataGridView1.ReadOnly = False
            Me.DataGridView1.BeginEdit(False)
        Else
            MsgBox("此格无法修改")
            Return
        End If
    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        If Me.DataGridView1.ReadOnly = False Then
            Me.DataGridView1.ReadOnly = True
            Me.DataGridView1.EndEdit()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Ds.HasChanges() Then
            Dim cb As New SqlClient.SqlCommandBuilder(Sda)
            Dim ca As Integer = Sda.Update(Ds.Tables(0))
            Ds.Tables(0).AcceptChanges()
            Me.DataGridView1.Update()
            MsgBox("已更新")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [报表$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try

            mSQLS1.CommandText = "Truncate table ProjectHR_PM_Only"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Dim Tran1 As SqlClient.SqlTransaction = mConnection.BeginTransaction()
            mSQLS1.Transaction = Tran1
            For i As Int16 = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
                If IsDBNull(DS.Tables("table1").Rows(i).Item(0)) Or DS.Tables("table1").Rows(i).Item(0).ToString() = "#N/A" Then
                    Exit For
                End If
                If DS.Tables("table1").Rows(i).Item(0).ToString() = "实际周报表" Then
                    Exit For
                End If

                    'oCommand.CommandText = "select count(pja01) from pja_file where (pjaacti = 'N' or pjaclose = 'Y') and pja01 = '" & DS.Tables("table1").Rows(i).Item(1).ToString() & "'"    '200828 mark by Brady
                oCommand.CommandText = "select count(pja01) from pja_file where (pjaacti = 'N' or (pjaclose = 'Y' and (to_char(pjaud14,'yyyy/mm/dd') < '" & Convert.ToDateTime(DS.Tables("table1").Rows(i).Item(4).ToString()).ToString("yyyy/MM/dd") & "'))) and pja01 = '" & DS.Tables("table1").Rows(i).Item(1).ToString() & "'"    '200828 add by Brady
                Dim FailCount As Int16 = oCommand.ExecuteScalar()
                If FailCount > 0 Then
                    Tran1.Rollback()
                    MsgBox("Failed to Upload:" & DS.Tables("table1").Rows(i).Item(1).ToString())
                    Exit For
                End If
                If DS.Tables("table1").Rows(i).Item(2).ToString() = "PL_RESP" Then
                    mSQLS1.CommandText = "INSERT INTO ProjectHR_PM_Only (EProject,EDate,EUser,EHour,eUserName,EAP,ModelID,WorkDesc,Remark, EDepartNo) VALUES ('" & DS.Tables("table1").Rows(i).Item(1).ToString() & "','"
                    mSQLS1.CommandText += Convert.ToDateTime(DS.Tables("table1").Rows(i).Item(4).ToString()).ToString("yyyy/MM/dd") & "','" & DS.Tables("table1").Rows(i).Item(3).ToString() & "'," & DS.Tables("table1").Rows(i).Item(8).ToString() & ",'"
                    mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(0).ToString() & "','" & DS.Tables("table1").Rows(i).Item(6).ToString() & "','" & DS.Tables("table1").Rows(i).Item(5).ToString() & "','"
                    mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(7).ToString() & "','" & DS.Tables("table1").Rows(i).Item(9).ToString() & "','" & DS.Tables("table1").Rows(i).Item(2).ToString() & "')"
                Else
                    mSQLS1.CommandText = "INSERT INTO ProjectHR (EProject,EDate,EUser,EHour,eUserName,EAP,ModelID,WorkDesc,Remark, EDepartNo) VALUES ('" & DS.Tables("table1").Rows(i).Item(1).ToString() & "','"
                    mSQLS1.CommandText += Convert.ToDateTime(DS.Tables("table1").Rows(i).Item(4).ToString()).ToString("yyyy/MM/dd") & "','" & DS.Tables("table1").Rows(i).Item(3).ToString() & "'," & DS.Tables("table1").Rows(i).Item(8).ToString() & ",'"
                    mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(0).ToString() & "','" & DS.Tables("table1").Rows(i).Item(6).ToString() & "','" & DS.Tables("table1").Rows(i).Item(5).ToString() & "','"
                    mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(7).ToString() & "','" & DS.Tables("table1").Rows(i).Item(9).ToString() & "','" & DS.Tables("table1").Rows(i).Item(2).ToString() & "')"
                End If
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    Tran1.Rollback()
                    MsgBox(ex.Message())
                    Exit For
                End Try

            Next

            If Not IsDBNull(Tran1.Connection) Then
                Tran1.Commit()
            End If
            Tran1.Dispose()

            ' add by cloud 2021060
            mSQLS2.CommandText = "Select CONVERT(varchar(100), EDate, 25) as Edate,euser, sum(ehour) as t1 from ProjectHR_PM_Only group by edate,euser order by edate,euser"
            mSQLReader = mSQLS2.ExecuteReader
            If mSQLReader.HasRows() Then
                While mSQLReader.Read()
                    Dim SS As Decimal = mSQLReader.Item("t1")
                    If SS <= 8 Then
                        mSQLS1.CommandText = "Insert into ProjectHR Select EProject , EDate, Euser, EHour , EUserName , RecordTime , EAP, ModelID , WorkDesc , Remark , EDepartNo  from ProjectHR_PM_ONLY where edate = '" & mSQLReader.Item("eDate").ToString() & "' and euser = '" & mSQLReader.Item("eUser").ToString() & "'"
                    Else
                        mSQLS1.CommandText = "Insert into ProjectHR Select EProject , EDate, Euser, Round(EHour /" & SS & " * 8,2) as Ehour, EUserName , RecordTime , EAP, ModelID , WorkDesc , Remark , EDepartNo  from ProjectHR_PM_ONLY where edate = '" & mSQLReader.Item("eDate").ToString() & "' and euser = '" & mSQLReader.Item("eUser").ToString() & "'"
                    End If
                    Try
                        mSQLS1.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.Message())
                    End Try
                End While
            End If
            mSQLReader.Close()

            
            ReloadGrid()
        End If
    End Sub
End Class