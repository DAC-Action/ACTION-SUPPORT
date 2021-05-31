Public Class Form116
    Dim oConnection As New Oracle.ManagedDataAccess.Client.OracleConnection
    Dim oCommand As New Oracle.ManagedDataAccess.Client.OracleCommand
    Dim tYear As Decimal = 0
    Dim tWeek As Decimal = 0

    Private Sub Form116_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        oConnection.ConnectionString = Module1.OpenOracleConnection("actiontest")
        tYear = Today.Year
        Me.TextBox1.Text = tYear
        If oConnection.State <> ConnectionState.Open Then
            Try
                oConnection.Open()
                oCommand.Connection = oConnection
                oCommand.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        oCommand.CommandText = "SELECT tc_azn05 FROM tc_azn_file WHERE tc_azn01 = to_date('" & Today.ToString("yyyy/MM/dd") & "','yyyy/mm/dd')"
        tWeek = oCommand.ExecuteScalar()
        Me.TextBox2.Text = tWeek
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT ERPPN,Year,WeekNum,Inbound FROM [sheet1$] Where Year > " & tYear & " or (year = " & tYear & " and WeekNum >=" & tWeek & ")"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Me.DataGridView1.DataSource = DS.Tables("table1")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        tYear = TextBox1.Text
        '刪除所有當週週及之後的資料
        'Me.Label3.Text = "DELETE DATA"
        'For I As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
        '    oCommand.CommandText = "DELETE tc_prk_file WHERE tc_prk01 = '" & DataGridView1.Rows(I).Cells("ERPPN").Value & "'"
        '    Try
        '        oCommand.ExecuteNonQuery()
        '    Catch ex As Exception
        '        'MsgBox(ex.Message())
        '        'Return
        '    End Try
        'Next

        ' 匯入Datagridview 資料
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step 1
            Me.Label3.Text = "DELETE DATA" & i
            Me.Label3.Refresh()
            oCommand.CommandText = "DELETE tc_prk_file WHERE tc_prk01 = '" & DataGridView1.Rows(i).Cells("ERPPN").Value & "' and tc_prk02 = " & DataGridView1.Rows(i).Cells("Year").Value & " AND tc_prk03 = " & DataGridView1.Rows(i).Cells("WeekNum").Value
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception

            End Try

            Me.Label3.Text = "INSERT DATA" & i
            Me.Label3.Refresh()
            oCommand.CommandText = "INSERT INTO tc_prk_file (tc_prk01,tc_prk02,tc_prk03,tc_prk04,tc_prklegal) VALUES ('"
            oCommand.CommandText += DataGridView1.Rows(i).Cells("ERPPN").Value & "'," & DataGridView1.Rows(i).Cells("Year").Value & "," & DataGridView1.Rows(i).Cells("WeekNum").Value & "," & DataGridView1.Rows(i).Cells("Inbound").Value & ",'ACTIONTEST')"
            Try
                oCommand.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message())
                Return
            End Try
        Next
        Me.Label3.Text = "FINISHED"
    End Sub
End Class