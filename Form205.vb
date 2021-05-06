Public Class Form205
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim DS As Data.DataSet = New DataSet()
    Dim RowsCount As Integer = 0
    Dim RowsCount1 As Integer = 0
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [明细$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            'Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
        RowsCount = 0
        RowsCount = DS.Tables("table1").Rows.Count
        Label6.Text = RowsCount
        If RowsCount > 0 Then
            Button2.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RowsCount1 = 0
        Dim g_success As Boolean = True
        Dim Tran1 As SqlClient.SqlTransaction = mConnection.BeginTransaction()
        mSQLS1.Transaction = Tran1
        For i As Integer = 0 To RowsCount - 1 Step 1
            ' 檢查是否第1行和第5行是空白
            If IsDBNull(DS.Tables("table1").Rows(i).Item(0)) Then
                g_success = False
                Tran1.Rollback()
                Exit For
            End If
            If IsDBNull(DS.Tables("table1").Rows(i).Item(4)) Then
                g_success = False
                Tran1.Rollback()
                Exit For
            End If
            mSQLS1.CommandText = "DELETE FIIETIME Where ModelID = '" & DS.Tables("table1").Rows(i).Item(0).ToString() & "' AND StationCode = '" & DS.Tables("table1").Rows(i).Item(4) & "'"
            Try
                mSQLS1.ExecuteNonQuery()
            Catch ex As Exception
                'MsgBox(ex.Message())
            End Try

            mSQLS1.CommandText = "INSERT INTO FIIETIME VALUES ('" & DS.Tables("table1").Rows(i).Item(0).ToString() & "','" & DS.Tables("table1").Rows(i).Item(1).ToString() & "','" & DS.Tables("table1").Rows(i).Item(3).ToString()
            mSQLS1.CommandText += "'," & DS.Tables("table1").Rows(i).Item(6).ToString() & ",'" & DS.Tables("table1").Rows(i).Item(4).ToString() & "','" & DS.Tables("table1").Rows(i).Item(5).ToString() & "')"
            Try
                mSQLS1.ExecuteNonQuery()
                RowsCount1 += 1
                Label8.Text = RowsCount1
                Label8.Refresh()
            Catch ex As Exception
                MsgBox(ex.Message())
                g_success = False
                Tran1.Rollback()
                Exit For
            End Try
        Next
        If g_success = True Then
            Tran1.Commit()
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Form205_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT()
        
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
End Class