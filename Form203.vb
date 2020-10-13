Public Class Form203
    Dim mConnection As New SqlClient.SqlConnection
    Dim mSQLS1 As New SqlClient.SqlCommand
    Dim mSQLReader As SqlClient.SqlDataReader
    Dim Ds As New DataSet()
    Dim Sda As New SqlClient.SqlDataAdapter
    Private Sub Form203_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mConnection.ConnectionString = Module1.OpenConnectionOfERPSUPPORT
        If mConnection.State <> ConnectionState.Open Then
            Try
                mConnection.Open()
                mSQLS1.Connection = mConnection
                mSQLS1.CommandType = CommandType.Text
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "OPEN EXCEL"
        Dim selectOK As DialogResult = OpenFileDialog1.ShowDialog()
        If selectOK = System.Windows.Forms.DialogResult.OK Then
            Dim ExcelPath As String = OpenFileDialog1.FileName
            Dim Source As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " & ExcelPath & ";Extended Properties = 'Excel 12.0;HDR=YES';"
            Dim Excelconn As OleDb.OleDbConnection = New OleDb.OleDbConnection(Source)
            Excelconn.Open()
            Dim ExcelString = "SELECT * FROM [Sheet1$]"
            Dim ExcelAdapater As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(ExcelString, Excelconn)
            Dim DS As Data.DataSet = New DataSet()
            Try
                ExcelAdapater.Fill(DS, "table1")
            Catch ex As Exception
                MsgBox(ex.Message())
            End Try
            Dim Tran1 As SqlClient.SqlTransaction = mConnection.BeginTransaction()
            mSQLS1.Transaction = Tran1
            For i As Int16 = 0 To DS.Tables("table1").Rows.Count - 1 Step 1
                If IsDBNull(DS.Tables("table1").Rows(i).Item(0)) Or IsDBNull(DS.Tables("table1").Rows(i).Item(2)) Then
                    Exit For
                End If

                mSQLS1.CommandText = "DELETE HR_Temp_Att WHERE Date1 = '" & DS.Tables("table1").Rows(i).Item(0) & "' AND WorkerNo = '" & DS.Tables("table1").Rows(i).Item(2).ToString() & "'"
                Try
                    mSQLS1.ExecuteNonQuery()
                Catch ex As Exception
                    Tran1.Rollback()
                    MsgBox(ex.Message())
                    Exit For
                End Try
                mSQLS1.CommandText = "INSERT INTO HR_Temp_Att VALUES ("
                For j As Int16 = 0 To DS.Tables("table1").Columns.Count - 1 Step 1
                    'MsgBox(ExcelDataReader.Item(i).GetType().ToString())
                    If DS.Tables("table1").Rows(i).Item(j).GetType.ToString() = "System.String" Then
                        mSQLS1.CommandText += "'" & DS.Tables("table1").Rows(i).Item(j) & "',"
                    End If
                    If DS.Tables("table1").Rows(i).Item(j).GetType.ToString() = "System.Double" Then
                        mSQLS1.CommandText += DS.Tables("table1").Rows(i).Item(j) & ","
                    End If
                    If DS.Tables("table1").Rows(i).Item(j).GetType.ToString() = "System.DBNull" Then
                        mSQLS1.CommandText += "NULL,"
                    End If
                    If DS.Tables("table1").Rows(i).Item(j).GetType.ToString() = "System.DateTime" Then
                        mSQLS1.CommandText += "'" & DS.Tables("table1").Rows(i).Item(j) & "',"
                    End If
                Next
                mSQLS1.CommandText = mSQLS1.CommandText.Remove(mSQLS1.CommandText.Length - 1)
                mSQLS1.CommandText += ") "

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
            MsgBox("DONE")
        End If
    End Sub
End Class