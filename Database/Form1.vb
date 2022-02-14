Public Class Form1
    Private myDS As New DataSet
    Private nRow As Integer

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim myDA As New OleDb.OleDbDataAdapter
        Dim myCommand As New OleDb.OleDbCommand
        Dim myConnection As New OleDb.OleDbConnection

        myConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\delus\Documents\Database1.accdb"
        myCommand.CommandText = "SELECT Employees.DeptCode, Employees.DateHire, Employees.DateBirth, Employees.PhoneNumber, Employees.Zip, Employees.State, Employees.City, Employees.Address2, Employees.Address1, Employees.FirstName,  Employees.LastName, Employees.ID,Employees.Activeind, JobLookUp.[Desc], DeptLookUp.[Desc] AS Expr1 FROM ((DeptLookUp INNER JOIN Employees ON DeptLookUp.ID = Employees.DeptCode) INNER JOIN JobLookUp ON Employees.JobCode = JobLookUp.Code)"
        myCommand.Connection = myConnection
        myDA.SelectCommand = myCommand
        myDA.Fill(myDS)

        nRow = 0
        LoadFormData()
    End Sub

    Private Sub LoadFormData()
        txtEmpID.Text = myDS.Tables(0).Rows(nRow)("ID").ToString
        txtFirstName.Text = myDS.Tables(0).Rows(nRow)("FirstName").ToString
        txtLastName.Text = myDS.Tables(0).Rows(nRow)("LastName").ToString
        txtAddress1.Text = myDS.Tables(0).Rows(nRow)("Address1").ToString
        txtAddress2.Text = myDS.Tables(0).Rows(nRow)("Address2").ToString
        txtCity.Text = myDS.Tables(0).Rows(nRow)("City").ToString
        txtState.Text = myDS.Tables(0).Rows(nRow)("State").ToString
        txtZIP.Text = myDS.Tables(0).Rows(nRow)("Zip").ToString
        txtPhone.Text = myDS.Tables(0).Rows(nRow)("PhoneNumber").ToString
        txtDOB.Text = myDS.Tables(0).Rows(nRow)("DateBirth").ToString
        txtDOH.Text = myDS.Tables(0).Rows(nRow)("DateHire").ToString
        txtDept.Text = myDS.Tables(0).Rows(nRow)("Expr1").ToString
        txtJob.Text = myDS.Tables(0).Rows(nRow)("Desc").ToString
        CheckBox1.Checked = myDS.Tables(0).Rows(nRow)("ActiveInd").ToString
        lblPosition.Text = "Record " & nRow + 1 & " of " & myDS.Tables(0).Rows.Count

    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        nRow = 0
        LoadFormData()
    End Sub

    Private Sub btnPrior_Click(sender As Object, e As EventArgs) Handles btnPrior.Click
        nRow -= 1

        If nRow < 0 Then
            nRow = 0
        End If
        LoadFormData()
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        nRow += 1

        If nRow > myDS.Tables(0).Rows.Count - 1 Then
            nRow = myDS.Tables(0).Rows.Count - 1
        End If
        LoadFormData()
    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        nRow = myDS.Tables(0).Rows.Count - 1
        LoadFormData()

    End Sub
End Class
