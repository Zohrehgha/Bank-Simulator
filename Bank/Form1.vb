Public Class Form1
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'setting up the ADO objects
        conn.Provider = "Microsoft.Jet.OleDB.4.0"
        ' Setting up the jet DB driver (access) 
        conn.ConnectionString = "C:\ITD\Term 3\Visual Basic.Net\assignment\assignment 2\Database2.mdb"
        conn.Open()
        rs.Open("select * from AccountHolders", conn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic)
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub
End Class
