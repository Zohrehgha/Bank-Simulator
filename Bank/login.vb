Public Class login
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Private Sub login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'setting up the ADO objects
        conn.Provider = "Microsoft.Jet.OleDB.4.0"
        ' Setting up the jet DB driver (access) 
        conn.ConnectionString = "C:\ITD\Term 3\Visual Basic.Net\assignment\assignment 2\Database2.mdb"
        conn.Open()
        rs.Open("select * from AccountHolders", conn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim clientID As String = ""
        'Dim password As String

        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Plz Fill Client ID and Password")
        Else
            'clientID = TextBox1.Text
            'password = TextBox2.Text
            Dim Criteria As String
            Criteria = "ClientID =" + TextBox1.Text
            rs.MoveFirst()
            rs.Find(Criteria)
            If rs.EOF Then
                MsgBox("Client ID does not exit")
            Else
                'Dim Criteria2 As String
                'Criteria2 = "Password =" + TextBox2.Text
                'rs.MoveFirst()
                'rs.Find(Criteria)

                If (rs.Fields("ClientID").Value.ToString = TextBox1.Text And rs.Fields("Password").Value.ToString = TextBox2.Text) Then
                    Dim x As Transcript = New Transcript
                    x.Show()

                    'If Form1.CheckBox1.CheckState = 1 Then
                    '    Me.Button2.Enabled = True
                    '    'MessageBox.Show("Checked")
                    'End If
                    'Dim F2 As New Form1() With CheckBox1  ' Object variable
                    'If rs.Fields("AccountType").Value.ToString() =  Then

                    'End If
                Else
                    MessageBox.Show("Invalid Account")
                    TextBox1.Clear()
                    TextBox2.Clear()
                End If
            End If
        End If


    End Sub

    Private Function CheckBox1() As Object
        Throw New NotImplementedException()
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub
End Class