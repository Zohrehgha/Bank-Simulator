Public Class Transcript
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset    'table accountHolder
    Dim rs2 As New ADODB.Recordset   'table transcript
    Private Sub Transcript_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'setting up the ADO objects
        conn.Provider = "Microsoft.Jet.OleDB.4.0"
        ' Setting up the jet DB driver (access) 
        conn.ConnectionString = "C:\ITD\Term 3\Visual Basic.Net\assignment\assignment 2\Database2.mdb"
        conn.Open()
        rs.Open("select * from AccountHolders", conn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic)
        rs2.Open("select * from Transcript", conn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic)
        Label2.Enabled = False
        Label6.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        Button1.Enabled = False
    End Sub
    'perform all transactions by this button
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox9.Text = "" Then
            MessageBox.Show("Please Fill up the blank boxes")
        ElseIf ComboBox2.Text = "Withdraw" Then
            Call Withdraw()
            rs2.AddNew()
            Call SaveinTable()
            rs.Update()
            rs2.Update()
            MessageBox.Show("Success")
            TextBox9.Clear()

        ElseIf ComboBox2.Text = "Deposit" Then
            Call Deposit()
            rs2.AddNew()
            Call SaveinTable()
            rs.Update()
            rs2.Update()
            MessageBox.Show("Success")
            TextBox9.Clear()

        ElseIf ComboBox2.Text = "Transfer" Then
            Call Transfer()
            rs2.AddNew()
            Call SaveinTable()
            rs.Update()
            rs2.Update()
            MessageBox.Show("Success")
            TextBox9.Clear()
        End If
    End Sub
    'perform the Withdraw transaction 
    Private Sub Withdraw()
        If ComboBox1.Text = "Checking" And TextBox9.Text <= TextBox5.Text Then
            TextBox5.Text = Convert.ToDouble(TextBox5.Text) - Convert.ToDouble(TextBox9.Text)

        ElseIf ComboBox1.Text = "Saving" And TextBox9.Text <= TextBox6.Text Then
            TextBox6.Text = Convert.ToDouble(TextBox6.Text) - Convert.ToDouble(TextBox9.Text)

        ElseIf ComboBox1.Text = "Visa" And TextBox9.Text <= TextBox7.Text Then
            TextBox7.Text = Convert.ToDouble(TextBox7.Text) - Convert.ToDouble(TextBox9.Text)
        End If
    End Sub
    'perform the Deposit transaction
    Private Sub Deposit()
        If ComboBox1.Text = "Checking" Then
            TextBox5.Text = Convert.ToDouble(TextBox5.Text) + Convert.ToDouble(TextBox9.Text)
        ElseIf ComboBox1.Text = "Saving" Then
            TextBox6.Text = Convert.ToDouble(TextBox6.Text) + Convert.ToDouble(TextBox9.Text)
        ElseIf ComboBox1.Text = "Visa" Then
            TextBox7.Text = Convert.ToDouble(TextBox7.Text) + Convert.ToDouble(TextBox9.Text)
        End If
    End Sub
    'perform the Transfer transaction from one account to another
    Private Sub Transfer()
        If ComboBox3.Text = "Checking" And TextBox9.Text <= TextBox5.Text And ComboBox4.Text = "Saving" Then
            TextBox5.Text = Convert.ToDouble(TextBox5.Text) - Convert.ToDouble(TextBox9.Text)
            TextBox6.Text = Convert.ToDouble(TextBox6.Text) + Convert.ToDouble(TextBox9.Text)
        ElseIf ComboBox3.Text = "Checking" And TextBox9.Text <= TextBox5.Text And ComboBox4.Text = "Visa" Then
            TextBox7.Text = Convert.ToDouble(TextBox7.Text) + Convert.ToDouble(TextBox9.Text)
            TextBox5.Text = Convert.ToDouble(TextBox5.Text) - Convert.ToDouble(TextBox9.Text)
        ElseIf ComboBox3.Text = "Saving" And TextBox9.Text <= TextBox6.Text And ComboBox4.Text = "Checking" Then
            TextBox5.Text = Convert.ToDouble(TextBox5.Text) + Convert.ToDouble(TextBox9.Text)
            TextBox6.Text = Convert.ToDouble(TextBox6.Text) - Convert.ToDouble(TextBox9.Text)
        ElseIf ComboBox3.Text = "Saving" And TextBox9.Text <= TextBox6.Text And ComboBox4.Text = "Visa" Then
            TextBox7.Text = Convert.ToDouble(TextBox7.Text) + Convert.ToDouble(TextBox9.Text)
            TextBox6.Text = Convert.ToDouble(TextBox6.Text) - Convert.ToDouble(TextBox9.Text)
        End If
    End Sub
    'save all modifications of transactions in DB
    Private Sub SaveinTable()
        rs.Fields("ClientID").Value = TextBox1.Text
        rs.Fields("Password").Value = TextBox2.Text
        rs2.Fields("dateTime").Value = Convert.ToDateTime(DateTimePicker2.Text)
        rs.Fields("AccountNumber").Value = TextBox4.Text
        rs.Fields("BranchNumber").Value = TextBox3.Text
        rs.Fields("AccountType").Value = ComboBox1.Text
        rs.Fields("CheckingBalance").Value = TextBox5.Text
        rs.Fields("SavingBalance").Value = TextBox6.Text
        rs.Fields("VisaBalance").Value = TextBox7.Text
        rs2.Fields("BranchNumber").Value = TextBox3.Text
        rs2.Fields("CheckingBalance").Value = TextBox5.Text
        rs2.Fields("SavingBalance").Value = TextBox6.Text
        rs2.Fields("VisaBalance").Value = TextBox7.Text
        rs2.Fields("Amount").Value = TextBox9.Text
        rs2.Fields("TransactionType").Value = ComboBox2.Text
        rs2.Fields("ClientID").Value = TextBox1.Text
    End Sub
    ' transfer transaction needs "From" and "To"
    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Transfer" Then
            Label2.Enabled = True
            Label6.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox1.Enabled = False
            Label7.Enabled = False
        End If
    End Sub
    'verify clientID and Password by reading the info from the database 
    Private Sub Button3_Click_2(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Plz Fill Client ID and Password")
        Else
            Dim Criteria As String
            Criteria = "ClientID =" + TextBox1.Text
            rs.MoveFirst()
            rs.Find(Criteria)
            If rs.EOF Then
                MsgBox("Client ID does not exit")
            Else
                If (rs.Fields("ClientID").Value.ToString = TextBox1.Text And rs.Fields("Password").Value.ToString = TextBox2.Text) Then
                    MessageBox.Show("Login succesfully")
                    Call showdata()
                    Button1.Enabled = True
                Else
                    MessageBox.Show("Invalid Account")
                    TextBox1.Clear()
                    TextBox2.Clear()
                End If
            End If
        End If
    End Sub
    'read all Clients accounts from Data Base
    Private Sub showdata()
        TextBox3.Text = rs.Fields("BranchNumber").Value.ToString()
        TextBox4.Text = rs.Fields("AccountNumber").Value.ToString()
        TextBox5.Text = rs.Fields("CheckingBalance").Value.ToString()
        TextBox6.Text = rs.Fields("SavingBalance").Value.ToString()
        TextBox7.Text = rs.Fields("VisaBalance").Value.ToString()
        TextBox8.Text = rs.Fields("CreditLimit").Value.ToString()
    End Sub
    'search and view all transcripts filter between specific dates
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Criteria As String
        Criteria = ""
        If (TextBox1.Text <> "") Then
            Criteria = Criteria + "ClientID = " + TextBox1.Text
        End If
        If (DateTimePicker1.Text <> "") Then
            If Criteria <> "" Then
                Criteria = Criteria + " AND dateTime = '" + Convert.ToDateTime(DateTimePicker1.Text) + "'"
            Else
                Criteria = Criteria + "dateTime = '" + Convert.ToDateTime(DateTimePicker1.Text) + "'"
            End If
        End If
        If (DateTimePicker3.Text <> "") Then
            If Criteria <> "" Then
                Criteria = Criteria + " AND dateTime = '" + Convert.ToDateTime(DateTimePicker3.Text) + "'"
            Else
                Criteria = Criteria + "dateTime = '" + Convert.ToDateTime(DateTimePicker3.Text) + "'"
            End If
        End If
        rs2.MoveFirst()
        'go to the beginning to start serach 
        rs2.Filter = Criteria
        ' Either We find the record(s), which is the first record if there are more than one
        ' If record is found the file pointer stays at it
        'if not found, the file pointer has passed eof meaning eof = true
        If rs2.EOF Then
            'not found
            MessageBox.Show("Record with this ID does not exist")
            Exit Sub
        Else
            'found
            'showreport()
            MessageBox.Show("Record found succesfully")
            Call showreport()
            rs2.Filter = ""
        End If
    End Sub
    'read all transaction from transcript table in DataGridView
    Private Sub showreport()
        DataGridView1.ColumnCount = 8
        DataGridView1.Columns(0).Name = "ClientID"
        DataGridView1.Columns(1).Name = "Date"
        DataGridView1.Columns(2).Name = "BranchNumber"
        DataGridView1.Columns(3).Name = "TransactionType"
        DataGridView1.Columns(4).Name = "CheckingBalance"
        DataGridView1.Columns(5).Name = "SavingBalance"
        DataGridView1.Columns(6).Name = "VisaBalance"
        DataGridView1.Columns(7).Name = "Amount"
        Dim row As ArrayList = New ArrayList
        row.Add(rs2.Fields("ClientID").Value.ToString())
        row.Add(rs2.Fields("dateTime").Value.ToString())
        row.Add(rs2.Fields("BranchNumber").Value.ToString())
        row.Add(rs2.Fields("TransactionType").Value.ToString())
        row.Add(rs2.Fields("CheckingBalance").Value.ToString())
        row.Add(rs2.Fields("SavingBalance").Value.ToString())
        row.Add(rs2.Fields("VisaBalance").Value.ToString())
        row.Add(rs2.Fields("Amount").Value.ToString())
        DataGridView1.Rows.Add(row.ToArray())
    End Sub
    'Private Sub dg()
    '    'Dim ds As New DataSet
    '    'Dim dt As New DataTable
    '    ''Dim da As New OleDbDataAdapter
    '    ''da = New OleDbDataAdapter()
    '    ''rs2.Fill(dt)
    '    'DataGridView1.DataSource = dt.DefaultView
    '    'ds.Tables.Add(dt)
    '    DataGridView1.Text = rs2.Fields("dateTime").Value.ToString()
    '    DataGridView1.Text = rs2.Fields("ClientID").Value.ToString()
    '    DataGridView1.Text = rs2.Fields("CheckingBalance").Value.ToString()
    '    DataGridView1.Text = rs2.Fields("SavingBalance").Value.ToString()
    '    DataGridView1.Text = rs2.Fields("VisaBalance").Value.ToString()
    '    DataGridView1.Text = rs2.Fields("Amount").Value.ToString()
    '    DataGridView1.Text = rs2.Fields("TransactionType").Value.ToString()
    'End Sub
End Class