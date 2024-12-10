Imports MySql.Data.MySqlClient
Imports Microsoft.VisualBasic.ApplicationServices
Imports Microsoft.Win32
Imports System.Data.OleDb

Public Class Form1

    Dim sqlConn As New MySqlConnection
    Dim sqlCmd As New MySqlCommand
    Dim sqlRd As MySqlDataReader
    Dim sqlDt As New DataTable
    Dim DtA As New MySqlDataAdapter
    Dim sqlQuery As String

    Dim server As String = "localhost"
    Dim Username As String = "root"
    Dim Password As String = "12345"
    Dim database As String = "massivevac"


    Dim flgHighlight As Boolean


    Private Sub updateTable()

        sqlConn.ConnectionString = "server =" + server + ";" + "user id=" + Username + ";" _
         + "password=" + Password + ";" + "database =" + database
        'sqlConn.ConnectionString = "server=localhost;userid=root;password=12345;database=dailytrans"

        sqlConn.Open()
        sqlCmd.Connection = sqlConn
        sqlCmd.CommandText = "SELECT * From massivevac.massivevac"

        sqlRd = sqlCmd.ExecuteReader
        sqlDt.Load(sqlRd)
        sqlRd.Close()
        sqlConn.Close()
        DataGridView1.DataSource = sqlDt

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView1.Rows(e.RowIndex)

            cmbBarangay.Text = row.Cells("BARANGAY").Value.ToString
            cmbPurok.Text = row.Cells("PUROK").Value.ToString
            cmbBHW.Text = row.Cells("BHW").Value.ToString
            cmbVaccinator.Text = row.Cells("VACCINATOR").Value.ToString
            cmbLocation.Text = row.Cells("LOCATION").Value.ToString
            cmbOwnerName.Text = row.Cells("OWNERNAME").Value.ToString
            cmbGender.Text = row.Cells("GENDER").Value.ToString
            txtAge.Text = row.Cells("AGE").Value.ToString
            cmbAnimalName.Text = row.Cells("ANIMALNAME").Value.ToString
            cmbAnimalage.Text = row.Cells("ANIMALAGE").Value.ToString
            cmbColor.Text = row.Cells("COLOR").Value.ToString
            cmbSex.Text = row.Cells("SEX").Value.ToString
            cmbType.Text = row.Cells("ANIMALTYPE").Value.ToString
            txtHeads.Text = row.Cells("HEADS").Value.ToString
            cmbLocation.Text = row.Cells("LOCATION").Value.ToString
            cmbEncoded.Text = row.Cells("ENCODEDBY").Value.ToString
        End If

    End Sub




    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim iExit As DialogResult

        iExit = MessageBox.Show("Are you sure you want to Exit?", "MySql Connector", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If iExit = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        updateTable()
        Timer1.Enabled = True
    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        sqlConn.ConnectionString = "server =" + server + ";" + "user id=" + Username + ";" _
         + "password=" + Password + ";" + "database =" + database



        Try
            sqlConn.Open()
            sqlQuery = "insert into massivevac.massivevac (OWNERNAME, GENDER, AGE, ANIMALNAME, ANIMALAGE, COLOR, SEX, ANIMALTYPE, 
                        HEADS, DATEVACCINATED, BARANGAY, PUROK, BHW, VACCINATOR, LOCATION, ENCODEDBY, TIME )
                    values('" & cmbOwnerName.Text & "','" & cmbGender.Text & "','" & txtAge.Text & "','" & cmbAnimalName.Text & "','" & cmbAnimalage.Text & "','" & cmbColor.Text & "','" & cmbSex.Text & "','" & cmbType.Text & "','" & txtHeads.Text & "','" & dtDateVac.Text & "','" & cmbBarangay.Text & "','" & cmbPurok.Text & "','" & cmbBHW.Text & "','" & cmbVaccinator.Text & "','" & cmbLocation.Text & "','" & cmbEncoded.Text & "','" & TIME.Text & "')"

            sqlCmd = New MySqlCommand(sqlQuery, sqlConn)
            sqlRd = sqlCmd.ExecuteReader
            sqlConn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "MySQL Connector", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Finally
            sqlConn.Dispose()
        End Try
        updateTable()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TIME.Text = Date.Now.ToString("hh:mm:ss")
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            DataGridView1.Rows.Remove(row)
        Next
        updateTable()
    End Sub



    Private Sub txtConsole_TextChanged(sender As Object, e As EventArgs) Handles txtConsole.TextChanged
        If txtConsole.Text = "" Then initValue()

    End Sub

    Private Sub txtConsole_KeyDown(sender As Object, e As KeyEventArgs) Handles txtConsole.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnNew_Click(Nothing, Nothing)
            flgHighlight = False
        Else
            Exit Sub
        End If

        e.SuppressKeyPress = True
    End Sub

    Sub initValue()
        cmbOwnerName.Text = ""
        cmbGender.Text = ""
        txtAge.Text = 1
        cmbAnimalName.Text = ""
        cmbAnimalage.Text = 1
        cmbColor.Text = "Black"
        cmbSex.Text = ""
        cmbType.Text = ""
        txtHeads.Text = 1

        flgHighlight = False
    End Sub

    Private Sub txtConsole_KeyUp(sender As Object, e As EventArgs) Handles txtConsole.KeyUp
        extractDATA(txtConsole.Text)
    End Sub
    Sub extractDATA(st As String)
        Dim temp() As String
        Dim c As Integer
        Dim neym As String


        ' On Error Resume Next

        'for Lastname, Firstname
        'check for comma if there is comma then first word is lastname. if there is no comma then last word is lastname
        Console.Text = "Input OWNER NAME (Lastname, Firstname) or (Firstname Lastname)"
        txtConsole.ForeColor = Color.Black

        st = StrConv(LTrim(st), vbProperCase)
        If st = "" Then Exit Sub
        neym = Trim(Split(st, ";")(0))
        cmbOwnerName.Text = neym
        If InStr(1, neym, ",") = 0 Then
            temp = Split(neym, " ")
            cmbOwnerName.Text = temp(UBound(temp)) & IIf(UBound(temp) > 0, ",", "")
            If UBound(temp) > 0 Then
                For c = 0 To UBound(temp) - 1
                    cmbOwnerName.Text = cmbOwnerName.Text & " " & temp(c)
                Next c
            End If
        End If

        If UBound(Split(st, ";")) = 0 Then Exit Sub
        Console.Text = "Input GENDER OF OWNER (M or F)"

        st = UCase(LTrim(Split(st, ";")(1)))
        If st = "" Then Exit Sub
        temp = Split(st, " ")

        'for gender
        cmbGender.Text = "Male"

        Select Case LCase(temp(0))
            Case = "m" : cmbGender.Text = "Male"
            Case = "f" : cmbGender.Text = "Female"
            Case = "o" : cmbGender.Text = "Other"

        End Select



        If UBound(temp) < 1 Then Exit Sub
        Console.Text = "INPUT ANIMAL AGE"
        txtAge.Text = temp(1)

        'for Animal name
        If UBound(temp) < 2 Then Exit Sub
        Console.Text = "INPUT ANIMAL NAME (NO SPACE)"
        cmbAnimalName.Text = temp(2)

        'for Animal age
        If UBound(temp) < 3 Then Exit Sub
        Console.Text = "INPUT ANIMAL AGE"
        cmbAnimalage.Text = temp(3)

        'for Animal Color
        If UBound(temp) < 4 Then Exit Sub
        Console.Text = "INPUT COLOR OF THE ANIMAL"
        Select Case LCase(temp(4))
            Case = "w" : cmbColor.Text = "White"
            Case = "br" : cmbColor.Text = "Brown"
            Case = "b" : cmbColor.Text = "Black"
            Case = "g" : cmbColor.Text = "Gray"
            Case = "t" : cmbColor.Text = "Tiger"
            Case = "k" : cmbColor.Text = "Kabang"
            Case = "o" : cmbColor.Text = "Orange"
                Exit Sub
        End Select

        'for animal Sex
        If UBound(temp) < 5 Then Exit Sub
        Console.Text = "INPUT ANIMAL SEX"
        Select Case LCase(temp(5))
            Case = "m" : cmbSex.Text = "Male"
            Case = "f" : cmbSex.Text = "Female"
        End Select

        'for animal type
        If UBound(temp) < 6 Then Exit Sub
        Console.Text = "PLEASE INPUT ANIMAL TYPE (DOG, CAT, RABBIT, MONKEY, HORSE, GOAT, COW, CARABAO)"
        Select Case LCase(temp(6))
            Case = "d" : cmbType.Text = "Dog"
            Case = "c" : cmbType.Text = "Cat"
            Case = "r" : cmbType.Text = "Rabbit"
            Case = "m" : cmbType.Text = "Monkey"
            Case = "h" : cmbType.Text = "Horse"
            Case = "g" : cmbType.Text = "Goat"
            Case = "co" : cmbType.Text = "Cow"
            Case = "ca" : cmbType.Text = "Carabao"
            Case Else
                txtConsole.ForeColor = Color.Red
                Exit Sub
        End Select
        If UBound(temp) < 7 Then Exit Sub
        Console.Text = "NO. OF HEADS"
        txtHeads.Text = temp(7)

    End Sub



    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        Dim DV As New DataView(sqlDt)
        DV.RowFilter = String.Format("OWNERNAME like '%{0}%' or ANIMALNAME like '%{0}%' or BARANGAY like '%{0}%'", txtSearch.Text)
        DataGridView1.DataSource = DV
    End Sub

    Private Sub cmbEncoded_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEncoded.SelectedIndexChanged

    End Sub
End Class