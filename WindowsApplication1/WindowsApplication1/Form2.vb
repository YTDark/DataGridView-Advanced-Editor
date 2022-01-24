Imports System.Data.SqlClient
Public Class Form2
    Public STR As String = ""
    Public CON As New SqlConnection("Data Source=Home-PC;Initial Catalog=DATA_;Integrated Security=True")

    Public Function SELECT_DT(ByVal TXT_ As String) As DataTable
        Dim DT As New DataTable
        DT.Clear()
        Dim ADP As New SqlDataAdapter(TXT_, CON)
        ADP.Fill(DT)
        Return DT
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        Form1.DGV1.DataSource = Nothing
        STR = "Select " & COLUM_() & " from BILL_T where  " & Select_text() + ORDER_BY() & ""
        Dim dt As DataTable = SELECT_DT(STR)

        If CheckBox4.Checked = False Then
            Form1.DGV1.DataSource = dt
        Else
            Form1.DGV1.DataSource = SELECT_DT("Select " & COLUM_() & " from BILL_T where " & Select_text() & " UNION  ALL (Select " & COLUM_() & " from BILL_T  except (Select " & COLUM_() & " from BILL_T where " & Select_text() & "))")
            For t As Integer = 0 To dt.Rows.Count - 1
                Form1.DGV1.Rows(t).DefaultCellStyle.BackColor = Label10.BackColor
            Next
        End If
        Me.Visible = False
    End Sub
    Public Function COLUM_() As String
        On Error Resume Next
        Dim str3(13) As String
        If CheckedListBox1.GetItemChecked(0) = True Then str3(0) = "NAME_T,"
        If CheckedListBox1.GetItemChecked(1) = True Then str3(1) = "DATE_T,"
        If CheckedListBox1.GetItemChecked(2) = True Then str3(2) = "ID,"
        If CheckedListBox1.GetItemChecked(3) = True Then str3(3) = "P_ITEM,"
        If CheckedListBox1.GetItemChecked(4) = True Then str3(4) = "BOX_T,"
        If CheckedListBox1.GetItemChecked(5) = True Then str3(5) = "NAME_ITEM,"
        If CheckedListBox1.GetItemChecked(6) = True Then str3(6) = "COST_T,"
        If CheckedListBox1.GetItemChecked(7) = True Then str3(7) = "VAL_ITEM,"
        If CheckedListBox1.GetItemChecked(8) = True Then str3(8) = "Q_ITEM,"
        If CheckedListBox1.GetItemChecked(9) = True Then str3(9) = "STORE_T,"
        If CheckedListBox1.GetItemChecked(10) = True Then str3(10) = "NAME_T,"
        If CheckedListBox1.GetItemChecked(11) = True Then str3(11) = "W_ITEM,"
        If CheckedListBox1.GetItemChecked(12) = True Then str3(12) = "NUMBER_T,"
        If CheckedListBox1.GetItemChecked(13) = True Then str3(13) = "TYPE_T,"
        Dim STR2 As String = ""
        For i As Integer = 0 To str3.Count
            STR2 += str3(i)
        Next
        STR2 = RSet(STR2, (Len(STR2) - 1)) 'لحذف اخر فاصلة في السطر
        Return STR2 + Form3.STR_3
    End Function
    Public Function Select_text() As String
        On Error Resume Next
        Dim STR1(15) As String
        STR1(0) = "DATE_T between '" & Format(DateTimePicker1.Value, "yyyy/MM/dd") & "'AND'" & Format(DateTimePicker2.Value, "yyyy/MM/dd") & "'"
        If Not ComboBox3.Text = "" Then STR1(1) = "and COST_T like '" & ComboBox3.Text & "'"
        If Not ComboBox5.Text = "" Then STR1(2) = "and STORE_T like '" & ComboBox5.Text & "'"
        If Not ComboBox4.Text = "" Then STR1(3) = "and NAME_ITEM like '" & ComboBox4.Text & "'"
        If Not ComboBox6.Text = "" Then STR1(4) = "and W_ITEM like '" & ComboBox6.Text & "'"
        If Not ComboBox7.Text = "" Then STR1(5) = "and TO_T like '" & ComboBox7.Text & "'"

        If CheckBox1.Checked = True Then
            If RadioButton1.Checked = True Then STR1(6) = "and Q_ITEM > '" & Val(TextBox1.Text) & "'"
            If RadioButton2.Checked = True Then STR1(7) = "and Q_ITEM < '" & Val(TextBox2.Text) & "'"
            If RadioButton5.Checked = True Then STR1(8) = "and Q_ITEM between '" & Val(TextBox1.Text) & "'and '" & Val(TextBox2.Text) & "'"
        End If
        If CheckBox2.Checked = True Then
            If RadioButton1.Checked = True Then STR1(9) = "and P_ITEM > '" & Val(TextBox1.Text) & "'"
            If RadioButton2.Checked = True Then STR1(10) = "and P_ITEM < '" & Val(TextBox2.Text) & "'"
            If RadioButton5.Checked = True Then STR1(11) = "and P_ITEM between '" & Val(TextBox1.Text) & "'and '" & Val(TextBox2.Text) & "'"
        End If
        If CheckBox3.Checked = True Then
            If RadioButton1.Checked = True Then STR1(12) = "and VAL_ITEM > '" & Val(TextBox1.Text) & "'"
            If RadioButton2.Checked = True Then STR1(13) = "and VAL_ITEM < '" & Val(TextBox2.Text) & "'"
            If RadioButton5.Checked = True Then STR1(14) = "and VAL_ITEM between '" & Val(TextBox1.Text) & "'and '" & Val(TextBox2.Text) & "'"
        End If
        If Not TextBox3.Text = "" And RadioButton3.Checked = True Then STR1(15) = "and NOTE_T like '%" & TextBox3.Text & "%'"
        If Not TextBox3.Text = "" And RadioButton4.Checked = True Then STR1(16) = "and not NOTE_T  like '%" & TextBox3.Text & "%'"
        Dim STR2 As String = ""
        For i As Integer = 0 To STR1.Count
            STR2 += STR1(i)
        Next
        Return STR2
    End Function

    Public Function ORDER_BY() As String
        On Error Resume Next
        Dim str3 As String = ""
        Dim str4 As String = ""
        If ComboBox1.Text = "" Or ComboBox1.Text = "اساسي" Then str3 = "ORDER BY DATE_T,NUMBER_T,ID"
        If ComboBox1.Text = "المستند" Then str3 = "ORDER BY NAME_T"
        If ComboBox1.Text = "رقم المستند" Then str3 = "ORDER BY NUMBER_T"
        If ComboBox1.Text = "التاريخ" Then str3 = "ORDER BY DATE_T"
        If ComboBox1.Text = "الصنف" Then str3 = "ORDER BY NAME_ITEM"
        If ComboBox1.Text = "الوحدة" Then str3 = "ORDER BY W_ITEM"
        If ComboBox1.Text = "الكمية" Then str3 = "ORDER BY Q_ITEM"
        If ComboBox1.Text = "السعر" Then str3 = "ORDER BY P_ITEM"
        If ComboBox1.Text = "القيمة" Then str3 = "ORDER BY VAL_ITEM"
        If ComboBox2.Text = "تنازلي" Then str4 = "DESC"
        If ComboBox2.Text = "تصاعدي" Then str4 = "ASC"
        Return str3 + " " + str4
    End Function

    Private Sub ComboBox3_DropDown(sender As Object, e As EventArgs) Handles ComboBox3.DropDown
        ComboBox3.DataSource = SELECT_DT("Select distinct cost_t from BILL_T")
        ComboBox3.DisplayMember = "cost_t"
    End Sub

    Private Sub ComboBox4_DropDown(sender As Object, e As EventArgs) Handles ComboBox4.DropDown
        ComboBox4.DataSource = SELECT_DT("Select distinct STORE_T from BILL_T")
        ComboBox4.DisplayMember = "STORE_T"
    End Sub
    Private Sub ComboBox5_DropDown(sender As Object, e As EventArgs) Handles ComboBox5.DropDown
        ComboBox5.DataSource = SELECT_DT("Select distinct NAME_ITEM from BILL_T")
        ComboBox5.DisplayMember = "NAME_ITEM"
    End Sub

    Private Sub ComboBox6_DropDown(sender As Object, e As EventArgs) Handles ComboBox6.DropDown
        ComboBox6.DataSource = SELECT_DT("Select distinct W_ITEM from BILL_T")
        ComboBox6.DisplayMember = "W_ITEM"
    End Sub
    Private Sub ComboBox7_DropDown(sender As Object, e As EventArgs) Handles ComboBox7.DropDown
        ComboBox7.DataSource = SELECT_DT("Select distinct TO_T from BILL_T")
        ComboBox7.DisplayMember = "TO_T"
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = "1 / " & Month(Now) & " / " & Year(Now) & ""
        DateTimePicker2.Value = Date.Now
        RadioButton1.Checked = True
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        On Error Resume Next
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
            CheckBox3.Checked = False
        Else

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        On Error Resume Next
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
            CheckBox3.Checked = False
        Else

        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        On Error Resume Next
        If CheckBox3.Checked = True Then
            CheckBox1.Checked = False
            CheckBox2.Checked = False
        Else

        End If
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        On Error Resume Next
        Dim color_ As DialogResult = ColorDialog1.ShowDialog()
        If color_ = Windows.Forms.DialogResult.OK Then
            Label10.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Me.Visible = False
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        If Form3.Visible = False Then
            Form3.Visible = True
        Else
            Form3.Visible = False
        End If
    End Sub
End Class