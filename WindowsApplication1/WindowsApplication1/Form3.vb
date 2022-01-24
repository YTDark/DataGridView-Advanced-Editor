Public Class Form3
    Public STR_3 As String = ""
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        For i As Integer = 0 To DGV1.Rows.Count - 2
            If DGV1.Rows(i).Cells(2).Value = "=" Or DGV1.Rows(i).Cells(2).Value = "<" Or DGV1.Rows(i).Cells(2).Value = ">" Then
                STR_3 += ",(case when " & COLUM_(i, 1) + DGV1.Rows(i).Cells(2).Value.ToString + "'" + DGV1.Rows(i).Cells(3).Value + "'" & "  then " & COLUM_(i, 1) & " end)as " & DGV1.Rows(i).Cells(0).Value.ToString & ""
            Else
                STR_3 += (",( " & COLUM_(i, 1) + DGV1.Rows(i).Cells(2).Value.ToString + "'" + DGV1.Rows(i).Cells(3).Value + "'" & " )as " & DGV1.Rows(i).Cells(0).Value.ToString & "")
            End If
        Next
        Me.Visible = False
    End Sub
   
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        Dim bb As DataGridViewComboBoxColumn 'لتعبئه كمبوبوكس من الداتا قريد
        bb = DGV1.Columns(1)
        bb.Items.Clear()
        For i As Integer = 0 To Form2.CheckedListBox1.Items.Count - 1
            bb.Items.Add(Form2.CheckedListBox1.Items(i).ToString)
        Next
        bb = DGV1.Columns(2)
        bb.Items.Add("=")
        bb.Items.Add("<")
        bb.Items.Add(">")
        bb.Items.Add("+")
        bb.Items.Add("-")
        bb.Items.Add("*")
        bb.Items.Add("/")
    End Sub
  
    Public Function COLUM_(ii As Integer, iii As Integer) As String
        On Error Resume Next
        Dim str3 As String = ""
        If DGV1.Rows(ii).Cells(iii).Value = "الاسم" Then str3 = "to_T"
        If DGV1.Rows(ii).Cells(iii).Value = "التاريخ" Then str3 = "DATE_T"
        If DGV1.Rows(ii).Cells(iii).Value = "رقم المتسلسل" Then str3 = "ID"
        If DGV1.Rows(ii).Cells(iii).Value = "السعر" Then str3 = "P_ITEM"
        If DGV1.Rows(ii).Cells(iii).Value = "الصندوق" Then str3 = "BOX_T"
        If DGV1.Rows(ii).Cells(iii).Value = "الصنف" Then str3 = "NAME_ITEM"
        If DGV1.Rows(ii).Cells(iii).Value = "الفرع" Then str3 = "COST_T"
        If DGV1.Rows(ii).Cells(iii).Value = "القيمة" Then str3 = "VAL_ITEM"
        If DGV1.Rows(ii).Cells(iii).Value = "الكمية" Then str3 = "Q_ITEM"
        If DGV1.Rows(ii).Cells(iii).Value = "المخزن" Then str3 = "STORE_T"
        If DGV1.Rows(ii).Cells(iii).Value = "المستند" Then str3 = "NAME_T"
        If DGV1.Rows(ii).Cells(iii).Value = "الوحدة" Then str3 = "W_ITEM"
        If DGV1.Rows(ii).Cells(iii).Value = "رقم المستند" Then str3 = "NUMBER_T"
        If DGV1.Rows(ii).Cells(iii).Value = "نوع المستند" Then str3 = "TYPE_T"
        Return str3
    End Function
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Me.Visible = False
    End Sub
End Class