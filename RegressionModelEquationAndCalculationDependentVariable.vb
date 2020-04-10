Public Class RegressionModelEquation
    'Объявление переменных
    Private Grid As DataGridStatistControl
    Private Graph As BigFPicShow
    Public MyParentForm As MainForm
    Public NameOfAnalysis As String
    Public ResForName As New RegressionSummary.OutCollection
    Public EquationRes As STATISTICA.Spreadsheet
    Dim add = 40
    Dim result = 0
    Dim textboxarr(1)

    Private Sub RegressionModelEquation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim regres As New MultipleRegression
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Dim y = 0
        'Отрисовываем лейблы с названием значимых переменных и поля для ввода значений
        ReDim textboxarr(EquationRes.Data.GetLength(0) - 1)
        For i = 0 To ListOfVariable.ListName.Length - 1

            y = y + add
            Dim Label1 As New Label
            Label1.Top = y
            Label1.Left = 40
            Label1.AutoSize = True
            Label1.Text = ListOfVariable.ListName(i)
            Label1.Visible = True
            Me.Controls.Add(Label1)

            Dim TextBox1 As New TextBox
            TextBox1.Top = y
            TextBox1.Left = 100
            TextBox1.Visible = True
            textboxarr(i) = TextBox1
            Me.Controls.Add(TextBox1)
        Next
    End Sub

    'Предсказание зависимой переменной
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        result = 0
        For i = 2 To EquationRes.Data.GetLength(0)
            result = result + (CDbl(textboxarr(i - 1).Text) * Math.Round(EquationRes.Data.GetValue(i, 4), 3))
        Next
        result = result + Math.Round(EquationRes.Data.GetValue(1, 4), 3)
        Label2.Text = CStr(result)
        Label1.Visible = True
        Label2.Visible = True
    End Sub


    'Вывод уравнения множественной регрессии
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim equation As New Form2
        Label3.Text = ""
        For i = 2 To EquationRes.Data.GetLength(0)
            Label3.Text = Label3.Text + CStr(Math.Round(EquationRes.Data.GetValue(i, 4), 3)) + "x" + CStr(i - 1) + " + "
        Next
        Label3.Text = Label3.Text + CStr(Math.Round(EquationRes.Data.GetValue(1, 4), 3))
        equation.Label1.Text = ""
        equation.Label1.Text = Label3.Text
        equation.MyParentForm = MyParentForm
        equation.NameOfAnalysis = "Уравнение"
        equation.Show()
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
End Class