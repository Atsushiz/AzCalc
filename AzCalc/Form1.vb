Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "A'z Calc"
        TextBox_Work.Text = ""
        TextBox_Result.Text = ""

    End Sub

    Private Sub Button_One_Click(sender As Object, e As EventArgs) Handles Button_One.Click
        TextBox_Work.Text = TextBox_Work.Text + "1"

    End Sub

    Private Sub Button_Two_Click(sender As Object, e As EventArgs) Handles Button_Two.Click
        TextBox_Work.Text = TextBox_Work.Text + "2"

    End Sub

    Private Sub Button_Three_Click(sender As Object, e As EventArgs) Handles Button_Three.Click
        TextBox_Work.Text = TextBox_Work.Text + "3"

    End Sub

    Private Sub Button_Four_Click(sender As Object, e As EventArgs) Handles Button_Four.Click
        TextBox_Work.Text = TextBox_Work.Text + "4"

    End Sub

    Private Sub Button_Five_Click(sender As Object, e As EventArgs) Handles Button_Five.Click
        TextBox_Work.Text = TextBox_Work.Text + "5"

    End Sub

    Private Sub Button_Six_Click(sender As Object, e As EventArgs) Handles Button_Six.Click
        TextBox_Work.Text = TextBox_Work.Text + "6"

    End Sub

    Private Sub Button_Seven_Click(sender As Object, e As EventArgs) Handles Button_Seven.Click
        TextBox_Work.Text = TextBox_Work.Text + "7"

    End Sub

    Private Sub Button_Eight_Click(sender As Object, e As EventArgs) Handles Button_Eight.Click
        TextBox_Work.Text = TextBox_Work.Text + "8"

    End Sub

    Private Sub Button_Nine_Click(sender As Object, e As EventArgs) Handles Button_Nine.Click
        TextBox_Work.Text = TextBox_Work.Text + "9"

    End Sub

    Private Sub Button_Plus_Click(sender As Object, e As EventArgs) Handles Button_Plus.Click
        TextBox_Work.Text = TextBox_Work.Text + "+"

    End Sub

    Private Sub Button_Slash_Click(sender As Object, e As EventArgs) Handles Button_Slash.Click
        If (TextBox_Work.Text <> "") Then
            TextBox_Work.Text = TextBox_Work.Text + "/"
        End If

    End Sub

    Private Sub Button_Minus_Click(sender As Object, e As EventArgs) Handles Button_Minus.Click
        TextBox_Work.Text = TextBox_Work.Text + "-"

    End Sub

    Private Sub Button_Asterisk_Click(sender As Object, e As EventArgs) Handles Button_Asterisk.Click
        If (TextBox_Work.Text <> "") Then
            TextBox_Work.Text = TextBox_Work.Text + "*"
        End If
    End Sub

    Private Sub Button_LeftParen_Click(sender As Object, e As EventArgs) Handles Button_LeftParen.Click
        TextBox_Work.Text = TextBox_Work.Text + "("

    End Sub

    Private Sub Button_RightParen_Click(sender As Object, e As EventArgs) Handles Button_RightParen.Click
        TextBox_Work.Text = TextBox_Work.Text + ")"

    End Sub

    Private Sub Button_Equal_Click(sender As Object, e As EventArgs) Handles Button_Equal.Click
        calc()

    End Sub

    Private Sub Button_Zero_Click(sender As Object, e As EventArgs) Handles Button_Zero.Click
        TextBox_Work.Text = TextBox_Work.Text + "0"

    End Sub

    Private Sub Button_Clear_Click(sender As Object, e As EventArgs) Handles Button_Clear.Click
        TextBox_Work.Text = ""

    End Sub

    Private Sub TextBox_Work_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Work.TextChanged
        Dim laststring As String = RTrim(TextBox_Work.Text)


    End Sub

    Private Sub calc()
        '計算式
        Dim exp As String = TextBox_Work.Text

        Dim t As Type =
            Type.GetTypeFromProgID("MSScriptControl.ScriptControl")
        Dim obj As Object = Activator.CreateInstance(t)
        t.InvokeMember("Language",
            System.Reflection.BindingFlags.SetProperty,
            Nothing,
            obj,
            New Object() {"vbscript"})
        'Eval関数で計算を実行して結果を取得
        Dim result As Double = CDbl(
            t.InvokeMember("Eval",
                System.Reflection.BindingFlags.InvokeMethod,
                Nothing,
                obj,
                New Object() {exp}))

        Dim str As String() = result.ToString.Split("."c)
        Dim CustomFormat As String

        ' 配列要素の確認
        If (str.Length = 1) Then
            ' 小数点以下が無い場合
            CustomFormat = "#,0"
        End If

        If (str.Length > 1) Then
            If str(1).Length > 10 Then
                ' 小数点以下が11桁以上の場合
                CustomFormat = "n10"
            Else
                ' 小数点以下が10桁以下の場合
                CustomFormat = "n" + str(1).Length.ToString
            End If
        End If


        '結果を表示
        TextBox_Result.Text = result.ToString(CustomFormat)

    End Sub

    Private Sub Button_Period_Click(sender As Object, e As EventArgs) Handles Button_Period.Click
        TextBox_Work.Text = TextBox_Work.Text + "."

    End Sub
End Class
