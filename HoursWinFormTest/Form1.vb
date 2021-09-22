Imports HoursLibrary

Public Class Form1
    Private Sub TimeSpan1ValuesButton_Click(sender As Object, e As EventArgs) Handles TimeSpan1ValuesButton.Click
        MessageBox.Show($"No formatting: {TimeComboBox1.TimeSpan.ToString()}")
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        IncrementsComboBox.DataSource = System.Enum.GetValues(GetType(TimeIncrement)).
            Cast(Of TimeIncrement)().ToList()

        TimeComboBox1.Increment = TimeIncrement.Quarterly
        IncrementsComboBox.SelectedIndex = IncrementsComboBox.FindString(TimeComboBox1.Increment.ToString())

        Dim position = TimeComboBox1.FindString(Now.RoundedQuarterly().ToString(Hours.TimeFormat))

        TimeComboBox1.SelectedIndex = If(position > -1, position, 0)

    End Sub
End Class
