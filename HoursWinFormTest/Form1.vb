Imports HoursLibrary.Classes
Imports HoursLibrary.LanguageExtensions

Public Class Form1
    Private Sub TimeSpan1ValuesButton_Click(sender As Object, e As EventArgs) Handles TimeSpan1ValuesButton.Click
        MessageBox.Show($"No formatting: {TimeComboBox1.TimeSpan}{Environment.NewLine}Formatted: {TimeComboBox1.TimeSpan.Formatted()}")
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        IncrementsComboBox.DataSource = [Enum].GetValues(GetType(TimeIncrement)).Cast(Of TimeIncrement)().ToList()

        TimeComboBox1.Increment = TimeIncrement.Quarterly
        IncrementsComboBox.SelectedIndex = IncrementsComboBox.FindString(TimeComboBox1.Increment.ToString())

        Dim position = TimeComboBox1.FindString(Now.RoundedQuarterly().ToString(Hours.TimeFormat))

        TimeComboBox1.SelectedIndex = If(position > -1, position, 0)

        AddHandler IncrementsComboBox.SelectedIndexChanged,
            AddressOf IncrementsComboBox_SelectedIndexChanged

    End Sub

    Private Sub IncrementsComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)

        Dim current As String = TimeComboBox1.TimeSpan.CurrentHour()
        Dim selectedItem = CType(IncrementsComboBox.SelectedItem, TimeIncrement)

        TimeComboBox1.Increment = selectedItem
        TimeComboBox1.SelectedIndex = 0

        Dim position = TimeComboBox1.FindString(current)

        TimeComboBox1.SelectedIndex = If(position > -1, position, 0)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dataResults = DataOperations.Read()

        For Each dataResult As TimeTable In dataResults
            Debug.WriteLine($"{dataResult.Id}  {dataResult.StartTime.Value.Formatted()}")
        Next

    End Sub
End Class

