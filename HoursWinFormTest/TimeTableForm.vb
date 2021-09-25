Imports HoursLibrary.LanguageExtensions

Public Class TimeTableForm
    WithEvents TimesBindingSource As New BindingSource
    Private Sub TimeTableForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        TimesBindingSource.DataSource = DataOperations.Read()
        DataListBox.DataSource = TimesBindingSource

        AddHandler StartTimeTimePickerComboBox.SelectedIndexChanged,
            AddressOf StartTimeTimePickerComboBox_SelectedIndexChanged
    End Sub

    Private Sub TimesBindingSource_PositionChanged(sender As Object, e As EventArgs) Handles TimesBindingSource.PositionChanged

        Dim current = CType(TimesBindingSource.Current, TimeTable)
        Dim position = StartTimeTimePickerComboBox.FindString(current.StartTime.Value.Formatted())

        StartTimeTimePickerComboBox.SelectedIndex = If(position > -1, position, 0)

    End Sub

    Private Sub StartTimeTimePickerComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)

        Dim time As TimeSpan

        If TimeSpan.TryParse(StartTimeTimePickerComboBox.TimeSpan.ToString(), time) Then
            CType(TimesBindingSource.Current, TimeTable).StartTime = time
        End If

    End Sub
End Class