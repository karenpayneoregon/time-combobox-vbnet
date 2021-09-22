Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Globalization
Imports System.Windows.Forms.Design

Public Class TimeComboBox
    Inherits ComboBox

    Private _shown As Boolean = False
    Public Sub New()
        DropDownStyle = ComboBoxStyle.DropDownList

        If Not DesignMode Then

            Dim hoursInstance As New Hours()
            Items.AddRange(hoursInstance.Range(Increment))

        End If

        Size = New Size(80, 21)
        MaxDropDownItems = 10
        IntegralHeight = False

    End Sub
    ''' <summary>
    ''' Disable sorting
    ''' </summary>
#Disable Warning BC108, BC114
        Public Overloads ReadOnly Property Sorted() As Boolean
        Get
            Return False
        End Get
    End Property
#Enable Warning BC108, BC114
    ''' <summary>
    ''' Set current item in the ComboBox using a TimeSpan. 
    ''' If the value passed in is not in the ComboBox -1 is returned.
    ''' </summary>
    ''' <param name="pTime"></param>
    ''' <returns>Index of item or -1 if not found</returns>
    Public Function SetCurrentItem(pTime As TimeSpan) As Integer
        Dim dateInstance = Date.Today.Add(pTime)
        Dim displayTime = dateInstance.ToString("hh:mm tt")
        Dim index = FindString(displayTime)
        SelectedIndex = index

        Return index

    End Function
    ''' <summary>
    ''' Set current item by string which represents a valid TimeSpan
    ''' </summary>
    ''' <param name="pTime"></param>
    ''' <returns></returns>
    Public Function SetCurrentItem(ByVal pTime As String) As Integer

        Dim time As TimeSpan

        If TimeSpan.TryParse(pTime, time) Then

            Dim dateTime As Date = Date.Today.Add(time)
            Dim displayTime = dateTime.ToString("hh:mm tt")
            Dim index = FindString(displayTime)

            If index > -1 Then
                SelectedIndex = index
            End If

            Return index
        Else
            Return -1
        End If

    End Function

    Private _timeSpan As TimeSpan
    ''' <summary>
    ''' Get current selected item as a TimeSpan
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property TimeSpan() As TimeSpan
        Get
            Return Convert.ToDateTime(Text.TrimStart("0"c)).TimeOfDay
        End Get
    End Property

    Private _hours As Integer
    ''' <summary>
    ''' Get hour for selected item
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property SelectedHour() As Integer
        Get
            Return TimeSpan.Hours
        End Get
    End Property

    Private _minutes As Integer
    ''' <summary>
    ''' Get minutes for selected item
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property Minutes() As Integer
        Get
            Return TimeSpan.Minutes
        End Get
    End Property

    ''' <summary>
    ''' Determine if current selected item is AM
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property IsAM() As Boolean
        Get
            Dim check = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, TimeSpan.Hours, TimeSpan.Minutes, 0)

            Return check.ToString("tt") = "AM"

        End Get
    End Property
    ''' <summary>
    ''' Determine if current selected item is PM
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property IsPM() As Boolean
        Get
            Return Not IsAM
        End Get
    End Property

    Protected Shared Property ParentIncrement() As TimeIncrement

    Private _increment As TimeIncrement
    ''' <summary>
    ''' Get/set increment <see cref="TimeIncrement"/>
    ''' </summary>
    <Category("Behavior"), Browsable(True), Description("Time increment")>
    Public Property Increment() As TimeIncrement
        Set
            _increment = Value
            ParentIncrement = Value
            Items.Clear()

            Dim hoursInstance As New Hours()
            Items.AddRange(hoursInstance.Range(Increment))

        End Set
        Get
            Return _increment
        End Get
    End Property

    Private Shared _hour As String = ""

    <Category("Behavior"), Browsable(True), Editor(GetType(Editor), GetType(UITypeEditor)), Description("Hour get/set")>
    Public Property Time() As String
        Set
            _hour = Value
            SetHour()
        End Set
        Get
            Return _hour
        End Get
    End Property
    Private Function SetHour() As Boolean

        Dim success As Boolean = False

        If String.IsNullOrWhiteSpace(Time) Then
            Time = "00:00"
        End If

        Dim result As Integer = FindString(Time)

        If result > -1 Then
            SelectedIndex = result
            Dim item = Items.Count
            success = True
        Else
            SelectedIndex = 0
        End If

        Return success

    End Function
    Private Sub TimeComboBox_VisibleChanged(ByVal sender As Object, ByVal e As EventArgs)
        If Not _shown Then
            SetHour()
            _shown = True
        End If
    End Sub

    ''' <summary>
    ''' Development only
    ''' </summary>
    Public Sub ItemsToTimeSpan()
        For Each timeItem In Items.OfType(Of String)()

            Dim dt As DateTime = Nothing

            If Date.TryParseExact(timeItem, "HH:mm tt", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then

                Dim timeInstance As TimeSpan = dt.TimeOfDay
                Console.WriteLine($"{timeItem} - {timeInstance}")
            End If


        Next

    End Sub

    Friend Class Editor
        Inherits UITypeEditor

        Private _svc As IWindowsFormsEditorService
        Public Overrides Function GetEditStyle(ByVal context As ITypeDescriptorContext) As UITypeEditorEditStyle
            Return UITypeEditorEditStyle.DropDown
        End Function
        Public Overrides Function EditValue(context As ITypeDescriptorContext, provider As IServiceProvider, value As Object) As Object

            _svc = DirectCast(provider.GetService(GetType(IWindowsFormsEditorService)), IWindowsFormsEditorService)

            Dim listBox As New ListBox()

            Dim data() As String = Hours.Range(ParentIncrement)

            For Each item In data
                listBox.Items.Add(item)
            Next item

            If value IsNot Nothing Then
                listBox.SelectedItem = value
            End If

            _svc.DropDownControl(listBox)

            value = DirectCast(listBox.SelectedItem, String)

            Return value

        End Function

    End Class
End Class