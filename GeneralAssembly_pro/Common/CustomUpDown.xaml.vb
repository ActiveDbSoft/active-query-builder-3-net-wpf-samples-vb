''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Windows
Imports System.Windows.Input

Namespace Common
    ''' <summary>
    ''' Interaction logic for CustomUpDown.xaml
    ''' </summary>
    Partial Public Class CustomUpDown
        Public Delegate Sub ValueChangedHandler(sender As Object, e As EventArgs)

        Public Event ValueChanged As ValueChangedHandler

        Public Property Value() As Integer
            Set(value As Integer)
                TextBoxValue.Text = value.ToString()
                OnValueChanged()
            End Set
            Get
                Return Integer.Parse(TextBoxValue.Text)
            End Get
        End Property

        Public Sub New()
            InitializeComponent()
            Value = 0
        End Sub

        Private Sub ValueUpButton_OnClick(sender As Object, e As RoutedEventArgs)
            Value += 1
        End Sub

        Private Sub ValueDownButton_OnClick(sender As Object, e As RoutedEventArgs)
            Value -= 1
        End Sub

        Private Sub TextBox1_OnPreviewTextInput(sender As Object, e As TextCompositionEventArgs)
            e.Handled = Not TextBoxTextAllowed(e.Text)

            Dim localValue As Integer = Nothing

            Integer.TryParse(TextBoxValue.Text, localValue)
            Value = localValue
        End Sub

        Private Sub textBoxValue_Pasting(sender As Object, e As DataObjectPastingEventArgs)
            If e.DataObject.GetDataPresent(GetType(String)) Then
                Dim text1 = DirectCast(e.DataObject.GetData(GetType(String)), String)
                If Not TextBoxTextAllowed(text1) Then
                    e.CancelCommand()
                End If
            Else
                e.CancelCommand()
            End If
        End Sub

        Private Shared Function TextBoxTextAllowed(text2 As String) As Boolean
            Return Array.TrueForAll(text2.ToCharArray(), Function(c) Char.IsDigit(c) OrElse Char.IsControl(c))
        End Function

        Protected Overridable Sub OnValueChanged()
            Dim handler = ValueChangedEvent
            If handler IsNot Nothing Then
                handler(Me, EventArgs.Empty)
            End If
        End Sub
    End Class
End Namespace
