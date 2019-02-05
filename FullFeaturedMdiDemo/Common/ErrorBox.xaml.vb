'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel
Imports System.Threading
Imports System.Windows

Namespace Common
    Public Partial Class ErrorBox
        Public Shared ReadOnly MessageProperty As DependencyProperty = DependencyProperty.Register("Message", GetType(String), GetType(ErrorBox), New PropertyMetadata(Nothing))

        Public Property Message() As String
            Get
                Return DirectCast(GetValue(MessageProperty), String)
            End Get
            Set
                SetValue(MessageProperty, value)
            End Set
        End Property
        Public Sub New()
            InitializeComponent()

            Visibility = Visibility.Collapsed

            Dim [property] As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(MessageProperty, GetType(ErrorBox))
            [property].AddValueChanged(Me, AddressOf MessagePropertyChanged)
        End Sub

        Private Sub MessagePropertyChanged(sender As Object, e As EventArgs)
            TextBlockErrorPrompt.Text = Message

            Visibility = If(String.IsNullOrEmpty(Message), Visibility.Collapsed, Visibility.Visible)

            If String.IsNullOrEmpty(Message) Then
                Return
            End If

            Dim timer = New System.Threading.Timer(AddressOf CallBackPopup, Nothing, 3000, Timeout.Infinite)
        End Sub

        Private Sub CallBackPopup(state As Object)
            Dispatcher.BeginInvoke(DirectCast(Sub() 
                If GridError.Visibility = Visibility.Collapsed Then
                    Return
                End If

                TextBlockErrorPrompt.Text = String.Empty
                Visibility = Visibility.Collapsed

            End Sub, Action))
        End Sub
    End Class
End Namespace
