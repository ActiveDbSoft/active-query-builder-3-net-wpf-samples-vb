'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel
Imports System.Threading
Imports System.Windows

Partial Public Class ErrorBox
    Public Shared ReadOnly MessageProperty As DependencyProperty = DependencyProperty.Register("Message", GetType(String), GetType(ErrorBox), New PropertyMetadata(Nothing))
    Public Property Message() As String
        Get
            Return DirectCast(GetValue(MessageProperty), String)
        End Get
        Set
            SetValue(MessageProperty, Value)
        End Set
    End Property
    Public Sub New()
        InitializeComponent()
        Visibility = Visibility.Collapsed
        Dim [property] As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(MessageProperty, GetType(ErrorBox))
        [property].AddValueChanged(Me, MessagePropertyChanged)
    End Sub

    Private Function MessagePropertyChanged() As EventHandler
        TextBlockErrorPrompt.Text = Message
        Visibility = If(String.IsNullOrEmpty(Message), Visibility.Collapsed, Visibility.Visible)
        If String.IsNullOrEmpty(Message) Then
            Return Nothing
        End If
        Dim timer As Timer = New Timer(CallBackPopup, Nothing, 3000, Timeout.Infinite)
    End Function

    Private Function CallBackPopup() As TimerCallback
        Dispatcher.BeginInvoke(DirectCast(Reset, Action))
    End Function

    Private Function Reset()
        If Equals(Visibility, Visibility.Visible)
            TextBlockErrorPrompt.Text = String.Empty
                Visibility = Visibility.Collapsed
        End If
    End Function
End Class
