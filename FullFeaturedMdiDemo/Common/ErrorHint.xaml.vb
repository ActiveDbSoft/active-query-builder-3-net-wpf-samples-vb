'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows
Imports System.Windows.Input

Namespace Common
	''' <summary>
	''' Логика взаимодействия для ErrorHint.xaml
	''' </summary>
	Public Partial Class ErrorHint

	    Public Shared ReadOnly TextProperty As DependencyProperty = DependencyProperty.Register("Text", GetType(String), GetType(ErrorHint), new PropertyMetadata(Nothing))


		Public Property Text() As String
			Get
				Return DirectCast(GetValue(TextProperty), String)
			End Get
			Set
				SetValue(TextProperty, value)

			    Visibility = If(String.IsNullOrEmpty(value), Visibility.Collapsed, Visibility.Visible)
			End Set
		End Property

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub CloseImage_OnMouseUp(sender As Object, e As MouseButtonEventArgs)
			Text = String.Empty
		End Sub
	End Class
End Namespace
