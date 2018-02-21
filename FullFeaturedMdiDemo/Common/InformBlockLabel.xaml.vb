'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows

Namespace Common
	''' <summary>
	''' Логика взаимодействия для InformBlockLabel.xaml
	''' </summary>
	Public Partial Class InformBlockLabel
		Public Shared ReadOnly TextProperty As DependencyProperty = DependencyProperty.Register("Text", GetType(String), GetType(InformBlockLabel), New PropertyMetadata(Nothing))

		Public Property Text() As String
			Get
				Return DirectCast(GetValue(TextProperty), String)
			End Get
			Set
				SetValue(TextProperty, value)
			End Set
		End Property

		Public Sub New()
			InitializeComponent()
		End Sub
	End Class
End Namespace
