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
Imports System.Windows.Media

Namespace Common
	''' <summary>
	''' Interaction logic for MenuButton.xaml
	''' </summary>
	Public Partial Class MenuButton
		Public Shared ReadOnly SourceProperty As DependencyProperty = DependencyProperty.Register("Source", GetType(ImageSource), GetType(MenuButton), New PropertyMetadata(Nothing, Sub(o As DependencyObject, args As DependencyPropertyChangedEventArgs) 
		Dim mButton = TryCast(o, MenuButton)
		If mButton IsNot Nothing Then
			mButton.ImageContent.Source = TryCast(args.NewValue, ImageSource)
		End If

End Sub))

		Public Property Source() As ImageSource
			Get
				Return DirectCast(GetValue(SourceProperty), ImageSource)
			End Get
			Set
				SetValue(SourceProperty, value)
			End Set
		End Property

		Public Sub New()
			InitializeComponent()
		End Sub
	End Class
End Namespace
