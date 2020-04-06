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
Imports System.Windows.Controls
Imports System.Windows.Media

Namespace MdiControl.Converters
	Public Class MaximizeButtonTemplateSelector
		Inherits DataTemplateSelector

		Public Overrides Function SelectTemplate(item As Object, container As DependencyObject) As DataTemplate
			Dim obj = TryCast(container, FrameworkElement)
			If item IsNot Nothing Then
				Dim state = DirectCast(item, StateWindow)
				Dim mdi = FindParent(Of MdiChildWindow)(obj)

				Dim template As DataTemplate
				If state = StateWindow.Normal Then
					template = TryCast(mdi.FindResource("MaximizeDefaultTemplate"), DataTemplate)
				Else
					template = TryCast(mdi.FindResource("MaximizeTemplate"), DataTemplate)
				End If

				Return template
			End If
			Return Nothing
		End Function

		Private Shared Function FindParent(Of T As Class)(from As DependencyObject) As T
			Dim result As T = Nothing
			Dim parent = VisualTreeHelper.GetParent(from)

			If TypeOf parent Is T Then
				result = TryCast(parent, T)
			ElseIf parent IsNot Nothing Then
				result = FindParent(Of T)(parent)
			End If

			Return result
		End Function
	End Class
End Namespace
