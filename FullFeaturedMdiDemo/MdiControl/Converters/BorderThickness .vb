''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Namespace MdiControl.Converters
	Friend Class BorderThickness
		Implements IValueConverter

		Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
			If parameter IsNot Nothing Then
				Dim h = SystemParameters.ResizeFrameVerticalBorderWidth

				Return New Thickness(h, 0, h, 0)
			End If

			Dim w = SystemParameters.ResizeFrameVerticalBorderWidth

			Return New Thickness(w, 0, w, w)
		End Function

		Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
			Throw New NotImplementedException()
		End Function
	End Class
End Namespace
