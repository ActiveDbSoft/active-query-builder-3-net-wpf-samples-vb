'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Globalization
Imports System.Windows.Data
Imports System.Windows.Media

Namespace Common
	Friend Class ColorEnabledConverter
		Implements IValueConverter
		Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
			Dim enabled = value IsNot Nothing AndAlso CBool(value)

			Return If(enabled, New SolidColorBrush(Color.FromArgb(255, 47, 52, 53)), Brushes.Gray)
		End Function

		Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
			Throw New NotImplementedException()
		End Function
	End Class
End Namespace
