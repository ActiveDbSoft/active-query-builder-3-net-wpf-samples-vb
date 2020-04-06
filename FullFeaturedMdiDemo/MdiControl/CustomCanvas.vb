'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls

Namespace MdiControl
	Public Class CustomCanvas
		Inherits Canvas

		Protected Overrides Function MeasureOverride(constraint As Size) As Size
			Dim defaultValue = MyBase.MeasureOverride(constraint)
'INSTANT VB NOTE: The variable desiredSize was renamed since Visual Basic does not handle local variables named the same as class members well:
			Dim desiredSize_Renamed = New Size()

			desiredSize_Renamed = Children.Cast(Of UIElement)().Aggregate(desiredSize_Renamed, Function(current, child) New Size(Math.Max(current.Width, GetLeft(child) + child.DesiredSize.Width), Math.Max(current.Height, GetTop(child) + child.DesiredSize.Height)))

			Return If(Double.IsNaN(desiredSize_Renamed.Width) OrElse Double.IsNaN(desiredSize_Renamed.Height), defaultValue, desiredSize_Renamed)
		End Function
	End Class
End Namespace
