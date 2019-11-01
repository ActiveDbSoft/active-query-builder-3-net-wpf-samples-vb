'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls

Namespace MdiControl
	Public Class CustomCanvas
		Inherits Canvas
		Protected Overrides Function MeasureOverride(constraint As Size) As Size
            Dim defaultValue As Size = MyBase.MeasureOverride(constraint)
            Dim desiredSize As Size = New Size()

            desiredSize = Children.Cast(Of UIElement)().Aggregate(desiredSize, Function(current, child) New Size(Math.Max(current.Width, GetLeft(child) + child.DesiredSize.Width), Math.Max(current.Height, GetTop(child) + child.DesiredSize.Height)))

			Return If((Double.IsNaN(desiredSize.Width) OrElse Double.IsNaN(desiredSize.Height)), defaultValue, desiredSize)
		End Function
	End Class
End Namespace
