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
Imports System.Windows.Input
Imports System.Windows.Media

Namespace MdiControl.ButtonsIcon
	Public Class ButtonMinimizeIcon
		Inherits BaseButtonIcon


		Protected Overrides Sub OnRender(drawingContext As DrawingContext)
			MinHeight = SizeContent.Height
			MinWidth = SizeContent.Width

			Dim x = (ActualWidth - SizeContent.Width) / 2
			Dim y = (ActualHeight - SizeContent.Height) / 2

			Dim pen = New Pen(Stroke, 2)

			Dim brushPressed = New SolidColorBrush(CType(ColorConverter.ConvertFromString("#3d6099"), Color))

			Dim pressed = IsMouseOver AndAlso Mouse.LeftButton = MouseButtonState.Pressed

			Dim brush = If(pressed, brushPressed, (If(IsMouseOver, Background, Brushes.Transparent)))

			drawingContext.DrawRectangle(brush, Nothing, New Rect(New Point(0, 0), New Size(ActualWidth, ActualHeight)))



			drawingContext.DrawLine(pen, New Point(x, y + SizeContent.Height - 1), New Point(x + SizeContent.Width, y + SizeContent.Height - 1))
		End Sub
	End Class
End Namespace
