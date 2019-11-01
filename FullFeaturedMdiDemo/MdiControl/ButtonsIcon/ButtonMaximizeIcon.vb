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
Imports System.Windows.Media

Namespace MdiControl.ButtonsIcon
	Public Class ButtonMaximizeIcon
		Inherits BaseButtonIcon
		Protected Overrides Sub OnRender(drawingContext As DrawingContext)
			MinHeight = SizeContent.Height
			MinWidth = SizeContent.Width

            Dim x As Double = (ActualWidth - SizeContent.Width) / 2
            Dim y As Double = (ActualHeight - SizeContent.Height) / 2

            Dim pen As Pen = New Pen(Stroke, 1)

            Dim brushPressed As SolidColorBrush = New SolidColorBrush(CType(ColorConverter.ConvertFromString("#3d6099"), Color))

            Dim pressed As Boolean = IsMouseOver AndAlso Mouse.LeftButton = MouseButtonState.Pressed

            Dim brush As Brush = If(pressed, brushPressed, (If(IsMouseOver, Background, Brushes.Transparent)))

            drawingContext.DrawRectangle(brush, Nothing, New Rect(New Point(0, 0), New Size(ActualWidth, ActualHeight)))

			If Not IsMaximized Then
				drawingContext.DrawRectangle(Nothing, pen, New Rect(New Point(x, y), SizeContent))
				drawingContext.DrawLine(pen, New Point(x, y + 1), New Point(x + SizeContent.Width, y + 1))
				Return
			End If

			y += 2
			drawingContext.DrawRectangle(Nothing, pen, New Rect(New Point(x, y), SizeContent))
			drawingContext.DrawLine(pen, New Point(x, y + 1), New Point(x + SizeContent.Width, y + 1))

			drawingContext.DrawLine(pen, New Point(x + 2, y - 2), New Point(x + SizeContent.Width + 2, y - 2))
			drawingContext.DrawLine(pen, New Point(x + SizeContent.Width + 2, y - 2), New Point(x + SizeContent.Width + 2, SizeContent.Height + y - 2))
		End Sub
	End Class
End Namespace
