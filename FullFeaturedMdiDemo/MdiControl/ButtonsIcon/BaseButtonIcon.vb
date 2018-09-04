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
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media

Namespace MdiControl.ButtonsIcon
	Public MustInherit Class BaseButtonIcon
		Inherits Grid
		Implements IBaseButtonIcon
		Public Event Click As RoutedEventHandler

		Public Shared Shadows ReadOnly BackgroundProperty As DependencyProperty = DependencyProperty.Register("Background", GetType(Brush), GetType(BaseButtonIcon), New PropertyMetadata(Brushes.Transparent))

		Public Shadows Property Background() As Brush
			Get
				Return DirectCast(GetValue(BackgroundProperty), Brush)
			End Get
			Set
				SetValue(BackgroundProperty, value)
			End Set
		End Property

		Public Shared ReadOnly StrokeProperty As DependencyProperty = DependencyProperty.Register("Stroke", GetType(Brush), GetType(BaseButtonIcon), New FrameworkPropertyMetadata(Brushes.Black, FrameworkPropertyMetadataOptions.AffectsRender))

		Public Shared ReadOnly IsMaximizedProperty As DependencyProperty = DependencyProperty.Register("IsMaximized", GetType(Boolean), GetType(BaseButtonIcon), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.AffectsRender, AddressOf CallbackIsMaximized))

		Private Shared Sub CallbackIsMaximized(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim sender As BaseButtonIcon = TryCast(d, BaseButtonIcon)
			Dim value = CBool(e.NewValue)
			If sender IsNot Nothing Then
				sender.SizeContent = If(value, New Size(8, 8), New Size(11, 9))
			End If
		End Sub

		Protected Property SizeContent() As Size
			Get
				Return m_SizeContent
			End Get
			Set
				m_SizeContent = Value
			End Set
		End Property
		Private m_SizeContent As Size

		Public Property Stroke() As Brush Implements IBaseButtonIcon.Stroke
			Get
				Return DirectCast(GetValue(StrokeProperty), Brush)
			End Get
			Set
				SetValue(StrokeProperty, value)
			End Set
		End Property

		Public Property IsMaximized() As Boolean Implements IBaseButtonIcon.IsMaximized
			Get
				Return CBool(GetValue(IsMaximizedProperty))
			End Get
			Set
				SetValue(IsMaximizedProperty, value)
			End Set
		End Property

		Protected Sub New()
			SnapsToDevicePixels = True
			SetValue(RenderOptions.EdgeModeProperty, EdgeMode.Aliased)


			SizeContent = New Size(11, 9)
		End Sub

		Protected Overrides Sub OnPreviewMouseLeftButtonUp(e As MouseButtonEventArgs)
			InvalidateVisual()
			OnClick(e)
		End Sub


		Protected Overrides Sub OnPreviewMouseLeftButtonDown(e As MouseButtonEventArgs)
			InvalidateVisual()
		End Sub

		Protected Overrides Sub OnPreviewMouseDown(e As MouseButtonEventArgs)
			InvalidateVisual()
			'base.OnPreviewMouseDown(e);
		End Sub
		Protected Overridable Sub OnClick(e As RoutedEventArgs)
            RaiseEvent Click(Me, e)
		End Sub
	End Class
End Namespace
