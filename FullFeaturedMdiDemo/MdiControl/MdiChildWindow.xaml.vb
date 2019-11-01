'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Media

Namespace MdiControl
	''' <summary>
	''' Interaction logic for MdiChildWindow.xaml
	''' </summary>
	<ContentProperty("Children")> _
	Public Partial Class MdiChildWindow
		Private Const WidthMinimized As Double = 173
		Private _oldSize As Size = Size.Empty
		Private _oldPoint As New Point(0, 0)
		Private Shadows Property Content() As Grid
			Get
				Return m_Content
			End Get
			Set
				m_Content = Value
			End Set
		End Property
		Private Shadows m_Content As Grid

		Public Event Closing As EventHandler
		Public Event Minimize As EventHandler
		Public Event Maximize As EventHandler
		Public Event Resize As EventHandler

		#Region "Dependency"
		Public Shared ReadOnly LeftProperty As DependencyProperty = DependencyProperty.Register("Left", GetType(Double), GetType(MdiChildWindow), New PropertyMetadata(0.0, AddressOf CallBackLeft))

		Public Shared ReadOnly TopProperty As DependencyProperty = DependencyProperty.Register("Top", GetType(Double), GetType(MdiChildWindow), New PropertyMetadata(0.0, AddressOf CallBackTop))

		Public Shared ReadOnly IsActiveProperty As DependencyProperty = DependencyProperty.Register("IsActive", GetType(Boolean), GetType(MdiChildWindow), New PropertyMetadata(False))

		Public Shared ReadOnly TitleProperty As DependencyProperty = DependencyProperty.Register("Title", GetType(String), GetType(MdiChildWindow), New PropertyMetadata("Window"))

		Public Shared Shadows ReadOnly BackgroundProperty As DependencyProperty = DependencyProperty.Register("Background", GetType(Brush), GetType(MdiChildWindow), New PropertyMetadata(Brushes.White, AddressOf CallBackBackgound))

		Public Shared ReadOnly StateProperty As DependencyProperty = DependencyProperty.Register("State", GetType(StateWindow), GetType(MdiChildWindow), New PropertyMetadata(StateWindow.Normal, AddressOf CallbackStateChange))

		Private Shared Sub CallbackStateChange(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
			If CType(e.NewValue, StateWindow) <> StateWindow.Maximized Then
				Return
			End If

            Dim obj As MdiChildWindow = TryCast(d, MdiChildWindow)
            If obj IsNot Nothing Then
				obj.IsMaximized = True
				obj.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
				obj.Arrange(New Rect(obj.DesiredSize))
			End If
		End Sub

		Public Shared ReadOnly IsMaximizedProperty As DependencyProperty = DependencyProperty.Register("IsMaximized", GetType(Boolean), GetType(MdiChildWindow), New PropertyMetadata(False))
		#End Region

		Public Property IsMaximized() As Boolean
			Get
				Return CBool(GetValue(IsMaximizedProperty))
			End Get
			Set
				SetValue(IsMaximizedProperty, value)
			End Set
		End Property

		Public Property State() As StateWindow
			Get
				Return CType(GetValue(StateProperty), StateWindow)
			End Get
			Set
				SetValue(StateProperty, value)
			End Set
		End Property

		Public Shadows Property Background() As Brush
			Get
				Return DirectCast(GetValue(BackgroundProperty), Brush)
			End Get
			Set
				SetValue(BackgroundProperty, value)
			End Set
		End Property

		Public Property Title() As String
			Get
				Return DirectCast(GetValue(TitleProperty), String)
			End Get
			Set
				SetValue(TitleProperty, value)
			End Set
		End Property

		Public Property Children() As ObservableCollection(Of Visual)
			Get
				Return m_Children
			End Get
			Set
				m_Children = Value
			End Set
		End Property
		Private m_Children As ObservableCollection(Of Visual)

		Public Property IsActive() As Boolean
			Get
				Return CBool(GetValue(IsActiveProperty))
			End Get
			Set
				SetValue(IsActiveProperty, value)
			End Set
		End Property

		Public Property Top() As Double
			Get
				Return CDbl(GetValue(TopProperty))
			End Get
			Set
				SetValue(TopProperty, value)
			End Set
		End Property

		Public Property Left() As Double
			Get
				Return CDbl(GetValue(LeftProperty))
			End Get
			Set
				SetValue(LeftProperty, value)
			End Set
		End Property

		Public Overridable Sub Close()
			OnClose()
		End Sub

		Public ReadOnly Property Bounds() As Rect
			Get
				Return New Rect(Left, Top, ActualWidth, ActualHeight)
			End Get
		End Property

		Public Sub New()
			Children = New ObservableCollection(Of Visual)()
			AddHandler Children.CollectionChanged, AddressOf ChildrenOnCollectionChanged
			InitializeComponent()

			Canvas.SetLeft(Me, 0)
			Canvas.SetTop(Me, 0)

			DataContext = Me

			State = StateWindow.Normal
		End Sub

		Public Overrides Sub OnApplyTemplate()
			MyBase.OnApplyTemplate()

			Content = DirectCast(Template.FindName("GridRoot", Me), Grid)
		End Sub

		Public Function IsContainsPoint(point As Point) As Boolean
            Dim left As Double = Canvas.GetLeft(Me)
            Dim top As Double = Canvas.GetTop(Me)

			If Double.IsNaN(left) Then
				left = TranslatePoint(New Point(0, 0), DirectCast(Parent, UIElement)).X
			End If

			If Double.IsNaN(top) Then
				top = TranslatePoint(New Point(0, 0), DirectCast(Parent, UIElement)).Y
			End If

            Dim rect As Rect = New Rect(left, top, ActualWidth, ActualHeight)
			Return rect.Contains(point)
		End Function

		Public Sub SetLocation(point As Point)
			Left = point.X
			Top = point.Y
		End Sub

		Public Sub SetSize(size As Size)
			Width = size.Width
			Height = size.Height
		End Sub

		Private Shared Sub CallBackLeft(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj As MdiChildWindow = TryCast(d, MdiChildWindow)

			If obj IsNot Nothing Then
				obj.SetValue(Canvas.LeftProperty, CDbl(e.NewValue))
			End If
		End Sub

		Private Shared Sub CallBackTop(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj As MdiChildWindow = TryCast(d, MdiChildWindow)

			If obj IsNot Nothing Then
				obj.SetValue(Canvas.TopProperty, CDbl(e.NewValue))
			End If
		End Sub

		Private Shared Sub CallBackBackgound(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj As MdiChildWindow = TryCast(d, MdiChildWindow)
			If obj IsNot Nothing Then
				If obj.Content Is Nothing Then
					obj.ApplyTemplate()
				End If
				obj.Content.Background = DirectCast(e.NewValue, Brush)
			End If
		End Sub

		Private Sub ChildrenOnCollectionChanged(sender As Object, args As NotifyCollectionChangedEventArgs)
			If Content Is Nothing Then
				ApplyTemplate()
			End If

			Select Case args.Action
				Case NotifyCollectionChangedAction.Add
					Content.Children.Add(DirectCast(args.NewItems(0), FrameworkElement))
					Exit Select
				Case NotifyCollectionChangedAction.Remove
					Content.Children.Remove(DirectCast(args.NewItems(0), FrameworkElement))
					Exit Select
			End Select
		End Sub

		Protected Overridable Sub OnClose()
            RaiseEvent Closing(Me, EventArgs.Empty)
		End Sub

		Protected Overridable Sub OnMinimize()
            RaiseEvent Minimize(Me, EventArgs.Empty)
		End Sub

		Protected Overridable Sub OnMaximize()
            RaiseEvent Maximize(Me, EventArgs.Empty)
		End Sub

		Protected Overridable Sub OnResize()
            RaiseEvent Resize(Me, EventArgs.Empty)
		End Sub

		Private Sub ThumbHeight_OnDragDelta(sender As Object, e As DragDeltaEventArgs)
			If IsMaximized OrElse State = StateWindow.Minimized Then
				Return
			End If
			If ActualHeight + e.VerticalChange < 30 Then
				Return
			End If
            Dim newSize As Size = New Size(ActualWidth, ActualHeight + e.VerticalChange)

			Height = newSize.Height

			OnResize()
		End Sub

		Private Sub ThumbWidth_OnDragDelta(sender As Object, e As DragDeltaEventArgs)
			If IsMaximized OrElse State = StateWindow.Minimized Then
				Return
			End If

			If ActualWidth + e.HorizontalChange < 140 Then
				Return
			End If
            Dim newSize As Size = New Size(ActualWidth + e.HorizontalChange, ActualHeight)

			Width = newSize.Width
			OnResize()
		End Sub

		Private Sub ThumbHW_OnDragDelta(sender As Object, e As DragDeltaEventArgs)
			If IsMaximized OrElse State = StateWindow.Minimized Then
				Return
			End If
            Dim nWidth As Double = ActualWidth
            Dim nHeight As Double = ActualHeight

			If ActualHeight + e.VerticalChange > 30 Then
				nHeight += e.VerticalChange
			End If

			If ActualWidth + e.HorizontalChange > 140 Then
				nWidth += e.HorizontalChange
			End If

			Width = nWidth
			Height = nHeight
			OnResize()
		End Sub

		Private Sub ThumbMove_OnDragDelta(sender As Object, e As DragDeltaEventArgs)
			If IsMaximized OrElse State = StateWindow.Minimized Then
				Return
			End If

			Top += e.VerticalChange
			Left += e.HorizontalChange
			OnResize()
		End Sub

		Private Sub ButtonClose_OnClick(sender As Object, e As RoutedEventArgs)
			Close()
		End Sub

		Private Sub ButtonMaximize_OnClick(sender As Object, e As RoutedEventArgs)
			If State = StateWindow.Maximized Then
				IsMaximized = False
				State = StateWindow.Normal
				SetSize(If(_oldSize = Size.Empty, New Size(300, 200), _oldSize))
				SetLocation(_oldPoint)
			Else
				IsMaximized = True

				If State <> StateWindow.Minimized Then
					_oldSize = New Size(ActualWidth, ActualHeight)
					_oldPoint = New Point(Left, Top)
				End If
				State = StateWindow.Maximized
			End If

			OnMaximize()
		End Sub

		Private Sub ButtonMinimize_OnClick(sender As Object, e As RoutedEventArgs)
			If State = StateWindow.Minimized Then
				State = If(IsMaximized, StateWindow.Maximized, StateWindow.Normal)
				Width = _oldSize.Width
				Height = _oldSize.Height
				Left = _oldPoint.X
				Top = _oldPoint.Y
			Else
				State = StateWindow.Minimized
				_oldSize = New Size(ActualWidth, ActualHeight)
				_oldPoint = New Point(Left, Top)
				Height = SystemParameters.MinimizedWindowHeight
				Width = WidthMinimized
			End If

			OnMinimize()
		End Sub

		Private Sub Header_OnMouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
			ButtonMaximize_OnClick(sender, e)
		End Sub
	End Class

	Public Enum StateWindow
		Normal
		Minimized
		Maximized
	End Enum
End Namespace
