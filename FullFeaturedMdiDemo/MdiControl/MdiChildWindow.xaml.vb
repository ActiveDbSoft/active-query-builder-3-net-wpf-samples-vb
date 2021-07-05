//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System.Windows.Controls.Primitives
Imports System.Windows.Markup

Namespace MdiControl
	''' <summary>
	''' Interaction logic for MdiChildWindow.xaml
	''' </summary>
	<ContentProperty("Children")>
	Partial Public Class MdiChildWindow
		Private Const WidthMinimized As Double = 173
		Private _oldSize As Windows.Size = Windows.Size.Empty
		Private _oldPoint As New Windows.Point(0, 0)
		Private Shadows Property Content() As Grid

		Public Event Closing As EventHandler
		Public Event Minimize As EventHandler
		Public Event Maximize As EventHandler
		Public Event Resize As EventHandler

#Region "Dependency"
		Public Shared ReadOnly LeftProperty As DependencyProperty = DependencyProperty.Register("Left", GetType(Double), GetType(MdiChildWindow), New PropertyMetadata(0.0, AddressOf CallBackLeft))

		Public Shared ReadOnly TopProperty As DependencyProperty = DependencyProperty.Register("Top", GetType(Double), GetType(MdiChildWindow), New PropertyMetadata(0.0, AddressOf CallBackTop))

		Public Shared ReadOnly IsActiveProperty As DependencyProperty = DependencyProperty.Register("IsActive", GetType(Boolean), GetType(MdiChildWindow), New PropertyMetadata(False))

		Public Shared ReadOnly TitleProperty As DependencyProperty = DependencyProperty.Register("Title", GetType(String), GetType(MdiChildWindow), New PropertyMetadata("Window"))

		Public Shared Shadows ReadOnly BackgroundProperty As DependencyProperty = DependencyProperty.Register("Background", GetType(Windows.Media.Brush), GetType(MdiChildWindow), New PropertyMetadata(Windows.Media.Brushes.White, AddressOf CallBackBackgound))

		Public Shared ReadOnly StateProperty As DependencyProperty = DependencyProperty.Register("State", GetType(StateWindow), GetType(MdiChildWindow), New PropertyMetadata(StateWindow.Normal, AddressOf CallbackStateChange))

		Private Shared Sub CallbackStateChange(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
'INSTANT VB NOTE: The variable state was renamed since Visual Basic does not handle local variables named the same as class members well:
			Dim state_Renamed = DirectCast(e.NewValue, StateWindow)

			Dim obj = TryCast(d, MdiChildWindow)
			If obj IsNot Nothing Then
				Select Case state_Renamed
					Case StateWindow.Normal
						obj.SetSize(If(obj._oldSize = Windows.Size.Empty, New Windows.Size(300, 200), obj._oldSize))
						obj.SetLocation(obj._oldPoint)
					Case StateWindow.Minimized
						obj.State = StateWindow.Minimized
						obj._oldSize = New Windows.Size(obj.ActualWidth, obj.ActualHeight)
						obj._oldPoint = New Windows.Point(obj.Left, obj.Top)
						obj.Height = SystemParameters.MinimizedWindowHeight
						obj.Width = WidthMinimized
					Case StateWindow.Maximized
						obj.IsMaximized = True
						obj.Measure(New Windows.Size(Double.PositiveInfinity, Double.PositiveInfinity))
						obj.Arrange(New Rect(obj.DesiredSize))
					Case Else
						Throw New ArgumentOutOfRangeException()
				End Select

			End If
		End Sub

		Public Shared ReadOnly IsMaximizedProperty As DependencyProperty = DependencyProperty.Register("IsMaximized", GetType(Boolean), GetType(MdiChildWindow), New PropertyMetadata(False))
#End Region

		Public Property IsMaximized() As Boolean
			Get
				Return DirectCast(GetValue(IsMaximizedProperty), Boolean)
			End Get
			Set(value As Boolean)
				SetValue(IsMaximizedProperty, value)
			End Set
		End Property

		Public Property State() As StateWindow
			Get
				Return DirectCast(GetValue(StateProperty), StateWindow)
			End Get
			Set(value As StateWindow)
				SetValue(StateProperty, value)
			End Set
		End Property

		Public Shadows Property Background() As Windows.Media.Brush
			Get
				Return DirectCast(GetValue(BackgroundProperty), Windows.Media.Brush)
			End Get
			Set(value As Windows.Media.Brush)
				SetValue(BackgroundProperty, value)
			End Set
		End Property

		Public Property Title() As String
			Get
				Return DirectCast(GetValue(TitleProperty), String)
			End Get
			Set(value As String)
				SetValue(TitleProperty, value)
			End Set
		End Property

		Public Property Children() As ObservableCollection(Of Visual)

		Public Property IsActive() As Boolean
			Get
				Return DirectCast(GetValue(IsActiveProperty), Boolean)
			End Get
			Set(value As Boolean)
				SetValue(IsActiveProperty, value)
			End Set
		End Property

		Public Property Top() As Double
			Get
				Return DirectCast(GetValue(TopProperty), Double)
			End Get
			Set(value As Double)
				SetValue(TopProperty, value)
			End Set
		End Property

		Public Property Left() As Double
			Get
				Return DirectCast(GetValue(LeftProperty), Double)
			End Get
			Set(value As Double)
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

		Public Function IsContainsPoint(point As Windows.Point) As Boolean
			'INSTANT VB NOTE: The variable left was renamed since Visual Basic does not handle local variables named the same as class members well:
			Dim left_Renamed = Canvas.GetLeft(Me)
			'INSTANT VB NOTE: The variable top was renamed since Visual Basic does not handle local variables named the same as class members well:
			Dim top_Renamed = Canvas.GetTop(Me)

			If Double.IsNaN(left_Renamed) Then
				left_Renamed = TranslatePoint(New Windows.Point(0, 0), CType(Parent, UIElement)).X
			End If

			If Double.IsNaN(top_Renamed) Then
				top_Renamed = TranslatePoint(New Windows.Point(0, 0), CType(Parent, UIElement)).Y
			End If

			Dim rect = New Rect(left_Renamed, top_Renamed, ActualWidth, ActualHeight)
			Return rect.Contains(point)
		End Function

		Public Sub SetLocation(point As Windows.Point)
			Left = point.X
			Top = point.Y
		End Sub

		Public Sub SetSize(size As Windows.Size)
			Width = size.Width
			Height = size.Height
		End Sub

		Private Shared Sub CallBackLeft(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
			Dim obj = TryCast(d, MdiChildWindow)

			If obj IsNot Nothing Then
				obj.SetValue(Canvas.LeftProperty, DirectCast(e.NewValue, Double))
			End If
		End Sub

		Private Shared Sub CallBackTop(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
			Dim obj = TryCast(d, MdiChildWindow)

			If obj IsNot Nothing Then
				obj.SetValue(Canvas.TopProperty, DirectCast(e.NewValue, Double))
			End If
		End Sub

		Private Shared Sub CallBackBackgound(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
			Dim obj = TryCast(d, MdiChildWindow)
			If obj IsNot Nothing Then
				If obj.Content Is Nothing Then
					obj.ApplyTemplate()
				End If
				obj.Content.Background = DirectCast(e.NewValue, Windows.Media.Brush)
			End If
		End Sub

		Private Sub ChildrenOnCollectionChanged(sender As Object, args As NotifyCollectionChangedEventArgs)
			If Content Is Nothing Then
				ApplyTemplate()
			End If

			Select Case args.Action
				Case NotifyCollectionChangedAction.Add
					Content.Children.Add(DirectCast(args.NewItems(0), FrameworkElement))
				Case NotifyCollectionChangedAction.Remove
					Content.Children.Remove(DirectCast(args.NewItems(0), FrameworkElement))
			End Select
		End Sub

		Protected Overridable Sub OnClose()
			Dim handler = ClosingEvent
			If handler IsNot Nothing Then
				handler(Me, EventArgs.Empty)
			End If
		End Sub

		Protected Overridable Sub OnMinimize()
			Dim handler = MinimizeEvent
			If handler IsNot Nothing Then
				handler(Me, EventArgs.Empty)
			End If
		End Sub

		Protected Overridable Sub OnMaximize()
			Dim handler = MaximizeEvent
			If handler IsNot Nothing Then
				handler(Me, EventArgs.Empty)
			End If
		End Sub

		Protected Overridable Sub OnResize()
			Dim handler = ResizeEvent
			If handler IsNot Nothing Then
				handler(Me, EventArgs.Empty)
			End If
		End Sub

		Private Sub ThumbHeight_OnDragDelta(sender As Object, e As DragDeltaEventArgs)
			If IsMaximized OrElse State = StateWindow.Minimized Then
				Return
			End If
			If ActualHeight + e.VerticalChange < 30 Then
				Return
			End If
			Dim newSize = New Windows.Size(ActualWidth, ActualHeight + e.VerticalChange)

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
			Dim newSize = New Windows.Size(ActualWidth + e.HorizontalChange, ActualHeight)

			Width = newSize.Width
			OnResize()
		End Sub

		Private Sub ThumbHW_OnDragDelta(sender As Object, e As DragDeltaEventArgs)
			If IsMaximized OrElse State = StateWindow.Minimized Then
				Return
			End If
			Dim nWidth = ActualWidth
			Dim nHeight = ActualHeight

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
				SetSize(If(_oldSize = Windows.Size.Empty, New Windows.Size(300, 200), _oldSize))
				SetLocation(_oldPoint)
			Else
				IsMaximized = True

				If State <> StateWindow.Minimized Then
					_oldSize = New Windows.Size(ActualWidth, ActualHeight)
					_oldPoint = New Windows.Point(Left, Top)
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
				_oldSize = New Windows.Size(ActualWidth, ActualHeight)
				_oldPoint = New Windows.Point(Left, Top)
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
