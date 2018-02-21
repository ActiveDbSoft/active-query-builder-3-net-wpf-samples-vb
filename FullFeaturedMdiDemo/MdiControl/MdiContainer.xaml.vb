'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Media

Namespace MdiControl
    ''' <summary>
    ''' Interaction logic for MdiContainer.xaml
    ''' </summary>
    <ContentProperty("Children")> _
    Partial Public Class MdiContainer
        Public Delegate Sub ActiveWindowChangedHandler(sender As Object, args As EventArgs)

        Public Event ActiveWindowChanged As ActiveWindowChangedHandler

        Private Shared _zindex As Integer

        Public Shared ReadOnly PropertyTypeProperty As DependencyProperty = DependencyProperty.Register("Children", GetType(ObservableCollection(Of MdiChildWindow)), GetType(MdiContainer), New PropertyMetadata(Nothing))

        Public Shared ReadOnly FocusedWindowProperty As DependencyProperty = DependencyProperty.Register("FocusedWindow", GetType(MdiChildWindow), GetType(MdiContainer), New PropertyMetadata(Nothing, AddressOf CallbackFocusedWindow))

        Public Shared Shadows ReadOnly BackgroundProperty As DependencyProperty = DependencyProperty.Register("Background", GetType(Brush), GetType(MdiContainer), New PropertyMetadata(SystemColors.ControlBrush, AddressOf CallBackBackground))

        Public Shared ReadOnly ActiveChildProperty As DependencyProperty = DependencyProperty.Register("ActiveChild", GetType(MdiChildWindow), GetType(MdiContainer), New PropertyMetadata(Nothing))

        Public Property ActiveChild() As MdiChildWindow
            Get
                Return DirectCast(GetValue(ActiveChildProperty), MdiChildWindow)
            End Get
            Set(value As MdiChildWindow)
                SetValue(ActiveChildProperty, Value)
                OnActiveWindowChanged()
            End Set
        End Property

        Public Event ActivatedChild As DependencyPropertyChangedEventHandler

        Public Shadows Property Background() As Brush
            Get
                Return DirectCast(GetValue(BackgroundProperty), Brush)
            End Get
            Set(value As Brush)
                SetValue(BackgroundProperty, Value)
            End Set
        End Property

        Private Shared Sub CallBackBackground(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj = TryCast(d, MdiContainer)
            If obj IsNot Nothing Then
                obj.ScrollViewerRoot.Background = DirectCast(e.NewValue, Brush)
            End If
        End Sub

        Private Shared Sub CallbackFocusedWindow(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim oldWindow As MdiChildWindow = TryCast(e.OldValue, MdiChildWindow)
            Dim newWindow As MdiChildWindow = TryCast(e.NewValue, MdiChildWindow)

            If oldWindow IsNot Nothing Then
                oldWindow.IsActive = False
            End If

            If newWindow Is Nothing Then
                Return
            End If

            newWindow.IsActive = True
            Dim cont As MdiContainer = TryCast(d, MdiContainer)
            If cont IsNot Nothing Then
                cont.PromoteChildToFront(newWindow)
                cont.OnActivatedChild(e)
            End If
        End Sub

        Public Sub PromoteChildToFront(child As MdiChildWindow)
            Dim currentZindex As Int32 = Panel.GetZIndex(child)

            For Each localChild As MdiChildWindow In Children
                localChild.IsActive = localChild.Equals(child)

                Dim localZindex As Int32 = Panel.GetZIndex(localChild)

                Panel.SetZIndex(localChild, If(localChild.Equals(child), _zindex - 1, If(localZindex < currentZindex, localZindex, localZindex - 1)))
            Next
            child.Focus()
            ActiveChild = child
        End Sub

        Public Property FocusedWindow() As MdiChildWindow
            Get
                Return DirectCast(GetValue(FocusedWindowProperty), MdiChildWindow)
            End Get
            Set(value As MdiChildWindow)
                SetValue(FocusedWindowProperty, Value)
            End Set
        End Property

        Public Property Children() As ObservableCollection(Of MdiChildWindow)
            Get
                Return DirectCast(GetValue(PropertyTypeProperty), ObservableCollection(Of MdiChildWindow))
            End Get
            Set(value As ObservableCollection(Of MdiChildWindow))
                SetValue(PropertyTypeProperty, Value)
            End Set
        End Property

        Public Sub New()
            Children = New ObservableCollection(Of MdiChildWindow)()
            AddHandler Children.CollectionChanged, AddressOf ChildrenOnCollectionChanged
            InitializeComponent()
            AddHandler Loaded, AddressOf MdiContainer_Loaded
        End Sub

        Public Sub LayoutMdi(mdiLayout1 As MdiLayout)
            Dim point As Point = New Point(0, 0)

            Select Case mdiLayout1
                Case MdiLayout.Cascade
                    For Each mdiChildWindow As ChildWindow In Children
                        mdiChildWindow.State = StateWindow.Normal
                        mdiChildWindow.IsMaximized = False
                        mdiChildWindow.SetLocation(point)

                        point.X += 20
                        point.Y += 20
                    Next
                    Exit Select
                Case MdiLayout.TileHorizontal
                    Dim widthWindow As Double = ScrollViewerRoot.ViewportWidth / Children.Count
                    For Each mdiChildWindow As ChildWindow In Children
                        mdiChildWindow.State = StateWindow.Normal
                        mdiChildWindow.IsMaximized = False
                        mdiChildWindow.SetLocation(point)

                        mdiChildWindow.Width = widthWindow
                        mdiChildWindow.Height = ScrollViewerRoot.ViewportHeight

                        point.X += widthWindow
                    Next
                    Exit Select
                Case MdiLayout.TileVertical
                    Dim windowHeight As Double = ScrollViewerRoot.ViewportHeight / Children.Count
                    For Each mdiChildWindow As ChildWindow In Children
                        mdiChildWindow.State = StateWindow.Normal
                        mdiChildWindow.IsMaximized = False

                        mdiChildWindow.Width = ScrollViewerRoot.ViewportWidth
                        mdiChildWindow.Height = windowHeight

                        mdiChildWindow.SetLocation(point)
                        point.Y += windowHeight
                    Next
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException("mdiLayout", mdiLayout1, Nothing)
            End Select

            PerformLayout()
        End Sub

        Private Sub PerformLayout()
            Dim rect As Rect = New Rect()

            For Each mdiChildWindow As ChildWindow In Children
                rect.Union(mdiChildWindow.Bounds)
            Next

            GridRoot.Width = If(rect.Right > ActualWidth, rect.Right, ScrollViewerRoot.ViewportWidth)

            GridRoot.Height = If(rect.Bottom > ActualHeight, rect.Bottom, ScrollViewerRoot.ViewportHeight)
        End Sub

        Private Sub MdiContainer_Loaded(sender As Object, e As RoutedEventArgs)
            RemoveHandler Loaded, AddressOf MdiContainer_Loaded
            CheckVisibilityScrollbar()
        End Sub

        Private Sub ChildrenOnCollectionChanged(sender As Object, args As NotifyCollectionChangedEventArgs)
            Select Case args.Action
                Case NotifyCollectionChangedAction.Add
                    Dim objAdd As MdiChildWindow = TryCast(args.NewItems(0), MdiChildWindow)
                    If objAdd Is Nothing Then
                        Throw New ArgumentException("Error argument value")
                    End If
                    _zindex += 1

                    AddHandler objAdd.Closing, AddressOf ObjAddClosing
                    AddHandler objAdd.Minimize, AddressOf objAdd_Minimize
                    AddHandler objAdd.Maximize, AddressOf objAdd_Maximize
                    AddHandler objAdd.Resize, AddressOf objAdd_Resize

                    GridRoot.Children.Add(objAdd)
                    PromoteChildToFront(objAdd)
                    If objAdd.State = StateWindow.Maximized Then
                        objAdd.Width = ScrollViewerRoot.ViewportWidth
                    End If

                    objAdd.Height = ScrollViewerRoot.ViewportHeight

                    Exit Select
                Case NotifyCollectionChangedAction.Remove
                    Dim objRemove As MdiChildWindow = TryCast(args.OldItems(0), MdiChildWindow)
                    If objRemove Is Nothing Then
                        Throw New ArgumentException("Error argument value")
                    End If
                    GridRoot.Children.Remove(objRemove)
                    Exit Select
            End Select
        End Sub

        Private Sub CheckVisibilityScrollbar()
            Dim visibleBar As Boolean = Children.Any(Function(x) x.State = StateWindow.Maximized)

            ScrollViewerRoot.VerticalScrollBarVisibility = If(visibleBar, ScrollBarVisibility.Disabled, ScrollBarVisibility.Auto)
            ScrollViewerRoot.HorizontalScrollBarVisibility = If(visibleBar, ScrollBarVisibility.Disabled, ScrollBarVisibility.Auto)
        End Sub

        Private Sub objAdd_Resize(sender As Object, e As EventArgs)
            PerformLayout()
        End Sub

        Private Sub objAdd_Maximize(sender As Object, e As EventArgs)
            Dim mdiChild As MdiChildWindow = DirectCast(sender, MdiChildWindow)

            If mdiChild.State = StateWindow.Maximized Then
                ScrollViewerRoot.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
                ScrollViewerRoot.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled

                Dim width As Double = ScrollViewerRoot.ViewportWidth + (If(GridRoot.ActualWidth > ScrollViewerRoot.ViewportWidth, SystemParameters.VerticalScrollBarWidth, 0))

                Dim height As Double = ScrollViewerRoot.ViewportHeight + (If(GridRoot.ActualHeight > ScrollViewerRoot.ViewportHeight, SystemParameters.HorizontalScrollBarHeight, 0))

                mdiChild.SetLocation(New Point(0, 0))
                mdiChild.SetSize(New Size(width, height))
            Else
                ScrollViewerRoot.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                ScrollViewerRoot.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
            End If
        End Sub

        Private Sub objAdd_Minimize(sender As Object, e As EventArgs)
            Dim mdiChild As MdiChildWindow = DirectCast(sender, MdiChildWindow)

            Dim collectionMinimizedWindow = Children.Where(Function(x) x.State = StateWindow.Minimized AndAlso Math.Abs(x.Height - SystemParameters.MinimizedWindowHeight) < 0.2 AndAlso Not x.Equals(mdiChild)).ToList()

            Dim startY As Double = ScrollViewerRoot.ViewportHeight - SystemParameters.MinimizedWindowHeight

            Dim startX As Double = collectionMinimizedWindow.Sum(Function(mdiChildWindow) mdiChildWindow.ActualWidth + 2)

            If mdiChild.State = StateWindow.Minimized Then
                mdiChild.SetLocation(New Point(startX, startY))
            End If

            If Equals(FocusedWindow, mdiChild) Then
                FocusedWindow = Nothing
            End If
        End Sub

        Private Sub ObjAddClosing(sender As Object, e As EventArgs)
            Dim mdiChild As MdiChildWindow = DirectCast(sender, MdiChildWindow)

            If Equals(ActiveChild, mdiChild) Then
                ActiveChild = Nothing
            End If

            RemoveHandler mdiChild.Closing, AddressOf ObjAddClosing
            RemoveHandler mdiChild.Minimize, AddressOf objAdd_Minimize
            RemoveHandler mdiChild.Maximize, AddressOf objAdd_Maximize
            RemoveHandler mdiChild.Resize, AddressOf objAdd_Resize

            Children.Remove(mdiChild)
        End Sub

        Private Sub GridRoot_OnPreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
            Dim point as Point = e.GetPosition(Me)

            Dim window As MdiChildWindow = Children.FirstOrDefault(Function(mdiChildWindow) mdiChildWindow.IsContainsPoint(point) AndAlso mdiChildWindow.IsMouseOver)

            FocusedWindow = window
        End Sub

        Private Sub ScrollViewerRoot_OnSizeChanged(sender As Object, e As SizeChangedEventArgs)
            Dim mdiChild As MdiChildWindow = Children.FirstOrDefault(Function(child) child.State = StateWindow.Maximized)

            If mdiChild Is Nothing Then
                Return
            End If

            mdiChild.SetSize(New Size(ScrollViewerRoot.ActualWidth, ScrollViewerRoot.ActualHeight))
        End Sub

        Protected Overridable Sub OnActivatedChild(e As DependencyPropertyChangedEventArgs)
            RaiseEvent ActivatedChild(Me, e)
        End Sub

        Protected Overrides Sub OnRenderSizeChanged(sizeInfo As SizeChangedInfo)
            Dim rect As Rect = New Rect()

            For Each mdiChildWindow As MdiChildWindow In Children
                rect.Union(mdiChildWindow.Bounds)
            Next

            GridRoot.Width = If(rect.Right > ActualWidth, rect.Right, ScrollViewerRoot.ViewportWidth)
            GridRoot.Height = If(rect.Bottom > ActualHeight, rect.Bottom, ScrollViewerRoot.ViewportHeight)

            ScrollViewerRoot.VerticalScrollBarVisibility = If(rect.Bottom > ActualHeight, ScrollBarVisibility.Visible, ScrollBarVisibility.Disabled)
            ScrollViewerRoot.HorizontalScrollBarVisibility = If(rect.Right > ActualWidth, ScrollBarVisibility.Visible, ScrollBarVisibility.Disabled)

            MyBase.OnRenderSizeChanged(sizeInfo)

            CheckVisibilityScrollbar()
        End Sub

        Protected Overridable Sub OnActiveWindowChanged()
            RaiseEvent ActiveWindowChanged(Me, EventArgs.Empty)
        End Sub
    End Class

    Public Enum MdiLayout
        Cascade
        TileHorizontal
        TileVertical
    End Enum
End Namespace
