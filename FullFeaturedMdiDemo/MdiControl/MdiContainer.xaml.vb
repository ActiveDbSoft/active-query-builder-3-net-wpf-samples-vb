''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Windows.Markup

Namespace MdiControl
    ''' <summary>
    ''' Interaction logic for MdiContainer.xaml
    ''' </summary>
    <ContentProperty("Children")>
    Partial Public Class MdiContainer
        Public Delegate Sub ActiveWindowChangedHandler(sender As Object, args As EventArgs)

        Public Event ActiveWindowChanged As ActiveWindowChangedHandler

        Private Shared _zindex As Integer

        Public Shared ReadOnly PropertyTypeProperty As DependencyProperty = DependencyProperty.Register("Children", GetType(ObservableCollection(Of MdiChildWindow)), GetType(MdiContainer), New PropertyMetadata(Nothing))

        Public Shared ReadOnly FocusedWindowProperty As DependencyProperty = DependencyProperty.Register("FocusedWindow", GetType(MdiChildWindow), GetType(MdiContainer), New PropertyMetadata(Nothing, AddressOf CallbackFocusedWindow))

        Public Shared Shadows ReadOnly BackgroundProperty As DependencyProperty = DependencyProperty.Register("Background", GetType(Windows.Media.Brush), GetType(MdiContainer), New PropertyMetadata(Windows.SystemColors.ControlBrush, AddressOf CallBackBackground))

        Public Shared ReadOnly ActiveChildProperty As DependencyProperty = DependencyProperty.Register("ActiveChild", GetType(MdiChildWindow), GetType(MdiContainer), New PropertyMetadata(Nothing))

        Public Property ActiveChild() As MdiChildWindow
            Get
                Return DirectCast(GetValue(ActiveChildProperty), MdiChildWindow)
            End Get
            Set(value As MdiChildWindow)
                SetValue(ActiveChildProperty, value)
                FocusedWindow = value
                OnActiveWindowChanged()
            End Set
        End Property

        Public Event ActivatedChild As DependencyPropertyChangedEventHandler

        Public Shadows Property Background() As Windows.Media.Brush
            Get
                Return DirectCast(GetValue(BackgroundProperty), Windows.Media.Brush)
            End Get
            Set(value As Windows.Media.Brush)
                SetValue(BackgroundProperty, value)
            End Set
        End Property

        Private Shared Sub CallBackBackground(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim obj = TryCast(d, MdiContainer)
            If obj IsNot Nothing Then
                obj.ScrollViewerRoot.Background = DirectCast(e.NewValue, Windows.Media.Brush)
            End If
        End Sub

        Private Shared Sub CallbackFocusedWindow(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim oldWindow = TryCast(e.OldValue, MdiChildWindow)
            Dim newWindow = TryCast(e.NewValue, MdiChildWindow)

            If oldWindow IsNot Nothing Then
                oldWindow.IsActive = False
            End If

            If newWindow Is Nothing Then
                Return
            End If

            newWindow.IsActive = True

            If newWindow.State = StateWindow.Minimized Then
                newWindow.State = StateWindow.Normal
            End If

            Dim cont = TryCast(d, MdiContainer)

            If cont Is Nothing Then
                Return
            End If

            cont.PromoteChildToFront(newWindow)
            cont.OnActivatedChild(e)
        End Sub

        Public Sub PromoteChildToFront(child As MdiChildWindow)
            Dim currentZindex = Panel.GetZIndex(child)

            For Each localChild In Children
                localChild.IsActive = localChild.Equals(child)

                Dim localZindex = Panel.GetZIndex(localChild)

                Panel.SetZIndex(localChild,If(localChild.Equals(child), _zindex - 1, If(localZindex < currentZindex, localZindex, localZindex - 1)))
            Next localChild
            child.Focus()
            ActiveChild = child
        End Sub

        Public Property FocusedWindow() As MdiChildWindow
            Get
                Return DirectCast(GetValue(FocusedWindowProperty), MdiChildWindow)
            End Get
            Set(value As MdiChildWindow)
                SetValue(FocusedWindowProperty, value)
            End Set
        End Property

        Public Property Children() As ObservableCollection(Of MdiChildWindow)
            Get
                Return DirectCast(GetValue(PropertyTypeProperty), ObservableCollection(Of MdiChildWindow))
            End Get
            Set(value As ObservableCollection(Of MdiChildWindow))
                SetValue(PropertyTypeProperty, value)
            End Set
        End Property

        Public Sub New()
            Children = New ObservableCollection(Of MdiChildWindow)()
            AddHandler Children.CollectionChanged, AddressOf ChildrenOnCollectionChanged
            InitializeComponent()
            AddHandler Loaded, AddressOf MdiContainer_Loaded
        End Sub

        Public Sub LayoutMdi(mdiLayout As MdiLayout)
            Dim point = New Windows.Point(0, 0)

            Select Case mdiLayout
                Case MdiControl.MdiLayout.Cascade
'INSTANT VB NOTE: The variable mdiChildWindow was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
                    For Each mdiChildWindow_Renamed In Children
                        mdiChildWindow_Renamed.State = StateWindow.Normal
                        mdiChildWindow_Renamed.IsMaximized = False
                        mdiChildWindow_Renamed.SetLocation(point)

                        point.X += 20
                        point.Y += 20
                    Next mdiChildWindow_Renamed
                Case MdiControl.MdiLayout.TileHorizontal
'INSTANT VB WARNING: Instant VB cannot determine whether both operands of this division are integer types - if they are then you should use the VB integer division operator:
                    Dim widthWindow = ScrollViewerRoot.ViewportWidth/Children.Count
'INSTANT VB NOTE: The variable mdiChildWindow was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
                    For Each mdiChildWindow_Renamed In Children
                        mdiChildWindow_Renamed.State = StateWindow.Normal
                        mdiChildWindow_Renamed.IsMaximized = False
                        mdiChildWindow_Renamed.SetLocation(point)

                        mdiChildWindow_Renamed.Width = widthWindow
                        mdiChildWindow_Renamed.Height = ScrollViewerRoot.ViewportHeight

                        point.X += widthWindow
                    Next mdiChildWindow_Renamed
                Case MdiControl.MdiLayout.TileVertical
'INSTANT VB WARNING: Instant VB cannot determine whether both operands of this division are integer types - if they are then you should use the VB integer division operator:
                    Dim windowHeight = ScrollViewerRoot.ViewportHeight/Children.Count
'INSTANT VB NOTE: The variable mdiChildWindow was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
                    For Each mdiChildWindow_Renamed In Children
                        mdiChildWindow_Renamed.State = StateWindow.Normal
                        mdiChildWindow_Renamed.IsMaximized = False

                        mdiChildWindow_Renamed.Width = ScrollViewerRoot.ViewportWidth
                        mdiChildWindow_Renamed.Height = windowHeight

                        mdiChildWindow_Renamed.SetLocation(point)
                        point.Y += windowHeight
                    Next mdiChildWindow_Renamed
                Case Else
                    Throw New ArgumentOutOfRangeException("mdiLayout", mdiLayout, Nothing)
            End Select

           PerformLayout()
        End Sub

        Private Sub PerformLayout()
            Dim rect = New Rect()

'INSTANT VB NOTE: The variable mdiChildWindow was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
            For Each mdiChildWindow_Renamed In Children
                rect.Union(mdiChildWindow_Renamed.Bounds)
            Next mdiChildWindow_Renamed

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
                    Dim objAdd = TryCast(args.NewItems(0), MdiChildWindow)
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

                Case NotifyCollectionChangedAction.Remove
                    Dim objRemove = TryCast(args.OldItems(0), MdiChildWindow)
                    If objRemove Is Nothing Then
                        Throw New ArgumentException("Error argument value")
                    End If
                    GridRoot.Children.Remove(objRemove)
            End Select
        End Sub

        Private Sub CheckVisibilityScrollbar()
            Dim visibleBar = Children.Any(Function(x) x.State = StateWindow.Maximized)

            ScrollViewerRoot.VerticalScrollBarVisibility = If(visibleBar, ScrollBarVisibility.Disabled, ScrollBarVisibility.Auto)
            ScrollViewerRoot.HorizontalScrollBarVisibility = If(visibleBar, ScrollBarVisibility.Disabled, ScrollBarVisibility.Auto)
        End Sub

        Private Sub objAdd_Resize(sender As Object, e As EventArgs)
            PerformLayout()
        End Sub

        Private Sub objAdd_Maximize(sender As Object, e As EventArgs)
            Dim mdiChild = DirectCast(sender, MdiChildWindow)

            If mdiChild.State = StateWindow.Maximized Then
                ScrollViewerRoot.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
                ScrollViewerRoot.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled

'INSTANT VB NOTE: The variable width was renamed since Visual Basic does not handle local variables named the same as class members well:
                Dim width_Renamed = ScrollViewerRoot.ViewportWidth + (If(GridRoot.ActualWidth > ScrollViewerRoot.ViewportWidth, SystemParameters.VerticalScrollBarWidth, 0))

'INSTANT VB NOTE: The variable height was renamed since Visual Basic does not handle local variables named the same as class members well:
                Dim height_Renamed = ScrollViewerRoot.ViewportHeight + (If(GridRoot.ActualHeight > ScrollViewerRoot.ViewportHeight, SystemParameters.HorizontalScrollBarHeight, 0))

                mdiChild.SetLocation(New Windows.Point(0, 0))
                mdiChild.SetSize(New Windows.Size(width_Renamed, height_Renamed))
            Else
                ScrollViewerRoot.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                ScrollViewerRoot.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
            End If
        End Sub

        Private Sub objAdd_Minimize(sender As Object, e As EventArgs)
            Dim mdiChild = DirectCast(sender, MdiChildWindow)

            Dim collectionMinimizedWindow = Children.Where(Function(x) x.State = StateWindow.Minimized AndAlso Math.Abs(x.Height - SystemParameters.MinimizedWindowHeight) < 0.2 AndAlso Not x.Equals(mdiChild)).ToList()

            Dim startY = ScrollViewerRoot.ViewportHeight - SystemParameters.MinimizedWindowHeight

            Dim startX = collectionMinimizedWindow.Sum(Function(mdiChildWindow) mdiChildWindow.ActualWidth + 2)

            If mdiChild.State = StateWindow.Minimized Then
                mdiChild.SetLocation(New Windows.Point(startX, startY))
            End If

            If Equals(FocusedWindow, mdiChild) Then
                FocusedWindow = Nothing
            End If
        End Sub

        Private Sub ObjAddClosing(sender As Object, e As EventArgs)
            Dim mdiChild = DirectCast(sender, MdiChildWindow)

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
            Dim point = e.GetPosition(Me)
            Dim isFound = False

            For Each mdi In From mdiItem In Children
                Let isContains = mdiItem.IsContainsPoint(point)
                Let mouseOver = mdiItem.IsMouseOver Where mouseOver AndAlso isContains
                Select mdiItem
                isFound = True
                FocusedWindow = mdi
                Exit For
            Next

            If isFound Then Return
            FocusedWindow = Nothing
        End Sub

        Private Sub ScrollViewerRoot_OnSizeChanged(sender As Object, e As SizeChangedEventArgs)
            Dim mdiChild = Children.FirstOrDefault(Function(child) child.State = StateWindow.Maximized)

            If mdiChild Is Nothing Then
                Return
            End If

            mdiChild.SetSize(New Windows.Size(ScrollViewerRoot.ActualWidth, ScrollViewerRoot.ActualHeight))
        End Sub

        Protected Overridable Sub OnActivatedChild(e As DependencyPropertyChangedEventArgs)
            Dim handler = ActivatedChildEvent
            If handler IsNot Nothing Then
                handler(Me, e)
            End If
        End Sub

        Protected Overrides Sub OnRenderSizeChanged(sizeInfo As SizeChangedInfo)
            Dim rect = New Rect()

'INSTANT VB NOTE: The variable mdiChildWindow was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
            For Each mdiChildWindow_Renamed In Children
                rect.Union(mdiChildWindow_Renamed.Bounds)
            Next mdiChildWindow_Renamed

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
