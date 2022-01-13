''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Windows
Imports System.Windows.Input

Namespace Common.DataViewerControl
    ''' <summary>
    ''' Interaction logic for PaginationPanel.xaml
    ''' </summary>
    Partial Friend Class PaginationPanel
        Private Const TitleCheckBoxWithOffset As String = "Enable pagination"
        Private Const TitleCheckBoxWithoutOffset As String = "Enable limitations"
        Public Shared ReadOnly IsSupportLimitCountProperty As DependencyProperty = DependencyProperty.Register("IsSupportLimitCount", GetType(Boolean), GetType(PaginationPanel), New PropertyMetadata(True, Sub(o As DependencyObject, args As DependencyPropertyChangedEventArgs)
                    Dim obj = TryCast(o, PaginationPanel)
                    If obj IsNot Nothing Then
                        obj.PanelLimit.Visibility = If(DirectCast(args.NewValue, Boolean), Visibility.Visible, Visibility.Collapsed)
                    End If
        End Sub))

        Public Property IsSupportLimitCount() As Boolean
            Get
                Return CBool(GetValue(IsSupportLimitCountProperty))
            End Get
            Set(value As Boolean)
                SetValue(IsSupportLimitCountProperty, value)
            End Set
        End Property

        Public Shared ReadOnly IsSupportLimitOffsetProperty As DependencyProperty = DependencyProperty.Register("IsSupportLimitOffset", GetType(Boolean), GetType(PaginationPanel), New PropertyMetadata(True, Sub(o As DependencyObject, args As DependencyPropertyChangedEventArgs)
                    Dim value = DirectCast(args.NewValue, Boolean)
                    Dim obj = TryCast(o, PaginationPanel)

                    If obj Is Nothing Then
                        Return
                    End If

                    obj.PanelOffset.Visibility = If(DirectCast(args.NewValue, Boolean), Visibility.Visible, Visibility.Collapsed)
                    obj.CheckBoxEnabled.Content = If(value, TitleCheckBoxWithOffset, TitleCheckBoxWithoutOffset)
        End Sub))

        Public Property IsSupportLimitOffset() As Boolean
            Get
                Return CBool(GetValue(IsSupportLimitOffsetProperty))
            End Get
            Set(value As Boolean)
                SetValue(IsSupportLimitOffsetProperty, value)
            End Set
        End Property

        Public Event EnabledPaginationChanged As RoutedEventHandler
        Public Event CurrentPageChanged As RoutedEventHandler
        Public Event PageSizeChanged As RoutedEventHandler

        Public Shared ReadOnly PageSizeProperty As DependencyProperty = DependencyProperty.Register("PageSize", GetType(Integer), GetType(PaginationPanel), New PropertyMetadata(10, Sub(o, args)
              Dim obj = TryCast(o, PaginationPanel)
              If obj IsNot Nothing Then
                  obj.TextBoxPageSize.Text = args.NewValue.ToString()
              End If
        End Sub))

        Public Property PageSize() As Integer
            Get
                Return CInt(GetValue(PageSizeProperty))
            End Get
            Set(value As Integer)
                SetValue(PageSizeProperty, value)
                _maxPageCount = _countRows \ PageSize + (If(_maxPageCount * PageSize < _countRows, 1, 0))
            End Set
        End Property

        Private _maxPageCount As Integer

        Public Shared ReadOnly CurrentPageProperty As DependencyProperty = DependencyProperty.Register("CurrentPage", GetType(Integer), GetType(PaginationPanel), New PropertyMetadata(1, Sub(o, args)
              Dim obj = TryCast(o, PaginationPanel)
              If obj IsNot Nothing Then
                  obj.TextBoxCurrentPage.Text = args.NewValue.ToString()
              End If
        End Sub))

        Public Property CurrentPage() As Integer
            Get
                Return CInt(GetValue(CurrentPageProperty))
            End Get
            Set(value As Integer)
                SetValue(CurrentPageProperty, value)
            End Set
        End Property

        Private _countRows As Integer

        Public Property CountRows() As Integer
            Set(value As Integer)
                _countRows = value
                _maxPageCount = _countRows \ PageSize
                If _maxPageCount * PageSize < _countRows Then
                    _maxPageCount += 1
                End If
            End Set
            Get
                Return _countRows
            End Get
        End Property

        Public Shadows ReadOnly Property IsEnabled() As Boolean
            Get
                Dim isChecked = CheckBoxEnabled.IsChecked
                Return isChecked.HasValue AndAlso isChecked.Value = True
            End Get
        End Property

        Public Sub New()
            InitializeComponent()

            TextBoxCurrentPage.Text = CurrentPage.ToString()
            TextBoxPageSize.Text = PageSize.ToString()
        End Sub

        Public Sub Reset()
            ToggleEnabled(False)

            CurrentPage = 1
            PageSize = 10
            _maxPageCount = 1
        End Sub

        Private Sub CheckBoxEnabled_CheckedChanged(sender As Object, e As RoutedEventArgs)
            Dim isChecked = CheckBoxEnabled.IsChecked
            ToggleEnabled(isChecked.HasValue AndAlso isChecked.Value = True)

            OnEnabledPaginationChanged(e)
        End Sub

        Private Sub ButtonPreviewPage_OnClick(sender As Object, e As RoutedEventArgs)
            If CurrentPage - 1 = 0 Then
                Return
            End If

            CurrentPage -= 1
            OnCurrentPageChanged(e)

        End Sub

        Private Sub ButtonNextPage_OnClick(sender As Object, e As RoutedEventArgs)
            If CurrentPage + 1 > _maxPageCount Then
                Return
            End If

            CurrentPage += 1
            OnCurrentPageChanged(e)
        End Sub

        Private Sub TextBoxCurrentPage_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            Dim value As Integer = Nothing

            Dim success = Integer.TryParse(TextBoxCurrentPage.Text, value)
            If success Then
                CurrentPage = value
                OnCurrentPageChanged(New RoutedEventArgs(e.RoutedEvent))
            Else
                TextBoxCurrentPage.Text = CurrentPage.ToString()
            End If
        End Sub

        Private Sub TextBoxPageSize_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            Dim value As Integer = Nothing

            Dim success = Integer.TryParse(TextBoxPageSize.Text, value)
            If success Then
                PageSize = value
                OnPageSizeChanged(New RoutedEventArgs(e.RoutedEvent))
            Else
                TextBoxPageSize.Text = PageSize.ToString()
            End If
        End Sub

        Private Sub TextBox_OnKeyDown(sender As Object, e As KeyEventArgs)
            If e.Key <> Key.Enter Then
                Return
            End If

            Keyboard.ClearFocus()
        End Sub

        Private Sub ToggleEnabled(value As Boolean)
            CheckBoxEnabled.IsChecked = value
            TextBoxCurrentPage.IsEnabled = value
            TextBoxPageSize.IsEnabled = value
            ButtonPreviewPage.IsEnabled = value
            ButtonNextPage.IsEnabled = value
        End Sub

        Private Sub TextBoxCurrentPage_OnPreviewTextInput(sender As Object, e As TextCompositionEventArgs)
            e.Handled = IsTextNumeric(e.Text)
        End Sub

        Private Sub TextBoxPageSize_OnPreviewTextInput(sender As Object, e As TextCompositionEventArgs)
            e.Handled = IsTextNumeric(e.Text)
        End Sub

        Private Shared Function IsTextNumeric(str As String) As Boolean
            Dim reg = New System.Text.RegularExpressions.Regex("[^0-9]")
            Return reg.IsMatch(str)
        End Function

        #Region "Invokators"

        Protected Overridable Sub OnEnabledPaginationChanged(e As RoutedEventArgs)
            RaiseEvent EnabledPaginationChanged(Me, e)
        End Sub

        Protected Overridable Sub OnCurrentPageChanged(e As RoutedEventArgs)
            If CheckBoxEnabled.IsChecked <> True Then
                Return
            End If

            Dim handler = CurrentPageChangedEvent
            If handler IsNot Nothing Then
                handler(Me, e)
            End If
        End Sub

        Protected Overridable Sub OnPageSizeChanged(e As RoutedEventArgs)
            If CheckBoxEnabled.IsChecked <> True Then
                Return
            End If

            Dim handler = PageSizeChangedEvent
            If handler IsNot Nothing Then
                handler(Me, e)
            End If
        End Sub

        #End Region
    End Class
End Namespace
