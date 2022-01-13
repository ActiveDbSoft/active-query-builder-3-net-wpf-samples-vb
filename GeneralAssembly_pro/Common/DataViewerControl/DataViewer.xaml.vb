''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Threading.Tasks
Imports System.Windows.Controls.Primitives
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.QueryTransformer

Namespace Common.DataViewerControl
    ''' <summary>
    ''' Interaction logic for DataViewer.xaml
    ''' </summary>
    Partial Public Class DataViewer

        Private _currentTextSql As String

        Private _currentTask As Task(Of DataView)
        Private _nextTask As Task(Of DataView)

        Private _queryTransformer As QueryTransformer
        Private _sqlQuery As SQLQuery

        Private _sqlGenerationOptions As SQLGenerationOptions

        Public Property QueryTransformer() As QueryTransformer
            Get
                Return _queryTransformer
            End Get
            Set(value As QueryTransformer)
                _queryTransformer = value
                If value IsNot Nothing Then
                    _sqlQuery = DirectCast(_queryTransformer.Query, SQLQuery)
                End If
                CheckCanPagination()
            End Set
        End Property

        Public Property SqlGenerationOptions() As SQLGenerationOptions
            Get
                Return _sqlGenerationOptions
            End Get
            Set(value As SQLGenerationOptions)
                _sqlGenerationOptions = value
                If value IsNot Nothing AndAlso _queryTransformer IsNot Nothing Then
                    _queryTransformer.SQLGenerationOptions = _sqlGenerationOptions
                End If
            End Set
        End Property
        Public Property SqlQuery() As SQLQuery
            Set(value As SQLQuery)
                _sqlQuery = value

                If _sqlQuery Is Nothing Then
                    Return
                End If

                If _queryTransformer IsNot Nothing Then
                    Return
                End If

                _queryTransformer = New QueryTransformer With {.Query = _sqlQuery}

                If _sqlGenerationOptions IsNot Nothing Then
                    _queryTransformer.SQLGenerationOptions = _sqlGenerationOptions
                End If

                CheckCanPagination()
            End Set
            Get
                Return _sqlQuery
            End Get
        End Property

        Public ReadOnly Property CountRows() As Integer
            Get
                Return DataView.Items.Count
            End Get
        End Property

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub CheckCanPagination()
            ' The pagination panel is displayed if the current SyntaxProvider has support for pagination
            PaginationPanel.Visibility = If((QueryTransformer.IsSupportLimitCount OrElse QueryTransformer.IsSupportLimitOffset) AndAlso _sqlQuery.SQLContext.SyntaxProvider IsNot Nothing, Visibility.Visible, Visibility.Collapsed)


            PaginationPanel.IsSupportLimitCount = QueryTransformer.IsSupportLimitCount
            PaginationPanel.IsSupportLimitOffset = QueryTransformer.IsSupportLimitOffset
        End Sub

        Public Sub FillData(sqlCommand As String)
            BorderError.Visibility = Visibility.Collapsed
            ButtonBreakLoad.IsEnabled = True

            If _currentTextSql = sqlCommand OrElse String.IsNullOrWhiteSpace(sqlCommand) Then
                If Not String.IsNullOrWhiteSpace(sqlCommand) Then
                    Return
                End If

                DataView.ItemsSource = Nothing
                OnRowsLoaded()
                Return
            End If

            _currentTextSql = sqlCommand

            DataView.ItemsSource = Nothing

            Dim task = New Task(Of DataView)(Function() ExecuteSql(sqlCommand))

            If _currentTask Is Nothing Then
                _currentTask = task
                TryRunTask()
            Else
                _nextTask = task
            End If
        End Sub

        Public Sub Clear()
            DataView.ItemsSource = Nothing
        End Sub

        Private Function ExecuteSql(sqlCommand As String) As DataView
            If String.IsNullOrEmpty(sqlCommand) Then
                Return Nothing
            End If

            If SqlQuery.SQLContext.MetadataProvider Is Nothing Then
                Return Nothing
            End If

            If Not SqlQuery.SQLContext.MetadataProvider.Connected Then
                SqlQuery.SQLContext.MetadataProvider.Connect()
            End If

            Try

                Dim table As DataView = SqlHelpers.GetDataView(sqlCommand, SqlQuery)
                Return table

            Catch ex As Exception
                ShowException(ex)
            End Try

            Return Nothing
        End Function

        Private Sub TryRunTask()
            If _currentTask IsNot Nothing AndAlso (_currentTask.Status = TaskStatus.Running OrElse _currentTask.Status = TaskStatus.WaitingToRun) Then
                Return
            End If

            If _currentTask IsNot Nothing Then
                Dispatcher?.Invoke(Sub()
                    GridLoadMessage.Visibility = Visibility.Visible
                End Sub)

                _currentTask.ContinueWith(AddressOf TaskCompleted)

                _currentTask.Start()

                Return
            End If

            If _nextTask Is Nothing Then
                Return
            End If
            _currentTask = _nextTask
            _nextTask = Nothing

            TryRunTask()
        End Sub

        Private Sub TaskCompleted(obj As Task(Of DataView))
            _currentTask = Nothing

            Dispatcher?.Invoke(Sub()
                DataView.ItemsSource = obj.Result

                GridLoadMessage.Visibility = If(_currentTask IsNot Nothing, Visibility.Visible, Visibility.Collapsed)

                OnRowsLoaded()
            End Sub)
            TryRunTask()
        End Sub

        Private Sub ShowException(ex As Exception)
            Dispatcher?.BeginInvoke(CType(Sub()
                BorderError.Visibility = Visibility.Visible
                LabelError.Text = ex.Message
            End Sub, Action))
        End Sub

        Private Sub HeaderColumn_OnClick(sender As Object, e As RoutedEventArgs)
            If _queryTransformer Is Nothing Then
                Return
            End If

            Dim headerColumn = TryCast(sender, DataGridColumnHeader)
            If headerColumn Is Nothing Then
                Return
            End If

            If Not Keyboard.IsKeyDown(Key.LeftCtrl) AndAlso Not Keyboard.IsKeyDown(Key.RightCtrl) Then
                _queryTransformer.Sortings.Clear()
            End If

            Dim header = DirectCast(headerColumn.Content, HeaderDataModel)

            Dim columnForSorting = _queryTransformer.Columns.FindColumnByOriginalName(header.Title)

            Dim columnSorted As SortedColumn

            Select Case header.Sorting
                Case Sorting.Asc
                    columnSorted = columnForSorting.Desc()
                    _queryTransformer.Sortings.Add(columnSorted)
                Case Sorting.Desc
                    _queryTransformer.Sortings.Remove(columnForSorting)
                Case Sorting.None
                    columnSorted = columnForSorting.Asc()
                    _queryTransformer.Sortings.Add(columnSorted)
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select

            FillData(_queryTransformer.SQL)
        End Sub

        Private Sub DataView_AutoGeneratedColumns(sender As Object, e As EventArgs)
            If _queryTransformer Is Nothing Then
                Return
            End If

            For Each dataGridColumn In DataView.Columns
                Dim data = New HeaderDataModel With {.Title = dataGridColumn.SortMemberPath}
                dataGridColumn.Header = data

                Dim sorting = _queryTransformer.Sortings.FirstOrDefault(Function(sort) sort.Column.OriginalName = data.Title)

                If sorting Is Nothing Then
                    Continue For
                End If

                data.Counter = _queryTransformer.Sortings.IndexOf(sorting) + 1

                Select Case sorting.SortType
                    Case ItemSortType.None
                        data.Sorting = Common.DataViewerControl.Sorting.None
                    Case ItemSortType.Asc
                        data.Sorting = Common.DataViewerControl.Sorting.Asc
                    Case ItemSortType.Desc
                        data.Sorting = Common.DataViewerControl.Sorting.Desc
                    Case Else
                        Throw New ArgumentOutOfRangeException()
                End Select
            Next dataGridColumn
        End Sub

        Private Sub CloseImage_OnMouseUp(sender As Object, e As MouseButtonEventArgs)
            BorderError.Visibility = Visibility.Collapsed
        End Sub

        Private Sub ButtonCancelExecutingSql_OnClick(sender As Object, e As RoutedEventArgs)
            If _currentTask Is Nothing OrElse _currentTask.Status <> TaskStatus.Running Then
                Return
            End If

            ButtonBreakLoad.IsEnabled = False
            _nextTask = Nothing
        End Sub

        Protected Overridable Sub OnRowsLoaded()
            If Not PaginationPanel.IsEnabled Then
                PaginationPanel.CountRows = CountRows
            End If
            BorderBlockPagination.Visibility = Visibility.Collapsed
        End Sub

        Private Sub PaginationPanel_OnCurrentPageChanged(sender As Object, e As RoutedEventArgs)
            If _queryTransformer Is Nothing Then
                Return
            End If

            If PaginationPanel.CurrentPage = 1 Then
                _queryTransformer.Skip("")
                Return
            End If

            ' Select next n records
            _queryTransformer.Skip((PaginationPanel.PageSize * (PaginationPanel.CurrentPage - 1)).ToString())
        End Sub

        Private Sub PaginationPanel_OnEnabledPaginationChanged(sender As Object, e As RoutedEventArgs)
            If _queryTransformer Is Nothing Then
                Return
            End If
            ' Turn paging on and off
            If PaginationPanel.IsEnabled Then
                _queryTransformer.Take(PaginationPanel.PageSize.ToString())
            Else
                ResetPagination()
            End If
        End Sub

        Private Sub PaginationPanel_OnPageSizeChanged(sender As Object, e As RoutedEventArgs)
            _queryTransformer.Take(PaginationPanel.PageSize.ToString())
        End Sub

        Private Sub ResetPagination()
            If _queryTransformer Is Nothing Then
                Return
            End If

            PaginationPanel.Reset()
            _queryTransformer.Skip("")
            _queryTransformer.Take("")
        End Sub
    End Class
End Namespace
