'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data
Imports System.Data.Common
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.QueryTransformer

Namespace Common
    ''' <summary>
    ''' Interaction logic for CustomDataGrid.xaml
    ''' </summary>
    Partial Public Class CustomDataGrid
        Public Event RowsLoaded As EventHandler

        Private _currentTextSql As String
        Private _needCancelOperation As Boolean
        Private _currentTask As Task(Of DataView)
        Private _nextTask As Task(Of DataView)

        Public Property QueryTransformer As QueryTransformer
        Public Property SqlQuery As SQLQuery

        Public ReadOnly Property CountRows() As Integer
            Get
                Return DataView.Items.Count
            End Get
        End Property

        Public Sub New()
            InitializeComponent()
        End Sub

        Public Sub FillDataGrid(sqlCommand As String)
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

            Dim task As Task(Of DataView) = New Task(Of DataView)(Function() ExecuteSql(sqlCommand))

            If _currentTask Is Nothing Then
                _currentTask = task
                TryRunTask()
            Else
                _nextTask = task
                _needCancelOperation = True
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

            Dim command As DbCommand = Helpers.CreateSqlCommand(sqlCommand, SqlQuery)

            Dim table As New DataTable("result")

            Try
                Using dbReader As DbDataReader = command.ExecuteReader()

                    For i As Integer = 0 To dbReader.FieldCount - 1
                        table.Columns.Add(dbReader.GetName(i))
                    Next
                    While dbReader.Read() AndAlso Not _needCancelOperation
                        Dim values As Object() = New Object(dbReader.FieldCount - 1) {}
                        dbReader.GetValues(values)
                        table.Rows.Add(values)
                    End While

                    Return table.DefaultView
                End Using
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
                Dispatcher.Invoke(Sub() GridLoadMessage.Visibility = Visibility.Visible)

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
            _needCancelOperation = False

            Dispatcher.Invoke(Sub()
                                  DataView.ItemsSource = obj.Result

                                  GridLoadMessage.Visibility = If(_currentTask IsNot Nothing, Visibility.Visible, Visibility.Collapsed)

                                  OnRowsLoaded()

                              End Sub)
            TryRunTask()
        End Sub

        Private Sub ShowException(ex As Exception)
            Dispatcher.BeginInvoke(DirectCast(Sub()
                                                  BorderError.Visibility = Visibility.Visible
                                                  LabelError.Text = ex.Message

                                              End Sub, Action))
        End Sub

        Private Sub HeaderColumn_OnClick(sender As Object, e As RoutedEventArgs)
            Dim headerColumn As DataGridColumnHeader = TryCast(sender, DataGridColumnHeader)
            If headerColumn Is Nothing Then
                Return
            End If

            If Not Keyboard.IsKeyDown(Key.LeftCtrl) AndAlso Not Keyboard.IsKeyDown(Key.RightCtrl) Then
                QueryTransformer.Sortings.Clear()
            End If

            Dim header As HeaderData = DirectCast(headerColumn.Content, HeaderData)

            Dim columnForSorting As OutputColumn = QueryTransformer.Columns.FindColumnByOriginalName(header.Title)

            Dim columnSorted As SortedColumn

            Select Case header.Sorting
                Case Sorting.Asc
                    columnSorted = columnForSorting.Desc()
                    QueryTransformer.Sortings.Add(columnSorted)
                    Exit Select
                Case Sorting.Desc
                    QueryTransformer.Sortings.Remove(columnForSorting)
                    Exit Select
                Case Sorting.None
                    columnSorted = columnForSorting.Asc()
                    QueryTransformer.Sortings.Add(columnSorted)
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select

            FillDataGrid(QueryTransformer.SQL)
        End Sub

        Private Sub DataView_AutoGeneratedColumns(sender As Object, e As EventArgs)
            For Each dataGridColumn As DataGridColumn In DataView.Columns
                Dim data As HeaderData = New HeaderData() With {
                    .Title = dataGridColumn.SortMemberPath
                }
                dataGridColumn.Header = data

                Dim sorting1 As SortedColumn = QueryTransformer.Sortings.FirstOrDefault(Function(sort) sort.Column.OriginalName = data.Title)

                If sorting1 Is Nothing Then
                    Continue For
                End If

                data.Counter = QueryTransformer.Sortings.IndexOf(sorting1) + 1

                Select Case sorting1.SortType
                    Case ItemSortType.None
                        data.Sorting = Sorting.None
                        Exit Select
                    Case ItemSortType.Asc
                        data.Sorting = Sorting.Asc
                        Exit Select
                    Case ItemSortType.Desc
                        data.Sorting = Sorting.Desc
                        Exit Select
                    Case Else
                        Throw New ArgumentOutOfRangeException()
                End Select
            Next
        End Sub

        Private Sub CloseImage_OnMouseUp(sender As Object, e As MouseButtonEventArgs)
            BorderError.Visibility = Visibility.Collapsed
        End Sub

        Public Sub StopFill()
            If _currentTask Is Nothing OrElse _currentTask.Status <> TaskStatus.Running Then
                Return
            End If

            ButtonBreakLoad.IsEnabled = False
            _needCancelOperation = True
            _nextTask = Nothing
        End Sub

        Private Sub ButtonCancelExecutingSql_OnClick(sender As Object, e As RoutedEventArgs)
            StopFill()
        End Sub

        Protected Overridable Sub OnRowsLoaded()
            RaiseEvent RowsLoaded(Me, EventArgs.Empty)
        End Sub
    End Class
End Namespace
