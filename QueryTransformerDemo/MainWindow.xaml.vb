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
Imports System.ComponentModel
Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.QueryTransformer

Partial Public Class MainWindow
    Private _sqlContext As SQLContext
    Private _sqlQuery As SQLQuery
    Private _queryTransformer As QueryTransformer

    ' List of query output columns of the current SQL query used for turning their visibility on and off
    Private _collectionColumns As ObservableCollection(Of VisibleColumn)

    Public Sub New()
        InitializeComponent()
        AddHandler Loaded, AddressOf MainWindow_Loaded
    End Sub

    Private Sub LoadQuery()
        ' Clear the input fields
        ClearFieldsSorting()
        ClearFieldsAggregate()
        ClearFieldsWhere()

        ' Load a query into the SQLQuery object. 
        _sqlQuery.SQL = BoxSourceSql.Text

        ' Set counter values
        CounterSortingActive.Text = _queryTransformer.Sortings.Count.ToString()
        CounterAggregations.Text = _queryTransformer.Aggregations.Count.ToString()
        CounterWhere.Text = _queryTransformer.Filters.Count.ToString()

        _collectionColumns.Clear()

        ' Fill the list of output columns to be used as ItemsSource for a combobox
        For Each column As OutputColumn In _queryTransformer.Columns
            _collectionColumns.Add(New VisibleColumn(column))
        Next

        ListBoxVisibleColumns.ItemsSource = _collectionColumns

        RefreshBindings()

        CounterVisibleColumn.Text = _collectionColumns.Where(Function(x) x.Visible).Count().ToString()
    End Sub

    Private Sub RefreshBindings()
        ComboBoxColumnsSorting.ItemsSource = Nothing
        ComboBoxColumnAggregate.ItemsSource = Nothing
        ComboBoxColumnWhere.ItemsSource = Nothing

        ComboBoxColumnsSorting.ItemsSource = _queryTransformer.Columns
        ComboBoxColumnAggregate.ItemsSource = _queryTransformer.Columns
        ComboBoxColumnWhere.ItemsSource = _queryTransformer.Columns
    End Sub

    Private Sub ClearFieldsSorting()
        ComboBoxColumnsSorting.SelectedItem = Nothing
        ComboBoxSorting.SelectedItem = Nothing
    End Sub

    Private Sub ClearFieldsAggregate()
        ComboBoxColumnAggregate.SelectedItem = Nothing
        ComboBoxFunction.SelectedItem = Nothing
    End Sub

    Private Sub ClearFieldsWhere()
        ComboBoxColumnWhere.SelectedItem = Nothing
        ComboBoxConditions.SelectedItem = Nothing
        ComboBoxConditions.SelectedItem = Nothing
        BoxWhereValue.Text = String.Empty
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

        _collectionColumns = New ObservableCollection(Of VisibleColumn)()
        AddHandler _collectionColumns.CollectionChanged, AddressOf _collectionColumns_CollectionChanged

        _sqlContext = New SQLContext() With {
            .SyntaxProvider = New MSSQLSyntaxProvider() With {
                .ServerVersion = MSSQLServerVersion.MSSQL2012
            }
        }
        _sqlContext.MetadataContainer.LoadingOptions.OfflineMode = True
        _sqlContext.MetadataContainer.ImportFromXML("Northwind.xml")

        ' Load a sample query
        Dim sqlText As StringBuilder = New StringBuilder()
        sqlText.AppendLine("Select Categories.CategoryName,")
        sqlText.AppendLine("Products.QuantityPerUnit")
        sqlText.AppendLine("From Categories")
        sqlText.AppendLine("Inner Join Products On Categories.CategoryID = Products.CategoryID")

        _sqlQuery = New SQLQuery(_sqlContext)
        AddHandler _sqlQuery.SQLUpdated, AddressOf _sqlQuery_SQLUpdated
        _sqlQuery.SQL = sqlText.ToString()
        _queryTransformer = New QueryTransformer() With {
            .Query = _sqlQuery
        }

        AddHandler _queryTransformer.SQLUpdated, AddressOf _queryTransformer_SQLUpdated
        _queryTransformer.SQLGenerationOptions = New SQLFormattingOptions()
        LoadQuery()
    End Sub

    Private Sub _sqlQuery_SQLUpdated(sender As Object, e As EventArgs)
        BoxSourceSql.Text = _sqlQuery.SQL
    End Sub

    Private Sub _queryTransformer_SQLUpdated(sender As Object, e As EventArgs)
        ' Get the transformed SQL query text
        BoxResultSql.Text = _queryTransformer.SQL
    End Sub

    Private Sub _collectionColumns_CollectionChanged(sender As Object, e As Collections.Specialized.NotifyCollectionChangedEventArgs)
        If e.NewItems IsNot Nothing Then
            For Each item As INotifyPropertyChanged In e.NewItems
                AddHandler item.PropertyChanged, AddressOf ColumnVisibleChanged
            Next
        End If

        If e.OldItems IsNot Nothing Then
            For Each item As INotifyPropertyChanged In e.OldItems
                RemoveHandler item.PropertyChanged, AddressOf ColumnVisibleChanged
            Next
        End If
    End Sub

    Private Sub ColumnVisibleChanged(sender As Object, e As PropertyChangedEventArgs)
        CounterVisibleColumn.Text = _collectionColumns.Where(Function(x) x.Visible).Count().ToString()
    End Sub

    Private Sub ButtonLoad_OnClick(sender As Object, e As RoutedEventArgs)
        ' Load a query and updating the form controls
        LoadQuery()
    End Sub

    Private Sub ButtonClearVisibleColumn_OnClick(sender As Object, e As RoutedEventArgs)
        For Each column As VisibleColumn In _collectionColumns
            column.Visible = True
        Next
    End Sub

    Private Sub ButtonAddSorting_OnClick(sender As Object, e As RoutedEventArgs)
        Dim column As OutputColumn = DirectCast(ComboBoxColumnsSorting.SelectedItem, OutputColumn)
        If Not column.IsSupportSorting Then
            Return
        End If

        Dim sortedColumn As SortedColumn = Nothing
        If ComboBoxSorting.SelectedValue Is Nothing Then
            Return
        End If

        Select Case ComboBoxSorting.SelectedValue.ToString()
            Case "Asc"
                sortedColumn = column.Asc()
                Exit Select
            Case "Desc"
                sortedColumn = column.Desc()
                Exit Select
        End Select

        ' Add sorting to the query - the sort order of original query will be overridden.
        _queryTransformer.OrderBy(sortedColumn)

        CounterSortingActive.Text = _queryTransformer.Sortings.Count.ToString()

        ClearFieldsSorting()

    End Sub

    Private Sub ButtonClearSorting_OnClick(sender As Object, e As RoutedEventArgs)
        ClearFieldsSorting()

        ' Remove the added sortings from the query - the original sort order will be restored.
        _queryTransformer.Sortings.Clear()

        CounterSortingActive.Text = _queryTransformer.Sortings.Count.ToString()
    End Sub

    Private Sub ButtonAddAggregate_OnClick(sender As Object, e As RoutedEventArgs)
        Dim column As OutputColumn = DirectCast(ComboBoxColumnAggregate.SelectedItem, OutputColumn)

        Dim aggregatedColumn As AggregatedColumn = Nothing
        Select Case ComboBoxFunction.SelectedItem.ToString()
            Case "Count"
                aggregatedColumn = column.Count()
                Exit Select
            Case "Avg"
                aggregatedColumn = column.Avg()
                Exit Select
            Case "Sum"
                aggregatedColumn = column.Sum()
                Exit Select
            Case "Min"
                aggregatedColumn = column.Min()
                Exit Select
            Case "Max"
                aggregatedColumn = column.Max()
                Exit Select
        End Select

        ' Add an aggregate to the query - if any aggregates are added, all other columns will be removed from the query.
        _queryTransformer.Aggregations.Add(aggregatedColumn)

        CounterAggregations.Text = _queryTransformer.Aggregations.Count.ToString()

        ClearFieldsAggregate()
    End Sub

    Private Sub ButtonClearAggregate_Click(sender As Object, e As RoutedEventArgs)
        ' Clear all aggregates from the query - the columns of original query will be restored.
        _queryTransformer.Aggregations.Clear()

        CounterAggregations.Text = _queryTransformer.Aggregations.Count.ToString()

        ClearFieldsAggregate()
    End Sub

    Private Sub ButtonAddWhere_OnClick(sender As Object, e As RoutedEventArgs)
        Dim column As OutputColumn = DirectCast(ComboBoxColumnWhere.SelectedItem, OutputColumn)

        Dim condition As FilterCondition = Nothing
        Select Case ComboBoxConditions.SelectedItem.ToString()
            Case "Equal"
                condition = column.Equal(BoxWhereValue.Text)
                Exit Select
            Case "Not Equal"
                condition = column.Not_Equal(BoxWhereValue.Text)
                Exit Select
            Case "Greater"
                condition = column.Greater(BoxWhereValue.Text)
                Exit Select
            Case "GreaterEqual"
                condition = column.GreaterEqual(BoxWhereValue.Text)
                Exit Select
            Case "Not Grater"
                condition = column.Not_Greater(BoxWhereValue.Text)
                Exit Select
            Case "Not GreaterEqual"
                condition = column.Not_GreaterEqual(BoxWhereValue.Text)
                Exit Select
            Case "IsNull"
                condition = column.IsNull()
                Exit Select
            Case "Not IsNull"
                condition = column.Not_IsNull()
                Exit Select
            Case "IsNotNull"
                condition = column.IsNotNull()
                Exit Select
            Case "Less"
                condition = column.Less(BoxWhereValue.Text)
                Exit Select
            Case "LessEqual"
                condition = column.LessEqual(BoxWhereValue.Text)
                Exit Select
            Case "Not Less"
                condition = column.Not_Less(BoxWhereValue.Text)
                Exit Select
            Case "Not LessEqual"
                condition = column.Not_LessEqual(BoxWhereValue.Text)
                Exit Select
            Case "In"
                condition = column.[In](BoxWhereValue.Text)
                Exit Select
            Case "Not In"
                condition = column.Not_In(BoxWhereValue.Text)
                Exit Select
            Case "Like"
                condition = column.[Like](BoxWhereValue.Text)
                Exit Select
            Case "Not Like"
                condition = column.Not_Like(BoxWhereValue.Text)
                Exit Select
        End Select

        ' Add new filter to the query - the filter will be added to the WHERE clause of original query.
        _queryTransformer.Filters.Add(condition)

        CounterWhere.Text = _queryTransformer.Filters.Count.ToString()

        ClearFieldsWhere()
    End Sub

    Private Sub ButtonClearWhere_OnClick(sender As Object, e As RoutedEventArgs)
        ' Remove all additional filters from query - the original WHERE clause will be restored.
        _queryTransformer.Filters.Clear()

        CounterWhere.Text = _queryTransformer.Filters.Count.ToString()

        ClearFieldsWhere()
    End Sub

    Private Sub ButtonGetCode_OnClick(sender As Object, e As RoutedEventArgs)
        Dim window As Window = New Window() With {
            .Owner = Me,
            .WindowStartupLocation = WindowStartupLocation.CenterOwner,
            .Width = 600,
            .Height = 300,
            .ResizeMode = ResizeMode.NoResize,
            .Title = "Code Behind"
        }

        Dim grid1 As Grid = New Grid()
        grid1.RowDefinitions.Add(New RowDefinition())
        grid1.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})

        Dim buttonClose As Button = New Button() With {
            .Content = "Close",
            .Margin = New Thickness(0, 0, 10, 10),
            .HorizontalAlignment = HorizontalAlignment.Right,
            .Width = 75,
            .Height = 23
        }

        AddHandler buttonClose.Click, Sub()
                                          window.Close()
                                      End Sub

        buttonClose.SetValue(Grid.RowProperty, 1)

        Dim textbox As TextBox = New TextBox() With {
            .Margin = New Thickness(10),
            .Text = GetCodeBehind(),
            .Background = New SolidColorBrush(SystemColors.InfoColor)
        }
        textbox.SetValue(Grid.RowProperty, 0)

        window.Content = grid1
        grid1.Children.Add(textbox)
        grid1.Children.Add(buttonClose)
        window.ShowDialog()
    End Sub

    Private Function GetCodeBehind() As String
        Dim builder As StringBuilder = New StringBuilder()
        builder.AppendLine("_queryTransformer")

        For Each sorting As SortedColumn In _queryTransformer.Sortings
            Dim text as String = String.Format(vbTab & ".OrderBy(_queryTransformer.Columns[{0}], {1})", _queryTransformer.Columns.IndexOf(sorting.Column), (sorting.SortType = ItemSortType.Asc).ToString().ToLower(CultureInfo.CurrentCulture))
            builder.AppendLine(text)
        Next

        Dim reg As Regex = New Regex("([A-Z])\w+")
        For Each aggregation As AggregatedColumn In _queryTransformer.Aggregations
            Dim result As Match = reg.Match(aggregation.Expression)
            If Not result.Success Then
                Continue For
            End If

            Dim text as String = String.Format(vbTab & ".Select(_queryTransformer.Columns[{0}].{1}())", _queryTransformer.Columns.IndexOf(aggregation.Column), result.Value)
            builder.AppendLine(text)
        Next

        For Each filter As FilterColumnCondition In _queryTransformer.Filters

            Dim text as String = String.Format(vbTab & ".Where(""{0}"")", filter.Condition)
            builder.AppendLine(text)
        Next

        Return builder.ToString()
    End Function
End Class

Public Class VisibleColumn
    Implements INotifyPropertyChanged
    Private ReadOnly _column As OutputColumn

    Public Property Visible() As Boolean
        Get
            Return _column IsNot Nothing AndAlso _column.Visible
        End Get
        Set
            If _column Is Nothing Then
                Return
            End If
            _column.Visible = Value
            OnPropertyChanged("Visible")
        End Set
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return If(_column Is Nothing, String.Empty, _column.Column.Expression)
        End Get
    End Property

    Public Sub New(column As OutputColumn)
        _column = column
    End Sub

    Protected Overridable Sub OnPropertyChanged(propertyName As String)
        RaiseEvent INotifyPropertyChanged_PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Event INotifyPropertyChanged_PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
End Class
