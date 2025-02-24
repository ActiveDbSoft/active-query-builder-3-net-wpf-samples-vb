''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.QueryView
Imports QueryColumnListItem = ActiveQueryBuilder.View.WPF.QueryView.QueryColumnListItem

Namespace Frames
    ''' <summary>
    ''' Interaction logic for CustomComobBoxColumnDemoFrame.xaml
    ''' </summary>
    Public Partial Class CustomComobBoxColumnDemoFrame
        Private ReadOnly _customValues As List(Of String)
        Private ReadOnly _customValuesProvider As New List(Of String)()

        Public Sub New()
            InitializeComponent()

            ' Fill the source of custom values (for demo purposes)
            For i As Int32 = 0 To 99
                _customValuesProvider.Add("Some Value " & i)
            Next

            _customValues = _customValuesProvider.ToList()

            ' Fill query builder with demo data
            QueryBuilder1.SyntaxProvider = New MSSQLSyntaxProvider()
            QueryBuilder1.MetadataLoadingOptions.OfflineMode = True
            QueryBuilder1.MetadataContainer.ImportFromXML("Northwind.xml")
            QueryBuilder1.InitializeDatabaseSchemaTree()
            QueryBuilder1.SQL = "select OrderID, CustomerID, OrderDate from Orders"
        End Sub

        Private Sub QueryBuilder1_OnQueryElementControlCreated(owner As QueryElement, control As IQueryElementControl)
            If Not (TypeOf control Is IQueryColumnListControl) Then
                Return
            End If

            Dim queryColumnListControl As IQueryColumnListControl = DirectCast(control, IQueryColumnListControl)
            Dim dataGridView As DataGrid = DirectCast(queryColumnListControl.DataGrid, DataGrid)

            ' Create custom column

                ' Bind this column to the QueryColumnListItem.CustomData object, which is expected to be a string.
            Dim customColumn As DataGridComboBoxColumn = New DataGridComboBoxColumn() With { _
                .Header = "Custom Column", _
                .Width = New DataGridLength(200), _
                .HeaderStyle = New Style(),
                .ItemsSource = _customValuesProvider, _
                .SelectedItemBinding = New Binding("CustomData") With { _
                    .UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged _
                } _
            }
            customColumn.HeaderStyle.Setters.Add(New Setter(FontFamilyProperty, New FontFamily("Tahoma")))
            customColumn.HeaderStyle.Setters.Add( New Setter(FontWeightProperty, FontWeights.Bold))

            ' Insert new column to the specified position
            dataGridView.Columns.Insert(2, customColumn)

            ' Handle the necessary events
            AddHandler dataGridView.BeginningEdit, AddressOf DataGridView_BeginningEdit
            AddHandler dataGridView.CellEditEnding, AddressOf DataGridView_CellEditEnding
            AddHandler dataGridView.LoadingRow, AddressOf dataGridView_LoadingRow
        End Sub

        Private Sub dataGridView_LoadingRow(sender As Object, e As DataGridRowEventArgs)
            AddHandler e.Row.Loaded, AddressOf row_Loaded
        End Sub

        ''' <summary>
        ''' Defining the initial value for a newly added row.
        ''' </summary>
        Private Sub row_Loaded(sender As Object, e As RoutedEventArgs)
            Dim row As DataGridRow = DirectCast(sender, DataGridRow)
            RemoveHandler row.Loaded, AddressOf row_Loaded

            Dim grid As DataGrid = FindParent(Of DataGrid)(row)

            If row.GetIndex() >= grid.Items.Count - 1 Then
                Return
            End If

            Dim item As QueryColumnListItem = TryCast(row.Item, QueryColumnListItem)
            If item Is Nothing Then
                Return
            End If

            ' Initial setting of the custom column data
            item.CustomData = _customValues(row.GetIndex())
        End Sub

        ''' <summary>
        ''' This handler is fired when the user finishes the editing of a cell. 
        ''' </summary>
        Private Sub DataGridView_CellEditEnding(sender As Object, e As DataGridCellEditEndingEventArgs)
            If e.Column.DisplayIndex <> 2 Then
                Return
            End If

            Dim oldValue As String = _customValuesProvider(e.Row.GetIndex())

            ' get new value
            Dim item As QueryColumnListItem = TryCast(e.Row.Item, QueryColumnListItem)

            If item Is Nothing Then
                Return
            End If

            Dim newValue As String = TryCast(item.CustomData, String)

            ' do something with the new value
            If oldValue <> newValue Then
                _customValuesProvider(e.Row.GetIndex()) = newValue
            End If
        End Sub

        ''' <summary>
        ''' This handler determines whether a custom column should be editable or not.
        ''' </summary>
        Private Shared Sub DataGridView_BeginningEdit(sender As Object, e As DataGridBeginningEditEventArgs)
            Dim grid As DataGrid = DirectCast(sender, DataGrid)
            If e.Column.DisplayIndex = 2 AndAlso e.Row.GetIndex() < grid.Items.Count - 1 Then
                    ' Change to "true" if you need a read-only cell.    
                e.Cancel = False
            End If
        End Sub

        Public Shared Function FindParent(Of T As Class)(from As DependencyObject) As T
            Dim result As T = Nothing
            Dim parent As Object = VisualTreeHelper.GetParent(from)

            If TypeOf parent Is T Then
                result = TryCast(parent, T)
            ElseIf parent IsNot Nothing Then
                result = FindParent(Of T)(TryCast(parent, DependencyObject))
            End If

            Return result
        End Function
    End Class
End Namespace
