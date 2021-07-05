//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.QueryView
Imports CustomColumnsDemo.Annotations
Imports QueryColumnListItem = ActiveQueryBuilder.View.WPF.QueryView.QueryColumnListItem

Namespace Frames
	''' <summary>
	''' Interaction logic for CustomTextColumnDemoFrame.xaml
	''' </summary>
	Public Partial Class CustomTextColumnDemoFrame
		Private ReadOnly _customValuesProvider As New ObservableCollection(Of CustomTextData)()

		Public Sub New()
			InitializeComponent()

			' Fill the query builder with demo data
			QueryBuilder1.SyntaxProvider = New MSSQLSyntaxProvider()
			QueryBuilder1.MetadataLoadingOptions.OfflineMode = True
			QueryBuilder1.MetadataContainer.ImportFromXML("Northwind.xml")
			QueryBuilder1.InitializeDatabaseSchemaTree()
			QueryBuilder1.SQL = "select OrderID, CustomerID, OrderDate from Orders"

			' Fill the source of custom values (for the demo purposes)
			For i As Int32 = 0 To 99
				_customValuesProvider.Add(New CustomTextData() With { _
					.Description = "Some Value " & i _
				})
			Next
		End Sub

		Private Sub QueryBuilder1_OnQueryElementControlCreated(queryElement As QueryElement, queryElementControl As IQueryElementControl)
			If Not (TypeOf queryElementControl Is IQueryColumnListControl) Then
				Return
			End If

			Dim queryColumnListControl As IQueryColumnListControl = DirectCast(queryElementControl, IQueryColumnListControl)
			Dim dataGridView As DataGrid = DirectCast(queryColumnListControl.DataGrid, DataGrid)

			' Create custom column
				' Bind this column to the QueryColumnListItem.CustomData object, which must contain an object with the Description field (defined below)
			Dim customColumn As DataGridTextColumn = New DataGridTextColumn() With { _
				.Header = "Custom Column", _
				.Width = New DataGridLength(200), _
				.HeaderStyle = New Style(),
				.Binding = New Binding("CustomData.Description") With { _
					.Mode = BindingMode.TwoWay, _
					.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged _
				} _
			}
            customColumn.HeaderStyle.Setters.Add(New Setter(FontFamilyProperty, New FontFamily("Tahoma")))
            customColumn.HeaderStyle.Setters.Add(New Setter(FontWeightProperty, FontWeights.Bold))

			' Insert new column to the specified position
			dataGridView.Columns.Insert(2, customColumn)

			' Handle the necessary events
			AddHandler dataGridView.BeginningEdit, AddressOf dataGridView_BeginningEdit
			AddHandler dataGridView.CellEditEnding, AddressOf dataGridView_CellEditEnding
			AddHandler dataGridView.LoadingRow, AddressOf dataGridView_LoadingRow
		End Sub

		''' <summary>
		''' This handler is fired when the user finishes the editing of a cell. 
		''' </summary>
		Private Sub dataGridView_CellEditEnding(sender As Object, e As DataGridCellEditEndingEventArgs)
			If e.Column.DisplayIndex <> 2 Then
				Return
			End If

			Dim oldValue As CustomTextData = _customValuesProvider(e.Row.GetIndex())

			' get new value
			Dim item As QueryColumnListItem = TryCast(e.Row.Item, ActiveQueryBuilder.View.WPF.QueryView.QueryColumnListItem)

			If item Is Nothing Then
				Return
			End If

			Dim newValue As CustomTextData = TryCast(item.CustomData, CustomTextData)

			' do something with the new value
			If oldValue.Description <> newValue.Description Then
				_customValuesProvider(e.Row.GetIndex()) = newValue
			End If
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
			Dim exists As TextBlock = TryCast(grid.Columns(2).GetCellContent(row), TextBlock)

			If (row.GetIndex() >= grid.Items.Count - 1) OrElse exists Is Nothing Then
				Return
			End If

			Dim item As QueryColumnListItem = TryCast(row.Item, QueryColumnListItem)
			If item Is Nothing Then
				Return
			End If

			' Initial setting of the custom column data
			item.CustomData = _customValuesProvider(row.GetIndex())
		End Sub

		''' <summary>
		''' This handler determines whether a custom column should be editable or not.
		''' </summary>
		Private Shared Sub dataGridView_BeginningEdit(sender As Object, e As DataGridBeginningEditEventArgs)
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

        Private Class CustomTextData
			Implements INotifyPropertyChanged
			Private _description As String

			Public Property Description() As String
				Get
					Return _description
				End Get
				Set
					_description = value
					OnPropertyChanged("Description")
				End Set
			End Property

			<NotifyPropertyChangedInvocator> _
			Private Sub OnPropertyChanged(propertyName As String)
				RaiseEvent INotifyPropertyChanged_PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
			End Sub

			Public Event INotifyPropertyChanged_PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
		End Class
	End Class


End Namespace
