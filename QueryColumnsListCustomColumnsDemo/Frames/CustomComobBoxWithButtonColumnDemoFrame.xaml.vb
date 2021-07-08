''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.QueryView
Imports QueryColumnListItem = ActiveQueryBuilder.View.WPF.QueryView.QueryColumnListItem

Namespace Frames
	''' <summary>
	''' Interaction logic for CustomComobBoxWithButtonColumnDemoFrame.xaml
	''' </summary>
	Public Partial Class CustomComobBoxWithButtonColumnDemoFrame
		Private ReadOnly _customValues As List(Of String)
		Private ReadOnly _customValuesProvider As New List(Of String)()

		Public Sub New()
			InitializeComponent()
			' Fill query builder with demo data
			QueryBuilder1.SyntaxProvider = New MSSQLSyntaxProvider()
			QueryBuilder1.MetadataLoadingOptions.OfflineMode = True
			QueryBuilder1.MetadataContainer.ImportFromXML("Northwind.xml")
			QueryBuilder1.InitializeDatabaseSchemaTree()
			QueryBuilder1.SQL = "select OrderID, CustomerID, OrderDate from Orders"

			' Fill the source of custom values (for the demo purposes)
			For i As Int32 = 0 To 99
				_customValuesProvider.Add("Some Value " & i)
			Next
			_customValues = _customValuesProvider.ToList()
		End Sub

		Private Sub QueryBuilder1_OnQueryElementControlCreated(owner As QueryElement, control As IQueryElementControl)
			If Not (TypeOf control Is IQueryColumnListControl) Then
				Return
			End If

			Dim queryColumnListControl As IQueryColumnListControl = DirectCast(control, IQueryColumnListControl)
			Dim dataGridView As DataGrid = DirectCast(queryColumnListControl.DataGrid, DataGrid)

			' Create binding and templates for the custom column
			Dim textBlock1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(TextBlock))
			textBlock1.SetValue(VerticalAlignmentProperty, VerticalAlignment.Center)
			textBlock1.SetValue(MarginProperty, New Thickness(2, 0, 0, 0))
			textBlock1.SetBinding(TextBlock.TextProperty, New Binding("CustomData") With { _
				.Mode = BindingMode.OneWay, _
				.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged _
			})


			' Creating a template to browse a cell
			Dim templateCell As DataTemplate = New DataTemplate() With { _
				.VisualTree = textBlock1 _
			}

			Dim columnLeft As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
			columnLeft.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Star))

			Dim columnRight As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
			columnRight.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))

			Dim gridRoot As FrameworkElementFactory = New FrameworkElementFactory(GetType(Grid))
			gridRoot.SetValue(BackgroundProperty, Brushes.White)
			gridRoot.AppendChild(columnLeft)
			gridRoot.AppendChild(columnRight)

			Dim comboBox As FrameworkElementFactory = New FrameworkElementFactory(GetType(ComboBox))
			comboBox.SetBinding(Selector.SelectedItemProperty, New Binding("CustomData") With { _
				.Mode = BindingMode.TwoWay, _
				.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged _
			})
			comboBox.SetValue(ItemsControl.ItemsSourceProperty, _customValuesProvider)
			comboBox.SetValue(MarginProperty, New Thickness(0, 0, 0, 0))
			comboBox.SetValue(BorderThicknessProperty, New Thickness(0))
			comboBox.SetValue(VerticalContentAlignmentProperty, VerticalAlignment.Center)
			comboBox.SetValue(Grid.ColumnProperty, 0)
			comboBox.SetValue(PaddingProperty, New Thickness(1, 0, 0, 0))

			' Defining a handler which is fired on loading the editor for a cell
			comboBox.[AddHandler](LoadedEvent, New RoutedEventHandler(AddressOf LoadComboBox))
			gridRoot.AppendChild(comboBox)

			Dim button As FrameworkElementFactory = New FrameworkElementFactory(GetType(Button))
			button.SetValue(Grid.ColumnProperty, 1)
			button.SetValue(MarginProperty, New Thickness(2))
			button.SetValue(ContentProperty, "...")
			button.SetValue(VerticalAlignmentProperty, VerticalAlignment.Center)
			button.SetValue(VerticalContentAlignmentProperty, VerticalAlignment.Center)
			button.SetValue(WidthProperty, 16.0)

			' Defining a handler of cliking the ellipsis button in a cell
			button.[AddHandler](ButtonBase.ClickEvent, New RoutedEventHandler(AddressOf ClickButton))
			gridRoot.AppendChild(button)

			' Creating a template to edit the custom cell
			Dim templateCellEdit As DataTemplate = New DataTemplate() With { _
				.VisualTree = gridRoot _
			}

				' assigning templates to a column
			Dim customColumn As DataGridTemplateColumn = New DataGridTemplateColumn() With {
				.Header = "Custom Column", _
				.Width = New DataGridLength(200), _
				.HeaderStyle = New Style(),
				.CellTemplate = templateCell, _
				.CellEditingTemplate = templateCellEdit _
			}
            customColumn.HeaderStyle.Setters.Add(New Setter(FontFamilyProperty, New FontFamily("Arial")))
            customColumn.HeaderStyle.Setters.Add(New Setter(FontWeightProperty, FontWeights.Bold))

			' Insert new column to the specified position
			dataGridView.Columns.Insert(2, customColumn)

			' Handle the necessary events
			AddHandler dataGridView.BeginningEdit, AddressOf DataGridView_BeginningEdit
			AddHandler dataGridView.CellEditEnding, AddressOf dataGridView_CellEditEnding
			AddHandler dataGridView.LoadingRow, AddressOf dataGridView_LoadingRow
		End Sub

		''' <summary>
		''' The combo-box editor load handler
		''' </summary>
		Private Sub LoadComboBox(sender As Object, e As RoutedEventArgs)
			' opening the drop-down list box on start editing a cell
			Dim comboBox As ComboBox = DirectCast(sender, ComboBox)
			comboBox.IsDropDownOpen = True
		End Sub

		''' <summary>
		''' The ellipsis button click event handler
		''' </summary>
		Private Shared Sub ClickButton(sender As Object, e As RoutedEventArgs)
			Dim row As DataGridRow = FindParent(Of DataGridRow)(DirectCast(sender, FrameworkElement))

			MessageBox.Show(String.Format("Button at row {0} clicked.", row.GetIndex()))
		End Sub

		''' <summary>
		''' This handler is fired when the user finishes the editing of a cell. 
		''' </summary>
		Private Sub dataGridView_CellEditEnding(sender As Object, e As DataGridCellEditEndingEventArgs)
			If e.Column.DisplayIndex <> 2 Then
				Return
			End If

			Dim oldValue As String = _customValues(e.Row.GetIndex())

			' get new value
			Dim item As QueryColumnListItem = TryCast(e.Row.Item, QueryColumnListItem)

			If item Is Nothing Then
				Return
			End If

			Dim newValue As String = item.CustomData.ToString()

			' do something with the new value
			If oldValue <> newValue Then
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

			If row.GetIndex() >= grid.Items.Count - 1 Then
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
		Private Shared Sub DataGridView_BeginningEdit(sender As Object, e As DataGridBeginningEditEventArgs)
			' Checking the possibility to turn a cell into the edit mode.
			Dim grid As DataGrid = DirectCast(sender, DataGrid)
			If e.Column.DisplayIndex = 2 AndAlso e.Row.GetIndex() < grid.Items.Count - 1 Then
				' Make cell editable
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
