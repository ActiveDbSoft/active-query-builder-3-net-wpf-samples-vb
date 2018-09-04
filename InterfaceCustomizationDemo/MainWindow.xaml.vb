'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View
Imports ActiveQueryBuilder.View.NavigationBar
Imports ActiveQueryBuilder.View.QueryView
Imports ActiveQueryBuilder.View.WPF.QueryView
Imports QueryColumnListItem = ActiveQueryBuilder.Core.QueryColumnListItem

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Public Partial Class MainWindow
	Private _errorPopup As Popup

	Public Sub New()
		InitializeComponent()

		AddHandler Loaded, AddressOf MainWindow_Loaded
	End Sub

	Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
		RemoveHandler Loaded, AddressOf MainWindow_Loaded
		_errorPopup = New Popup()

		' set syntax provider
		QBuilder.SyntaxProvider = New MSSQLSyntaxProvider()

		' Fill metadata container from the XML file. (For demonstration purposes.)
		Try
			QBuilder.MetadataLoadingOptions.OfflineMode = True
			QBuilder.MetadataContainer.ImportFromXML("Northwind.xml")
			QBuilder.InitializeDatabaseSchemaTree()

			QBuilder.SQL = "SELECT Orders.OrderID, Orders.CustomerID, Orders.OrderDate, [Order Details].ProductID," & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "[Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount" & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "  FROM Orders INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID" & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "  WHERE Orders.OrderID > 0 AND [Order Details].Discount > 0"
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try
	End Sub

	Private Sub QBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
		' Update the text of SQL query when it's changed in the query builder.
		SqlEditor.Text = QBuilder.FormattedSQL
	End Sub

	Private Sub SqlEditor_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
		Try
			' Feed the text from text editor to the query builder when user exits the editor.
			QBuilder.SQL = SqlEditor.Text
			ShowErrorBanner(DirectCast(sender, FrameworkElement), "")
		Catch ex As SQLParsingException
			' Set caret to error position
			SqlEditor.SelectionStart = ex.ErrorPos.pos
			' Report error
			ShowErrorBanner(DirectCast(sender, FrameworkElement), ex.Message)
		End Try
	End Sub

	Public Sub ShowErrorBanner(control As FrameworkElement, text As String)
		' Display error banner if passed text is not empty
		If String.IsNullOrEmpty(text) Then
			If _errorPopup.IsOpen Then
				_errorPopup.IsOpen = False
			End If
			Return
		End If

		Dim label As TextBlock = New TextBlock() With {
			.Text = text, _
			.Background = Brushes.LightPink, _
			.Padding = New Thickness(5) _
		}

		_errorPopup.PlacementTarget = control

		_errorPopup.Child = label
		_errorPopup.IsOpen = True
		_errorPopup.HorizontalOffset = control.ActualWidth - label.ActualWidth - 2
		_errorPopup.VerticalOffset = control.ActualHeight - label.ActualHeight - 2
		Dim timer As Timer= New Timer(AddressOf CallBackPopup, Nothing, 3000, Timeout.Infinite)
	End Sub

	Private Sub CallBackPopup(state As Object)
		Dispatcher.BeginInvoke(DirectCast(Sub() If _errorPopup.IsOpen Then _errorPopup.IsOpen = False, Action))
	End Sub

	Private Sub QBuilder_OnQueryElementControlCreated(owner As QueryElement, control As IQueryElementControl)
		If QElementCreated.IsChecked <> True Then
			Return
		End If

		BoxLogEvents.Text = "QueryElementControl Created """ & control.[GetType]().Name & """" & Environment.NewLine & BoxLogEvents.Text

		If Not (TypeOf control Is DataSourceControl) Then
			Return
		End If

		Dim dsc As DataSourceControl = DirectCast(control, DataSourceControl)
		Dim element As UserControl = DirectCast(control, UserControl)

		Dim root As Grid = DirectCast(element.FindName("PART_CONTENT"), Grid)
		' Grid
		Dim list As ListView = DirectCast(element.FindName("ListBoxView"), ListView)
		' ListView
		If root Is Nothing OrElse list Is Nothing Then
			Return
		End If

		Dim customItemTemplate As DataTemplate = DirectCast(FindResource("CustomFieldTemplate"), DataTemplate)

		root.RowDefinitions.Add(New RowDefinition() With {
			.Height = GridLength.Auto _
		})
		root.RowDefinitions.Add(New RowDefinition() With {
			.Height = New GridLength(1, GridUnitType.Star) _
		})

		list.ItemTemplate = customItemTemplate
		list.SetValue(Grid.RowProperty, 1)

		Dim border As Border = New Border() With {
			.BorderThickness = New Thickness(0, 0, 0, 1), _
			.BorderBrush = Brushes.Gray, _
			.Background = SystemColors.InfoBrush _
		}

		Dim grid1 As Grid = New Grid() With {
			.Margin = New Thickness(3, 0, 3, 0), _
			.Height = 24, _
			.VerticalAlignment = VerticalAlignment.Center _
		}

		border.SetValue(Grid.RowProperty, 0)

		border.Child = grid1
		root.Children.Add(border)

		grid1.ColumnDefinitions.Add(New ColumnDefinition() With {
			.Width = New GridLength(1, GridUnitType.Star) _
		})
		grid1.ColumnDefinitions.Add(New ColumnDefinition() With {
			.Width = GridLength.Auto _
		})

		'#Region "Search box and button"
		Dim textSearchBox As TextBox = New TextBox() With {
			.VerticalContentAlignment = VerticalAlignment.Center, _
			.VerticalAlignment = VerticalAlignment.Center _
		}

		Dim buttonSearchClear As Button = New Button() With {
			.Content = "X", _
			.FontSize = 10, _
			.Padding = New Thickness(5, 0, 5, 0), _
			.VerticalAlignment = VerticalAlignment.Center, _
			.Margin = New Thickness(3, 0, 0, 0), _
			.VerticalContentAlignment = VerticalAlignment.Center, _
			.IsEnabled = False _
		}

		textSearchBox.SetValue(Grid.ColumnProperty, 0)

		' Filtering the list of fields in the DataSourceControl
		AddHandler textSearchBox.TextChanged, Sub() 
		list.Items.Filter = Function(x) 
		Dim dataSourceControlItem As DataSourceControlItem = TryCast(x, DataSourceControlItem)
		Return dataSourceControlItem IsNot Nothing AndAlso dataSourceControlItem.Name.Text.ToLower().Contains(textSearchBox.Text.ToLower())

End Function
		buttonSearchClear.IsEnabled = Not String.IsNullOrWhiteSpace(textSearchBox.Text)

		dsc.Refresh()

End Sub

		grid1.Children.Add(textSearchBox)

		AddHandler buttonSearchClear.Click, Function(sender, args) InlineAssignHelper(textSearchBox.Text, "")

		buttonSearchClear.SetValue(Grid.ColumnProperty, 1)
		'#End Region
		grid1.Children.Add(buttonSearchClear)
	End Sub

	Private Sub QBuilder_OnQueryElementControlDestroying(owner As QueryElement, control As IQueryElementControl)
		If QElementDestroying.IsChecked <> True Then
			Return
		End If

		BoxLogEvents.Text = "QueryElementControl Destroying """ & control.[GetType]().Name & """" & Environment.NewLine & BoxLogEvents.Text
	End Sub

	''' <summary>
	''' ValidateContextMenu event allows to modify or hide any query builder context menu.
	''' </summary>
	Private Sub QBuilder_OnValidateContextMenu(sender As Object, queryelement As QueryElement, menu As ICustomContextMenu)
		If ValidateContextMenu.IsChecked <> True Then
			Return
		End If

		BoxLogEvents.Text = "OnValidateContextMenu" & Environment.NewLine & BoxLogEvents.Text
		' Insert custom menu item to the top of any context menu
		menu.InsertItem(0, "Custom Item 1", AddressOf CustomItem1EventHandler)
		menu.InsertSeparator(1)
		' separator
		If TypeOf queryelement Is Link Then
			' Link context menu
			' Add another item in the Link's menu
			menu.AddSeparator()
			' separator
			menu.AddItem("Custom Item 2", AddressOf CustomItem2EventHandler)
		ElseIf TypeOf queryelement Is DataSourceObject Then
			' Datasource context menu
		ElseIf TypeOf queryelement Is UnionSubQuery Then
			If TypeOf sender Is IDesignPaneControl Then
				' diagram pane context menu
			ElseIf TypeOf sender Is NavigationPopupBase Then
			End If
		ElseIf TypeOf queryelement Is QueryColumnListItem Then
			' QueryColumnListControl context menu
			menu.ClearItems()
		End If
	End Sub

	Private Shared Sub CustomItem1EventHandler(sender As Object, e As EventArgs)
		MessageBox.Show("Custom Item 1")
	End Sub

	Private Shared Sub CustomItem2EventHandler(sender As Object, e As EventArgs)
		MessageBox.Show("Custom Item 2")
	End Sub

	Private Sub QBuilder_OnCustomizeDataSourceCaption(datasource As DataSource, ByRef caption As String)
		If CustomizeDataSourceCaption.IsChecked <> True Then
			Return
		End If

		BoxLogEvents.Text = "CustomizeDataSourceCaption for """ & caption & """" & Environment.NewLine & BoxLogEvents.Text
		caption = caption.ToUpper(CultureInfo.CurrentCulture)
	End Sub

	Private Sub QBuilder_OnCustomizeDataSourceFieldList(datasource As DataSource, fieldlist As List(Of FieldListItemData))
		If CustomizeDataSourceFieldList.IsChecked <> True Then
			Return
		End If

		BoxLogEvents.Text = "CustomizeDataSourceFieldList" & Environment.NewLine & BoxLogEvents.Text
		Dim comparer As FieldComparerByName = New FieldComparerByName()
		fieldlist.Sort(comparer)
	End Sub
	Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
		target = value
		Return value
	End Function
End Class

Public Class FieldComparerByName
	Implements IComparer(Of FieldListItemData)
    Public Function Compare(x As FieldListItemData, y As FieldListItemData) As Integer Implements IComparer(Of FieldListItemData).Compare
        Return String.Compare(x.Name.ToLower(CultureInfo.CurrentCulture), y.Name.ToLower(CultureInfo.CurrentCulture), StringComparison.Ordinal)
    End Function
End Class
