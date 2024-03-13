''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View
Imports ActiveQueryBuilder.View.NavigationBar
Imports ActiveQueryBuilder.View.QueryView
Imports ActiveQueryBuilder.View.WPF
Imports ActiveQueryBuilder.View.WPF.QueryView
Imports QueryColumnListItem = ActiveQueryBuilder.Core.QueryColumnListItem

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private _lastValidSql As String = String.Empty
    Private _errorPosition As Integer = -1

    Public Sub New()
        InitializeComponent()

        AddHandler Loaded, AddressOf MainWindow_Loaded
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

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
        ErrorBox.Visibility = Visibility.Collapsed
        _lastValidSql = QBuilder.FormattedSQL
    End Sub

    Private Sub SqlEditor_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Feed the text from text editor to the query builder when user exits the editor.
            QBuilder.SQL =SqlEditor.Text
            ErrorBox.Visibility = Visibility.Collapsed
            _lastValidSql = SqlEditor.Text
        Catch ex As SQLParsingException
            ' Set caret to error position
            SqlEditor.CaretIndex = ex.ErrorPos.pos
            ' Report error
            ErrorBox.Show(ex.Message, QBuilder.SyntaxProvider)
            _errorPosition = ex.ErrorPos.pos
        End Try
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

        list.ItemTemplate = customItemTemplate
        list.SetValue(Grid.RowProperty, 1)

        Dim border As Border = New Border() With {
            .BorderThickness = New Thickness(0, 0, 0, 1),
            .BorderBrush = Brushes.Gray,
            .Background = SystemColors.InfoBrush
        }

        Dim grid1 As Grid = New Grid() With {
            .Margin = New Thickness(3, 0, 3, 0),
            .Height = 24,
            .VerticalAlignment = VerticalAlignment.Center
        }

        border.SetValue(Grid.RowProperty, 0)

        border.Child = grid1
        root.Children.Add(border)

        grid1.ColumnDefinitions.Add(New ColumnDefinition() With {
            .Width = New GridLength(1, GridUnitType.Star)
        })
        grid1.ColumnDefinitions.Add(New ColumnDefinition() With {
            .Width = GridLength.Auto
        })

        '#Region "Search box and button"
        Dim textSearchBox As TextBox = New TextBox() With {
            .VerticalContentAlignment = VerticalAlignment.Center,
            .VerticalAlignment = VerticalAlignment.Center
        }

        Dim buttonSearchClear As Button = New Button() With {
            .Content = "X",
            .FontSize = 10,
            .Padding = New Thickness(5, 0, 5, 0),
            .VerticalAlignment = VerticalAlignment.Center,
            .Margin = New Thickness(3, 0, 0, 0),
            .VerticalContentAlignment = VerticalAlignment.Center,
            .IsEnabled = False
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
    Private Sub QBuilder_OnValidateContextMenu(ByVal sender As Object, ByVal queryelement As QueryElement, ByVal menu As ICustomContextMenu)
        If ValidateContextMenu.IsChecked <> True Then Return
        BoxLogEvents.Text = "OnValidateContextMenu" & Environment.NewLine + BoxLogEvents.Text
        menu.InsertItem(0, "Custom Item 1", AddressOf CustomItem1EventHandler)
        menu.InsertSeparator(1)

        If TypeOf queryelement Is Link Then
            menu.AddSeparator()
            menu.AddItem("Custom Item 2", AddressOf CustomItem2EventHandler)
        ElseIf TypeOf queryelement Is DataSourceObject Then
        ElseIf TypeOf queryelement Is UnionSubQuery Then

            If TypeOf sender Is IDesignPaneControl Then
            ElseIf TypeOf sender Is NavigationPopupBase Then
            End If
        ElseIf TypeOf queryelement Is QueryColumnListItem Then
            Dim queryColumnListItem As QueryColumnListItem= CType(queryelement, QueryColumnListItem)

            Dim queryColumnListHitTestInfo As QueryColumnListHitTestInfo = QBuilder.QueryColumnListControl.HitTest(Mouse.GetPosition(CType(QBuilder.QueryColumnListControl, UIElement)).ToCPoint())

            Select Case queryColumnListHitTestInfo.ItemProperty
                Case QueryColumnListItemProperty.Expression
                    menu.AddSeparator()
                    Dim menuItemExpression As ICustomSubMenu = menu.AddSubMenu("Expression property")
                    menuItemExpression.AddItem("Show full SQL", AddressOf ExpressionColumnEventHandler, False, True, Nothing, queryColumnListItem.Expression.GetSQL())
                Case QueryColumnListItemProperty.Selected,
                    QueryColumnListItemProperty.[Alias],
                    QueryColumnListItemProperty.SortType,
                    QueryColumnListItemProperty.SortOrder,
                    QueryColumnListItemProperty.Aggregate,
                    QueryColumnListItemProperty.Grouping,
                    QueryColumnListItemProperty.ConditionType,
                    QueryColumnListItemProperty.Condition,
                    QueryColumnListItemProperty.Custom
                    menu.AddSeparator()
                    menu.AddItem("Get info of current cell", Sub(o, args)
                                                                 Dim message As String = $"Item property [{queryColumnListHitTestInfo.ItemProperty}]{Environment.NewLine}Item index [{queryColumnListHitTestInfo.ItemIndex}]{Environment.NewLine}Condition index [{queryColumnListHitTestInfo.ConditionIndex}]{Environment.NewLine}Is now here [{queryColumnListHitTestInfo.IsNowhere}]"
                                                                 MessageBox.Show(Me, message, "Information")
                                                             End Sub)
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select
        End If
    End Sub

    Private Shared Sub ExpressionColumnEventHandler(sender As Object, e As EventArgs)
        Dim menuItem As ICustomMenuItem = CType(sender, ICustomMenuItem)
        MessageBox.Show(menuItem.Tag.ToString())
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

    Private Sub ErrorBox_OnGoToErrorPosition(sender As Object, e As EventArgs)
        SqlEditor.Focus()

        If _errorPosition <> -1 Then
            If SqlEditor.LineCount <> 1 Then SqlEditor.ScrollToLine(SqlEditor.GetLineIndexFromCharacterIndex(_errorPosition))
            SqlEditor.CaretIndex = _errorPosition
        End If
    End Sub

    Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
        SqlEditor.Text = _lastValidSql
        SqlEditor.Focus()
    End Sub
End Class

Public Class FieldComparerByName
    Implements IComparer(Of FieldListItemData)
    Public Function Compare(x As FieldListItemData, y As FieldListItemData) As Integer Implements IComparer(Of FieldListItemData).Compare
        Return String.Compare(x.Name.ToLower(CultureInfo.CurrentCulture), y.Name.ToLower(CultureInfo.CurrentCulture), StringComparison.Ordinal)
    End Function
End Class
