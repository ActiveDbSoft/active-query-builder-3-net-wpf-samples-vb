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
Imports System.Linq
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.QueryView

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Dim _lastValidSql as String
    Dim _errorPosition as Integer = -1

    Public Sub New()
        InitializeComponent()

        AddHandler Loaded, AddressOf MainWindow_Loaded
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

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
        ErrorBox.Visibility = Visibility.Collapsed
        SqlEditor.Text = QBuilder.FormattedSQL
        _lastValidSql = QBuilder.FormattedSQL
    End Sub

    Private Sub SqlEditor_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Feed the text from text editor to the query builder when user exits the editor.
            QBuilder.SQL =SqlEditor.Text
            ErrorBox.Visibility = Visibility.Collapsed
        Catch ex As SQLParsingException
            ' Set caret to error position
            SqlEditor.CaretIndex = ex.ErrorPos.pos
            ' Report error
            ErrorBox.Show(ex.Message, QBuilder.SyntaxProvider)
            _errorPosition = ex.ErrorPos.pos
        End Try
    End Sub

    ''' <summary>
    ''' Prompts the user if he/she really wants to add an object to the query.
    ''' </summary>
    Private Sub QBuilder_OnDataSourceAdding(metadataObject As MetadataObject, ByRef abort As Boolean)
        If CbDataSourceAdding.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "DataSourceAdding: adding object """ & metadataObject.Name & """" & Environment.NewLine & BoxLogEvents.Text

        Dim name As String = metadataObject.Name
        Dim msg As String = "Are you sure to add """ & name & """ to the query?"

        If MessageBox.Show(msg, "DataSourceAdding event handler", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If

    End Sub

    ''' <summary>
    ''' Displays a prompt on deleting an object from the query.
    ''' </summary>
    Private Sub QBuilder_OnDataSourceDeleting(datasource As DataSource, ByRef abort As Boolean)
        If CbDataSourceDeleting.IsChecked <> True Then
            Return
        End If
        Dim name As String = datasource.NameInQuery
        Dim msg As String = "Are you sure to remove """ & name & """ from the query?"

        If MessageBox.Show(msg, "DataSourceDeleting event handler", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    ''' <summary>
    ''' Prompts the user if he/she really wants to add a field to the SELECT list.
    ''' </summary>
    Private Sub QBuilder_OnDataSourceFieldAdding(datasource As DataSource, field As MetadataField, ByRef abort As Boolean)
        If CbDataSourceFieldAdding.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "DatasourceFieldAdding adding field """ & field.Name & """" & Environment.NewLine & BoxLogEvents.Text

        Dim msg As String = "Are you sure to add """ & field.Name & """ to the SELECT list?"

        If MessageBox.Show(msg, "DatasourceFieldAdding event handler", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    ' Displays a prompt on deleting a field from the SELECT list.
    Private Sub QBuilder_OnDataSourceFieldRemoving(datasource As DataSource, field As MetadataField, ByRef abort As Boolean)
        If CbDataSourceFieldRemoving.IsChecked <> True Then
            Return
        End If
        BoxLogEvents.Text = "DataSourceFieldRemoving removing field """ & field.Name & """ form """ & datasource.NameInQuery & """" & Environment.NewLine & BoxLogEvents.Text
        Dim msg As String = "Are you sure to remove """ & """.""" & datasource.NameInQuery & """" & field.Name & """ from the query?"

        If MessageBox.Show(msg, "DataSourceFieldRemoving event handler", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    '
    ' GridCellValueChanging event allows to prevent the cell value changing or modify the new cell value.
    ' Note: Some columns hide/unhide dynamically but this does not affect the column index in the event parameters -
    '       it includes hidden columns.
    '
    Private Sub QBuilder_OnQueryColumnListItemChanging(querycolumnlist As QueryColumnList, querycolumnlistitem As QueryColumnListItem, [property] As QueryColumnListItemProperty, conditionindex As Integer, oldvalue As Object, ByRef newValue As Object,
        ByRef abort As Boolean)
        If CbQueryColumnListItemChanging.IsChecked <> True Then
            Return
        End If
        BoxLogEvents.Text = "QueryColumnListItemChanging Changing the value for  property """ & Convert.ToString([property]) & """" & Environment.NewLine & BoxLogEvents.Text

        If [property] = QueryColumnListItemProperty.Expression Then
            ' Prevent changes in the Expression column.

            abort = True
            Return
        End If

        If [property] <> QueryColumnListItemProperty.[Alias] Then
            Return
        End If

        Dim s As String = TryCast(newValue, String)
        If s Is Nothing Then
            Return
        End If

        If s.Length <= 0 OrElse s.StartsWith("_") Then
            Return
        End If

        s = "_" & s
        newValue = s
    End Sub

    '
    ' LinkCreating event allows to prevent link creation
    '
    Private Sub QBuilder_OnLinkCreating(fromDatasource As DataSource, fromField As MetadataField, toDatasource As DataSource, toField As MetadataField, correspondingmetadataforeignkey As MetadataForeignKey, ByRef abort As Boolean)
        If CbLinkCreating.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "LinkCreating" & Environment.NewLine & BoxLogEvents.Text

        Dim msg As String = String.Format("Do you want to create the link between {0}.{1} and {2}.{3}?", fromDatasource.MetadataObject.Name, fromField.Name, toDatasource.MetadataObject.Name, toField.Name)

        If MessageBox.Show(msg, "LinkCreating event handler", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    '
    ' LinkDeleting event allows to prevent link deletion.
    '
    Private Sub QBuilder_OnLinkDeleting(link As Link, ByRef abort As Boolean)
        If CbLinkDeleting.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "LinkDeleting" & Environment.NewLine & BoxLogEvents.Text

        Dim msg As String = String.Format("Do you want to delete the link between {0} and {1}?", link.LeftField, link.RightField)

        If MessageBox.Show(msg, "LinkDeleting event handler", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    Private Sub QBuilder_OnQueryColumnListItemChanged(querycolumnlist As QueryColumnList, querycolumnlistitem As QueryColumnListItem, [property] As QueryColumnListItemProperty, conditionindex As Integer, newvalue As Object)
        If CbQueryColumnListItemChanged.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "QueryColumnListItemChanged property """ & Convert.ToString([property]) & """ changed" & Environment.NewLine & BoxLogEvents.Text
    End Sub

    Private Sub QBuilder_OnLinkChanged(link As Link)
        If CbLinkChanged.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "LinkChanged" & Environment.NewLine & BoxLogEvents.Text
    End Sub

    ''' <summary>
    ''' Denies changing of the link properties by the user.
    ''' </summary>
    Private Sub QBuilder_OnLinkChanging(link As Link, ByRef abort As Boolean)
        If CbLinkCreated.IsChecked = True Then
            BoxLogEvents.Text = """LinkChanging"" deny" & Environment.NewLine & BoxLogEvents.Text

            abort = True
            Return
        End If

        BoxLogEvents.Text = """LinkChanging"" allow" & Environment.NewLine & BoxLogEvents.Text
    End Sub

    Private Sub QBuilder_OnDatasourceFieldRemoved(datasource As DataSource, field As MetadataField)
        If CbDatasourceFieldRemoved.IsChecked <> True Then
            Return
        End If

        Dim fieldText As String = if(field Is Nothing, datasource.NameInQuery & ".*", field.Name)

        BoxLogEvents.Text = "DatasourceFieldRemoved " & fieldText & Environment.NewLine & BoxLogEvents.Text
    End Sub

    Private Sub QBuilder_OnDataSourceFieldAdded(datasource As DataSource, field As MetadataField, querycolumnlistitem As QueryColumnListItem, isFocused As Boolean)
        If CbDataSourceFieldAdded.IsChecked <> True Then
            Return
        End If

        Dim fieldText As String = if(field Is Nothing, datasource.NameInQuery & ".*", field.Name)

        BoxLogEvents.Text = "Datasource field " & fieldText & Environment.NewLine & BoxLogEvents.Text
    End Sub

    Private Sub QBuilder_OnDataSourceAdded(sender As SQLQuery, addedobject As DataSource)
        If CbDataSourceAdded.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "DataSourceAdded " & addedobject.NameInQuery & Environment.NewLine & BoxLogEvents.Text
    End Sub

    Private Sub QBuilder_OnLinkCreated(sender As SQLQuery, link As Link)
        If CbLinkCreated.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "LinkCreated" & Environment.NewLine & BoxLogEvents.Text
    End Sub

    Private Sub QBuilder_OnQueryColumnListItemAdded(sender As Object, item As QueryColumnListItem)
        BoxLogEvents.Text = "QueryColumnListItemAdded [" & item.ExpressionString & "]" & Environment.NewLine & BoxLogEvents.Text
    End Sub

    Private Sub QBuilder_OnQueryColumnListItemRemoving(sender As Object, item As QueryColumnListItem, ByRef abort As Boolean)
        If CbQclRemoving.IsChecked <> True Then
            Return
        End If

        BoxLogEvents.Text = "QueryColumnListItemRemoving [" & item.ExpressionString & "]" & Environment.NewLine & BoxLogEvents.Text

        Dim answer As MessageBoxResult = MessageBox.Show(Me, "Do you want to delete the QueryColumnListItem [" & item.ExpressionString & "]", "QueryColumnListItem removing", MessageBoxButton.YesNo)

        abort = answer = MessageBoxResult.No
    End Sub

    Private Sub ButtonCheckAll_OnClick(sender As Object, e As RoutedEventArgs)
        SwitchCheckBox(True)
    End Sub

    Private Sub ButtonUncheckAll_OnClick(sender As Object, e As RoutedEventArgs)
        SwitchCheckBox(False)
    End Sub

    Private Sub SwitchCheckBox(isChecked As Boolean)
        Dim collection As List(Of CheckBox) = FindVisualChildren(Of CheckBox)(CheckBoxPanel).ToList()

        For Each checkBox As CheckBox In collection
            checkBox.IsChecked = isChecked
        Next
    End Sub

    Public Shared Iterator Function FindVisualChildren(Of T As DependencyObject)(depObj As DependencyObject) As IEnumerable(Of T)
        If depObj Is Nothing Then
            Yield Nothing
        End If

        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(depObj) - 1
            Dim child As DependencyObject = VisualTreeHelper.GetChild(depObj, i)
            If child IsNot Nothing AndAlso TypeOf child Is T Then
                Yield CType(child, T)
            End If

            For Each childOfChild As T In FindVisualChildren(Of T)(child)
                Yield childOfChild
            Next
        Next
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
