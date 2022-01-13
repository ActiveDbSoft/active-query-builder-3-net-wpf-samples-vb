''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Threading
Imports System.Windows.Markup
Imports GeneralAssembly.Windows.QueryInformationWindows
Imports Microsoft.Win32
Imports GeneralAssembly.QueryBuilderProperties
Imports Helpers = ActiveQueryBuilder.View.Helpers
Imports Timer = System.Timers.Timer
Imports BuildInfo = ActiveQueryBuilder.Core.BuildInfo
Imports Brushes = System.Windows.Media.Brushes
Imports GeneralAssembly
Imports Connection

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private _showHintConnection As Boolean = True
    Private _selectedConnection As ConnectionInfo

    Private _lastValidSql As String
    Private _errorPosition As Integer = -1

    Private ReadOnly _openFileDialog As New OpenFileDialog()
    Private ReadOnly _saveFileDialog As New SaveFileDialog()

    Private ReadOnly _genericSyntaxProvider As GenericSyntaxProvider

    Public Sub New()
        InitializeComponent()

        _genericSyntaxProvider = New GenericSyntaxProvider()

        _openFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        _saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"

        QueryBuilder.SQLQuery.QueryRoot.AllowSleepMode = True
        AddHandler QueryBuilder.QueryAwake, AddressOf QueryBuilder_QueryAwake
        AddHandler QueryBuilder.SleepModeChanged, AddressOf QueryBuilder_SleepModeChanged
        DataGridResult.SqlQuery = QueryBuilder.SQLQuery

        ' DEMO WARNING
        If BuildInfo.GetEdition() = BuildInfo.Edition.Trial Then
            Dim trialNoticePanel = New Border With {
                .BorderBrush = Brushes.Black,
                .BorderThickness = New Thickness(1),
                .Background = Brushes.LightGreen,
                .Padding = New Thickness(5),
                .Margin = New Thickness(0, 0, 0, 2)
            }
            trialNoticePanel.SetValue(Grid.RowProperty, 1)

            Dim label = New TextBlock With {
                .Text = "Generation of random aliases for the query output columns is the limitation of the trial version. The full version is free from this behavior.",
                .HorizontalAlignment = HorizontalAlignment.Left,
                .VerticalAlignment = VerticalAlignment.Top
            }

            trialNoticePanel.Child = label
            GridRoot.Children.Add(trialNoticePanel)
        End If
    End Sub

    Private Sub QueryBuilder_SleepModeChanged(ByVal sender As Object, ByVal e As EventArgs)
        BorderSleepMode.Visibility = If(QueryBuilder.SleepMode, Visibility.Visible, Visibility.Collapsed)
        tbData.IsEnabled = Not QueryBuilder.SleepMode
    End Sub

    Private Sub QueryBuilder_QueryAwake(ByVal sender As QueryRoot, ByRef abort As Boolean)
        If MessageBox.Show("You had typed something that is not a SELECT statement in the text editor and continued with visual query building." & "Whatever the text in the editor is, it will be replaced with the SQL generated by the component. Is it right?", "Active Query Builder .NET Demo", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        QueryBuilder.SyntaxProvider = _genericSyntaxProvider

        Dim currentLang = Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName

        Language = XmlLanguage.GetLanguage(If(Helpers.Localizer.Languages.Contains(currentLang.ToLower()), currentLang, "en"))
        QueryBuilder.SQLQuery.QueryRoot.AllowSleepMode = True
    End Sub

    Private Sub menuItemAbout_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        QueryBuilder.ShowAboutDialog()
    End Sub

    ' TextBox lost focus by keyboard
    Private Sub textBox1_LostKeyboardFocus(ByVal sender As Object, ByVal e As System.Windows.Input.KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QueryBuilder.SQL = TextBox1.Text
            ErrorBox.Show(Nothing, QueryBuilder.SyntaxProvider)
            _lastValidSql = QueryBuilder.FormattedSQL
            _errorPosition = -1
        Catch ex As SQLParsingException
            ' Set caret to error position
            TextBox1.SelectionStart = ex.ErrorPos.pos
            _errorPosition = ex.ErrorPos.pos
            ' Report error
            ErrorBox.Show(ex.Message, QueryBuilder.SyntaxProvider)
        End Try
    End Sub

    Private Sub WarnAboutGenericSyntaxProvider()
        If TypeOf QueryBuilder.SyntaxProvider Is GenericSyntaxProvider Then
            panel1.Visibility = Visibility.Visible

            ' setup the panel to hide automatically
            Dim timer = New Timer()
            AddHandler timer.Elapsed, AddressOf TimerEvent
            timer.Interval = 10000
            timer.Start()
        End If
    End Sub

    Private Sub TimerEvent(ByVal source As Object, ByVal args As EventArgs)
        Dispatcher?.Invoke(Sub()
                               panel1.Visibility = Visibility.Collapsed
                           End Sub)

        DirectCast(source, Timer).Stop()
        DirectCast(source, Timer).Dispose()
    End Sub

    Private Sub RefreshMetadata_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Force the query builder to refresh metadata from current connection
        ' to refresh metadata, just clear MetadataContainer and reinitialize metadata tree

        If QueryBuilder.MetadataProvider Is Nothing OrElse Not QueryBuilder.MetadataProvider.Connected Then
            Return
        End If

        QueryBuilder.ClearMetadata()
        QueryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub ClearMetadata_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Clear the metadata

        If MessageBox.Show("Clear Metadata Container?", "Confirmation", MessageBoxButton.YesNo) = MessageBoxResult.Yes Then
            QueryBuilder.MetadataContainer.Clear()
        End If
    End Sub

    Private Sub LoadMetadata_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Load metadata from XML file
        If Not _openFileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        QueryBuilder.MetadataLoadingOptions.OfflineMode = True
        QueryBuilder.MetadataContainer.ImportFromXML(_openFileDialog.FileName)
        QueryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub SaveMetadata_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Save metadata to XML file
        _saveFileDialog.FileName = "Metadata.xml"

        If Not _saveFileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        QueryBuilder.MetadataContainer.LoadAll(True)
        QueryBuilder.MetadataContainer.ExportToXML(_saveFileDialog.FileName)
    End Sub

    Private Sub QueryBuilder_OnSQLUpdated(ByVal sender As Object, ByVal e As EventArgs)
        ' Handle the event raised by SQL Builder object that the text of SQL query is changed
        ' update the text box
        TextBox1.Text = QueryBuilder.FormattedSQL

        If Not Equals(TabControl.SelectedItem, tbData) Then
            Return
        End If

        ExecuteSql()
    End Sub

    Public Sub ResetQueryBuilder()
        QueryBuilder.ClearMetadata()
        QueryBuilder.MetadataProvider = Nothing
        QueryBuilder.SyntaxProvider = Nothing
        QueryBuilder.MetadataLoadingOptions.OfflineMode = False
    End Sub

    Private Sub FillProgrammatically_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ResetQueryBuilder()

        ' Fill the query builder metadata programmatically

        ' setup the query builder with metadata and syntax providers
        QueryBuilder.SyntaxProvider = _genericSyntaxProvider
        QueryBuilder.MetadataLoadingOptions.OfflineMode = True ' prevent querying objects from database

        ' create database and schema
        Dim database As MetadataNamespace = QueryBuilder.MetadataContainer.AddDatabase("MyDB")
        database.Default = True
        Dim schema As MetadataNamespace = database.AddSchema("MySchema")
        schema.Default = True

        ' create table
        Dim tableOrders As MetadataObject = schema.AddTable("Orders")
        tableOrders.AddField("OrderID")
        tableOrders.AddField("OrderDate")
        tableOrders.AddField("CustomerID")
        tableOrders.AddField("ResellerID")

        ' create another table
        Dim tableCustomers As MetadataObject = schema.AddTable("Customers")
        tableCustomers.AddField("CustomerID")
        tableCustomers.AddField("CustomerName")
        tableCustomers.AddField("CustomerAddress")

        ' add a relation between these two tables
        Dim relation As MetadataForeignKey = tableCustomers.AddForeignKey("FK_CustomerID")
        relation.Fields.Add("CustomerID")
        relation.ReferencedObjectName = tableOrders.GetQualifiedName()
        relation.ReferencedFields.Add("CustomerID")

        'create view
        Dim viewResellers As MetadataObject = schema.AddView("Resellers")
        viewResellers.AddField("ResellerID")
        viewResellers.AddField("ResellerName")

        ' kick the query builder to fill metadata tree
        QueryBuilder.InitializeDatabaseSchemaTree()

        WarnAboutGenericSyntaxProvider() ' show warning (just for demonstration purposes)
    End Sub

    Private Sub QueryStatistic_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim queryStatistics As QueryStatistics = QueryBuilder.QueryStatistics
        Dim builder As New StringBuilder()

        builder.Append("Used Objects (").Append(queryStatistics.UsedDatabaseObjects.Count).AppendLine("):")
        builder.AppendLine()

        For i As Integer = 0 To queryStatistics.UsedDatabaseObjects.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjects(i).ObjectName.QualifiedName)
        Next i

        builder.AppendLine().AppendLine()
        builder.Append("Used Columns (").Append(queryStatistics.UsedDatabaseObjectFields.Count).AppendLine("):")
        builder.AppendLine()

        For i As Integer = 0 To queryStatistics.UsedDatabaseObjectFields.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjectFields(i).FullName.QualifiedName)
        Next i

        builder.AppendLine().AppendLine()
        builder.Append("Output Expressions (").Append(queryStatistics.OutputColumns.Count).AppendLine("):")
        builder.AppendLine()

        For i As Integer = 0 To queryStatistics.OutputColumns.Count - 1
            builder.AppendLine(queryStatistics.OutputColumns(i).Expression)
        Next i

        Dim f = New QueryStatisticsWindow(builder.ToString()) With {.Owner = Me}
        f.ShowDialog()
    End Sub

    Private Sub Properties_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Show Properties form
        Dim f As New QueryBuilderPropertiesWindow(QueryBuilder)

        f.ShowDialog()

        WarnAboutGenericSyntaxProvider() ' show warning (just for demonstration purposes)
    End Sub

    Private Sub ExecuteSql()
        DataGridResult.FillData(QueryBuilder.SQL)
    End Sub

    Private Sub Selector_OnSelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        If e.AddedItems.Count = 0 Then
            Return
        End If
        Dim tab = TryCast(e.AddedItems(0), TabItem)
        If tab Is Nothing Then
            Return
        End If

        If QueryBuilder.FormattedSQL <> TextBox1.Text Then
            QueryBuilder.SQL = TextBox1.Text
        End If

        If Not Equals(TabControl.SelectedItem, tbData) OrElse QueryBuilder.SQL = String.Empty Then

            Return
        End If

        ExecuteSql()
    End Sub

    Private Sub TextBox1_OnTextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
        ErrorBox.Show(Nothing, QueryBuilder.SyntaxProvider)
    End Sub

    Private Sub MenuItemEditMetadata_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        QueryBuilder.EditMetadataContainer(QueryBuilder.SQLContext)
    End Sub

    Private Sub ErrorBox_OnSyntaxProviderChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        Dim oldSql = TextBox1.Text
        Dim caretPosition = TextBox1.CaretIndex

        QueryBuilder.SyntaxProvider = DirectCast(e.AddedItems(0), BaseSyntaxProvider)
        TextBox1.Text = oldSql
        TextBox1.Focus()
        TextBox1.CaretIndex = caretPosition
    End Sub

    Private Sub ErrorBox_OnGoToErrorPositionEvent(ByVal sender As Object, ByVal e As EventArgs)
        TextBox1.Focus()

        If _errorPosition = -1 Then
            Return
        End If

        If TextBox1.LineCount <> 1 Then
            TextBox1.ScrollToLine(TextBox1.GetLineIndexFromCharacterIndex(_errorPosition))
        End If
        TextBox1.CaretIndex = _errorPosition
    End Sub

    Private Sub ErrorBox_OnRevertValidTextEvent(ByVal sender As Object, ByVal e As EventArgs)
        TextBox1.Text = _lastValidSql
        TextBox1.Focus()
    End Sub

    Private Sub ConnectTo_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim cf = New DatabaseConnectionWindow(_showHintConnection) With {.Owner = Me}
        _showHintConnection = False
        If cf.ShowDialog() <> True Then
            Return
        End If
        _selectedConnection = cf.SelectedConnection

        InitializeSqlContext()
    End Sub
    Private Sub InitializeSqlContext()
        Try
            QueryBuilder.Clear()

            Dim metadataProvider As BaseMetadataProvider = Nothing

            If _selectedConnection Is Nothing Then
                Return
            End If

            ' create new SqlConnection object using the connections string from the connection form
            If Not _selectedConnection.IsXmlFile Then
                metadataProvider = _selectedConnection.ConnectionDescriptor?.MetadataProvider
            End If

            ' setup the query builder with metadata and syntax providers
            QueryBuilder.SQLContext.MetadataContainer.Clear()
            QueryBuilder.MetadataProvider = metadataProvider
            QueryBuilder.SyntaxProvider = _selectedConnection.ConnectionDescriptor.SyntaxProvider
            QueryBuilder.MetadataLoadingOptions.OfflineMode = metadataProvider Is Nothing

            If metadataProvider Is Nothing Then
                QueryBuilder.MetadataContainer.ImportFromXML(_selectedConnection.ConnectionString)
            End If

            ' Instruct the query builder to fill the database schema tree
            QueryBuilder.InitializeDatabaseSchemaTree()

        Finally

            DataGridResult.SqlQuery = QueryBuilder.SQLQuery
        End Try
    End Sub
End Class
