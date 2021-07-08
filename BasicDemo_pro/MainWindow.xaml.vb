''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Threading
Imports System.Windows.Markup
Imports GeneralAssembly.Windows.QueryInformationWindows
Imports BasicDemo.ConnectionWindow
Imports Microsoft.Win32
Imports Helpers = ActiveQueryBuilder.View.Helpers
Imports Timer = System.Timers.Timer
Imports BuildInfo = ActiveQueryBuilder.Core.BuildInfo
Imports QueryBuilderPropertiesWindow = GeneralAssembly.QueryBuilderProperties.QueryBuilderPropertiesWindow
Imports Brushes = System.Windows.Media.Brushes
Imports Image = System.Windows.Controls.Image
Imports GeneralAssembly

''' <summary>
''' Interaction logic for Window1.xaml
''' </summary>
Partial Public Class MainWindow
    Private ReadOnly _openFileDialog As New OpenFileDialog()
    Private ReadOnly _saveFileDialog As New SaveFileDialog()
    Private _errorPosition As Integer = -1
    Private _lastValidSql As String

    Public Sub New()
        InitializeComponent()

        _openFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        _saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"

        sqlTextEditor1.QueryProvider = QueryBuilder
        AddHandler QueryBuilder.SleepModeChanged, AddressOf QueryBuilder_SleepModeChanged
        AddHandler QueryBuilder.QueryAwake, AddressOf QueryBuilder_QueryAwake

        dataGridView1.SqlQuery = QueryBuilder.SQLQuery
        dataGridView1.SqlGenerationOptions = QueryBuilder.SQLGenerationOptions

        ' DEMO WARNING
        If BuildInfo.GetEdition() = BuildInfo.Edition.Trial Then
            Dim grid = New Grid()

            Dim trialNoticePanel = New Border With {
                .BorderBrush = Brushes.Black,
                .BorderThickness = New Thickness(1),
                .Background = Brushes.LightGreen,
                .Padding = New Thickness(5),
                .Margin = New Thickness(0, 0, 0, 2)
            }
            trialNoticePanel.SetValue(System.Windows.Controls.Grid.RowProperty, 0)

            Dim label = New TextBlock With {
                .Text = "Generation of random aliases for the query output columns is the limitation of the trial version. The full version is free from this behavior.",
                .HorizontalAlignment = HorizontalAlignment.Left,
                .VerticalAlignment = VerticalAlignment.Top
            }

            Dim button = New Button With {
                .Background = Brushes.Transparent,
                .Padding = New Thickness(0),
                .BorderThickness = New Thickness(0),
                .Cursor = Cursors.Hand,
                .Margin = New Thickness(0, 0, 5, 0),
                .HorizontalAlignment = HorizontalAlignment.Right,
                .VerticalAlignment = VerticalAlignment.Center,
                .Content = New Image With {
                    .Source = My.Resources.cancel.GetImageSource(),
                    .Stretch = Stretch.None
                }
            }

            AddHandler button.Click, Sub()
                                         PanelNotifications.Children.Remove(grid)
                                     End Sub

            button.SetValue(System.Windows.Controls.Grid.RowProperty, 0)

            trialNoticePanel.Child = label
            grid.Children.Add(trialNoticePanel)
            grid.Children.Add(button)

            PanelNotifications.Children.Add(grid)
        End If
    End Sub

    Private Sub QueryBuilder_QueryAwake(sender As QueryRoot, ByRef abort As Boolean)
        If MessageBox.Show("You had typed something that is not a SELECT statement in the text editor and continued with visual query building." & "Whatever the text in the editor is, it will be replaced with the SQL generated by the component. Is it right?", "Active Query Builder .NET Demo", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    Private Sub QueryBuilder_SleepModeChanged(sender As Object, e As EventArgs)
        BorderSleepMode.Visibility = If(QueryBuilder.SleepMode, Visibility.Visible, Visibility.Collapsed)
        tbData.IsEnabled = Not QueryBuilder.SleepMode
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        QueryBuilder.SyntaxProvider = New GenericSyntaxProvider()
        QueryBuilder.SQLQuery.QueryRoot.AllowSleepMode = True
        Dim currentLang = Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName

        Language = XmlLanguage.GetLanguage(If(Helpers.Localizer.Languages.Contains(currentLang.ToLower()), currentLang, "en"))
    End Sub

    Private Sub menuItemAbout_Click(sender As Object, e As RoutedEventArgs)
        QueryBuilder.ShowAboutDialog()
    End Sub

    ' TextBox lost focus by keyboard
    Private Sub SqlTextEditor_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QueryBuilder.SQL = sqlTextEditor1.Text
            ErrorBox.Show(Nothing, QueryBuilder.SyntaxProvider)
            _lastValidSql = QueryBuilder.FormattedSQL
        Catch ex As SQLParsingException
            ' Set caret to error position
            sqlTextEditor1.SelectionStart = ex.ErrorPos.pos
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

    Private Sub TimerEvent(source As Object, args As EventArgs)
        Dispatcher?.Invoke(Sub()
                               panel1.Visibility = Visibility.Collapsed
                           End Sub)

        DirectCast(source, Timer).Stop()
        DirectCast(source, Timer).Dispose()
    End Sub

    Private Sub RefreshMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Force the query builder to refresh metadata from current connection
        ' to refresh metadata, just clear MetadataContainer and reinitialize metadata tree

        If QueryBuilder.MetadataProvider Is Nothing OrElse Not QueryBuilder.MetadataProvider.Connected Then
            Return
        End If

        QueryBuilder.ClearMetadata()
        QueryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub ClearMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Clear the metadata

        If MessageBox.Show("Clear Metadata Container?", "Confirmation", MessageBoxButton.YesNo) = MessageBoxResult.Yes Then
            QueryBuilder.ClearMetadata()
        End If
    End Sub

    Private Sub LoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Load metadata from XML file
        If Not _openFileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        QueryBuilder.MetadataLoadingOptions.OfflineMode = True
        QueryBuilder.MetadataContainer.ImportFromXML(_openFileDialog.FileName)
        QueryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub SaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Save metadata to XML file
        _saveFileDialog.FileName = "Metadata.xml"

        If Not _saveFileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        QueryBuilder.MetadataContainer.LoadAll(True)
        QueryBuilder.MetadataContainer.ExportToXML(_saveFileDialog.FileName)
    End Sub

    Private Sub QueryBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
        ' Handle the event raised by SQL Builder object that the text of SQL query is changed
        ' update the text box
        sqlTextEditor1.Text = QueryBuilder.FormattedSQL
        CheckParameters()

        If Not Equals(TabControl.SelectedItem, tbData) Then
            Return
        End If

        ExecuteQuery()
    End Sub

    Private Sub CheckParameters()
        If SqlHelpers.CheckParameters(QueryBuilder.MetadataProvider, QueryBuilder.SyntaxProvider, QueryBuilder.Parameters) Then
            HideParametersErrorPanel()
        Else
            Dim acceptableFormats = SqlHelpers.GetAcceptableParametersFormats(QueryBuilder.MetadataProvider, QueryBuilder.SyntaxProvider)
            ShowParametersErrorPanel(acceptableFormats)
        End If
    End Sub

    Private Sub ShowParametersErrorPanel(acceptableFormats As List(Of String))
        Dim formats = acceptableFormats.Select(Function(x)
                                                   Dim s = x.Replace("n", "<number>")
                                                   Return s.Replace("s", "<name>")
                                               End Function)

        lbParamsError.Text = "Unsupported parameter notation detected. For this type of connection and database server use " & String.Join(", ", formats)
        pnlParamsError.Visibility = Visibility.Visible
    End Sub

    Private Sub HideParametersErrorPanel()
        pnlParamsError.Visibility = Visibility.Collapsed
    End Sub

    Public Sub ResetQueryBuilder()
        QueryBuilder.ClearMetadata()
        QueryBuilder.MetadataProvider = Nothing
        QueryBuilder.SyntaxProvider = Nothing
        QueryBuilder.MetadataLoadingOptions.OfflineMode = False
    End Sub

    Private Sub connect_OnClick(sender As Object, e As RoutedEventArgs)
        Dim form = New ConnectionEditWindow() With {.Owner = Me}
        Dim result = form.ShowDialog()
        If result.HasValue AndAlso result.Value Then
            Try
                QueryBuilder.SQLContext.Assign(form.Connection.GetSqlContext())
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub FillProgrammatically_OnClick(sender As Object, e As RoutedEventArgs)
        ResetQueryBuilder()

        ' Fill the query builder metadata programmatically

        ' setup the query builder with metadata and syntax providers
        QueryBuilder.SyntaxProvider = New GenericSyntaxProvider()
        QueryBuilder.MetadataLoadingOptions.OfflineMode = True ' prevent querying obejects from database

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

    Private Sub QueryStatistic_OnClick(sender As Object, e As RoutedEventArgs)
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

    Private Sub Properties_OnClick(sender As Object, e As RoutedEventArgs)
        ' Show Properties form
        Dim f = New QueryBuilderPropertiesWindow(QueryBuilder) With {.Owner = Me}

        f.ShowDialog()
        WarnAboutGenericSyntaxProvider() ' show warning (just for demonstration purposes)
    End Sub

    Private Sub ExecuteQuery()
        If sqlTextEditor1.Text <> QueryBuilder.FormattedSQL Then
            QueryBuilder.SQL = sqlTextEditor1.Text
        End If

        dataGridView1.FillData(QueryBuilder.SQL)
    End Sub

    Private Sub Selector_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If Not Equals(TabControl.SelectedItem, tbData) Then
            Return
        End If

        ExecuteQuery()
    End Sub

    Private Sub SqlTextEditor1_OnTextChanged(sender As Object, e As EventArgs)
        ErrorBox.Show(Nothing, QueryBuilder.SyntaxProvider)
    End Sub

    Private Sub MenuItemEditMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        QueryBuilder.EditMetadataContainer(QueryBuilder.SQLContext)
    End Sub

    Private Sub ErrorBox_OnSyntaxProviderChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim oldSql = sqlTextEditor1.Text
        Dim caretPosition = sqlTextEditor1.CaretOffset

        QueryBuilder.SyntaxProvider = DirectCast(e.AddedItems(0), BaseSyntaxProvider)
        sqlTextEditor1.Text = oldSql
        sqlTextEditor1.Focus()
        sqlTextEditor1.CaretOffset = caretPosition
    End Sub

    Private Sub ErrorBox_OnGoToErrorPositionEvent(sender As Object, e As EventArgs)
        sqlTextEditor1.Focus()

        If _errorPosition = -1 Then
            Return
        End If

        sqlTextEditor1.ScrollToPosition(_errorPosition)
        sqlTextEditor1.CaretOffset = _errorPosition
    End Sub

    Private Sub ErrorBox_OnRevertValidTextEvent(sender As Object, e As EventArgs)
        sqlTextEditor1.Text = _lastValidSql
        sqlTextEditor1.Focus()
    End Sub
End Class
