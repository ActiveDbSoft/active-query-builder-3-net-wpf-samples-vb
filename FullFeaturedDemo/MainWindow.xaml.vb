'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports Oracle.ManagedDataAccess.Client
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Markup
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.QueryTransformer
Imports FullFeaturedDemo.Connection
Imports FullFeaturedDemo.PropertiesForm
Imports FullFeaturedDemo.Windows
Imports Microsoft.Win32

Imports System.Text.RegularExpressions
Imports System.Windows.Media
Imports ActiveQueryBuilder.View.WPF
Imports MySql.Data.MySqlClient
Imports Npgsql

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private _selectedConnection As ConnectionInfo
    Private ReadOnly _sqlFormattingOptions As SQLFormattingOptions
    Private ReadOnly _sqlGenerationOptions As SQLGenerationOptions
    Private _showHintConnection As Boolean = True
    Private ReadOnly _transformerSql As QueryTransformer
    Private ReadOnly _timerStartingExecuteSql As Timer

    Public Sub New()
        ' Options to present the formatted SQL query text to end-user
        ' Use names of virtual objects, do not replace them with appropriate derived tables
        _sqlFormattingOptions = New SQLFormattingOptions() With {
            .ExpandVirtualObjects = False
        }

        ' Options to generate the SQL query text for execution against a database server
        ' Replace virtual objects with derived tables
        _sqlGenerationOptions = New SQLGenerationOptions() With {
            .ExpandVirtualObjects = True
        }

        InitializeComponent()

        AddHandler Closing, AddressOf MainWindow_Closing
        AddHandler Dispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive

        Dim currentLang As String = Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName

        LoadLanguage()

        Dim defLang As String = "en"

        If ActiveQueryBuilder.Core.Helpers.Localizer.Languages.Contains(currentLang.ToLower()) Then
            Language = XmlLanguage.GetLanguage(currentLang)
            defLang = currentLang.ToLower()
        End If

        Dim menuItem as MenuItem = MenuItemLanguage.Items.Cast(Of MenuItem)().First(Function(item) DirectCast(item.Tag, String) = defLang)
        menuItem.IsChecked = True

        QBuilder.SyntaxProvider = New GenericSyntaxProvider()

        _transformerSql = New QueryTransformer()

        _timerStartingExecuteSql = New Timer(AddressOf TimerStartingExecuteSql_Elapsed)

        ' DEMO WARNING
        If ActiveQueryBuilder.Core.BuildInfo.GetEdition() = ActiveQueryBuilder.Core.BuildInfo.Edition.Trial Then
            Dim trialNoticePanel As Border = New Border() With {
                .BorderBrush = Brushes.Black,
                .BorderThickness = New Thickness(1),
                .Background = Brushes.LightGreen,
                .Padding = New Thickness(5),
                .Margin = New Thickness(0, 0, 0, 2)
            }
            trialNoticePanel.SetValue(Grid.RowProperty, 1)

            Dim label As TextBlock = New TextBlock() With {
                .Text = "Generation of random aliases for the query output columns is the limitation of the trial version. The full version is free from this behavior.",
                .HorizontalAlignment = HorizontalAlignment.Left,
                .VerticalAlignment = VerticalAlignment.Top
            }

            Dim button As Button = New Button() With {
                .Background = Brushes.Transparent,
                .Padding = New Thickness(0),
                .BorderThickness = New Thickness(0),
                .Cursor = Cursors.Hand,
                .Margin = New Thickness(0, 0, 5, 0),
                .HorizontalAlignment = HorizontalAlignment.Right,
                .VerticalAlignment = VerticalAlignment.Center,
                .Content = "X"
            }

            AddHandler button.Click, Sub() GridRoot.Visibility = Visibility.Collapsed

            trialNoticePanel.Child = label
            GridRoot.Children.Add(trialNoticePanel)
            GridRoot.Children.Add(button)
        End If

        QBuilder.SQLQuery.QueryRoot.AllowSleepMode = True

        AddHandler QBuilder.SleepModeChanged, AddressOf SqlQuery_SleepModeChanged
        AddHandler QBuilder.QueryAwake, AddressOf SqlQuery_QueryAwake
    End Sub

    Private Shared Sub SqlQuery_QueryAwake(sender As QueryRoot, ByRef abort As Boolean)
        If MessageBox.Show("You had typed something that is not a SELECT statement in the text editor and continued with visual query building." & "Whatever the text in the editor is, it will be replaced with the SQL generated by the component. Is it right?", "Active Query Builder .NET Demo", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            abort = True
        End If
    End Sub

    Private Sub SqlQuery_SleepModeChanged(sender As Object, e As EventArgs)
        BorderSleepMode.Visibility = If(QBuilder.SleepMode, Visibility.Visible, Visibility.Collapsed)
        TabItemFastResult.Visibility = If(QBuilder.SleepMode, Visibility.Collapsed, Visibility.Visible)
        TabItemCurrentSubQuery.Visibility = If(QBuilder.SleepMode, Visibility.Collapsed, Visibility.Visible)
        TabItemData.IsEnabled = Not QBuilder.SleepMode
    End Sub

    Private Sub LoadLanguage()
        For Each language As String In ActiveQueryBuilder.Core.Helpers.Localizer.Languages
            If language.ToLower() = "auto" OrElse language.ToLower() = "default" Then
                Continue For
            End If

            Dim culture As CultureInfo = New CultureInfo(language)

            Dim stroke as String = String.Format("{0}", culture.DisplayName)

            Dim menuItem as MenuItem = New MenuItem() With {
                .Header = stroke,
                .Tag = language,
                .IsCheckable = True
            }

            MenuItemLanguage.Items.Add(menuItem)
            menuItem.SetValue(MenuBehavior.OptionGroupNameProperty, "group")
        Next
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        RemoveHandler Closing, AddressOf MainWindow_Closing
        RemoveHandler Dispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
    End Sub

    Private Sub Hooks_DispatcherInactive(sender As Object, e As EventArgs)
        MenuItemQueryProperties.IsEnabled = _selectedConnection IsNot Nothing
        MenuItemNewObject.IsEnabled = _selectedConnection IsNot Nothing
        MenuItemAddDerivedTable.IsEnabled = _selectedConnection IsNot Nothing
        MenuItemAddUnionSubquery.IsEnabled = _selectedConnection IsNot Nothing
        MenuItemCopyUnionSubquery.IsEnabled = _selectedConnection IsNot Nothing

        MenuItemSave.IsEnabled = _selectedConnection IsNot Nothing
        MenuItemQueryStatistics.IsEnabled = _selectedConnection IsNot Nothing
        MenuItemSaveIco.IsEnabled = _selectedConnection IsNot Nothing

        MenuItemUndo.IsEnabled = BoxSql.CanUndo
        MenuItemRedo.IsEnabled = BoxSql.CanRedo
        MenuItemCopyIco.IsEnabled = InlineAssignHelper(MenuItemCopy.IsEnabled, BoxSql.Selection.Text.Length > 0)
        MenuItemPasteIco.IsEnabled = InlineAssignHelper(MenuItemPaste.IsEnabled, Clipboard.ContainsText())
        MenuItemCutIco.IsEnabled = InlineAssignHelper(MenuItemCut.IsEnabled, BoxSql.Selection.Text.Length > 0)
        MenuItemSelectAll.IsEnabled = True

        MenuItemQueryAddDerived.IsEnabled = CanAddDerivedTable()

        MenuItemCopyUnionSq.IsEnabled = CanCopyUnionSubQuery()
        MenuItemAddUnionSq.IsEnabled = CanAddUnionSubQuery()
        MenuItemProp.IsEnabled = CanShowProperties()
        MenuItemAddObject.IsEnabled = CanAddObject()
        MenuItemProperties.IsEnabled = (_sqlFormattingOptions IsNot Nothing AndAlso QBuilder.SQLContext IsNot Nothing)

        For Each item As MenuItem In MetadataItemMenu.Items.Cast(Of FrameworkElement)().Where(Function(x) TypeOf x Is MenuItem).ToList()
            item.IsEnabled = QBuilder.SQLContext IsNot Nothing
        Next
    End Sub

    Public Function CanAddObject() As Boolean
        Return QBuilder.AddObjectDialog IsNot Nothing
    End Function

    Public Function CanShowProperties() As Boolean
        Return QBuilder.ActiveUnionSubQuery IsNot Nothing
    End Function

    Public Function CanAddUnionSubQuery() As Boolean
        If QBuilder.SyntaxProvider Is Nothing Then
            Return False
        End If

        If QBuilder.ActiveUnionSubQuery IsNot Nothing AndAlso QBuilder.ActiveUnionSubQuery.QueryRoot.IsSubQuery Then
            Return QBuilder.SyntaxProvider.IsSupportSubQueryUnions()
        End If

        Return QBuilder.SyntaxProvider.IsSupportUnions()
    End Function

    Public Function CanCopyUnionSubQuery() As Boolean
        Return CanAddUnionSubQuery()
    End Function

    Public Function CanAddDerivedTable() As Boolean
        If QBuilder.SyntaxProvider Is Nothing Then
            Return False
        End If

        If QBuilder.ActiveUnionSubQuery IsNot Nothing AndAlso QBuilder.ActiveUnionSubQuery.QueryRoot.IsMainQuery Then
            Return QBuilder.SyntaxProvider.IsSupportDerivedTables()
        End If

        Return QBuilder.SyntaxProvider.IsSupportSubQueryDerivedTables()
    End Function

    Private Sub InitializeSqlContext()
        Try
            QBuilder.Clear()

            Dim metadataProvider As BaseMetadataProvider = Nothing

            ' create new SqlConnection object using the connections string from the connection form
            If Not _selectedConnection.IsXmlFile Then
                Select Case _selectedConnection.ConnectionType
                    Case ConnectionTypes.MSSQL
                        metadataProvider = New MSSQLMetadataProvider() With {
                            .Connection = New SqlConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.MSAccess
                        metadataProvider = New OLEDBMetadataProvider() With {
                            .Connection = New OleDbConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.Oracle
                        ' previous version of this demo uses deprecated System.Data.OracleClient
                        ' current version uses Oracle.ManagedDataAccess.Client which doesn't support "Integrated Security" setting
                        Dim updatedConnectionString As String = Regex.Replace(_selectedConnection.ConnectionString, "Integrated Security=.*?;", "")

                        metadataProvider = New OracleNativeMetadataProvider() With {
                            .Connection = New OracleConnection(updatedConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.MySQL
                        metadataProvider = New MySQLMetadataProvider() With {
                            .Connection = New MySqlConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.PostgreSQL
                        metadataProvider = New PostgreSQLMetadataProvider() With {
                            .Connection = New NpgsqlConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.OLEDB
                        metadataProvider = New OLEDBMetadataProvider() With {
                            .Connection = New OleDbConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                        Case ConnectionTypes.ODBC
                        metadataProvider = New ODBCMetadataProvider() With {
                            .Connection = New OdbcConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case Else
                        Throw New ArgumentOutOfRangeException()
                End Select
         
            End If

            ' setup the query builder with metadata and syntax providers
            QBuilder.SQLContext.MetadataContainer.Clear()
            QBuilder.MetadataProvider = metadataProvider
            QBuilder.SyntaxProvider = _selectedConnection.SyntaxProvider
            QBuilder.MetadataLoadingOptions.OfflineMode = metadataProvider Is Nothing

            If metadataProvider Is Nothing Then
                QBuilder.MetadataContainer.ImportFromXML(_selectedConnection.ConnectionString)
            End If

            ' Instruct the query builder to fill the database schema tree

            QBuilder.InitializeDatabaseSchemaTree()
        Finally
            If QBuilder.MetadataContainer.LoadingOptions.OfflineMode Then
                TsmiOfflineMode.IsChecked = True
            End If

            If CBuilder.QueryTransformer IsNot Nothing Then
                RemoveHandler CBuilder.QueryTransformer.SQLUpdated, AddressOf QueryTransformer_SQLUpdated
            End If

            CBuilder.QueryTransformer = New QueryTransformer() With {
                .Query = QBuilder.SQLQuery,
                .SQLGenerationOptions = _sqlGenerationOptions
            }

            AddHandler CBuilder.QueryTransformer.SQLUpdated, AddressOf QueryTransformer_SQLUpdated

            DataGridResult.QueryTransformer = CBuilder.QueryTransformer
            DataGridResult.SqlQuery = QBuilder.SQLQuery

            ' The pagination panel is displayed if the current SyntaxProvider has support for pagination
            PaginationPanel.Visibility = If((CBuilder.QueryTransformer.IsSupportLimitCount OrElse CBuilder.QueryTransformer.IsSupportLimitOffset) AndAlso QBuilder.SyntaxProvider IsNot Nothing, Visibility.Visible, Visibility.Collapsed)


            PaginationPanel.IsSupportLimitCount = CBuilder.QueryTransformer.IsSupportLimitCount
            PaginationPanel.IsSupportLimitOffset = CBuilder.QueryTransformer.IsSupportLimitOffset
        End Try
    End Sub

    Private Sub QueryTransformer_SQLUpdated(sender As Object, e As EventArgs)
        ' Handle the event raised by Query Transformer object that the text of SQL query is changed
        ' update the text box
        If DataGridResult.QueryTransformer Is Nothing OrElse Not Equals(TabItemData, TabControl.SelectedItem) Then
            Return
        End If

        DataGridResult.FillDataGrid(DataGridResult.QueryTransformer.SQL)
        SetTextRichTextBox(DataGridResult.QueryTransformer.SQL, BoxSqlTransformer)
    End Sub

    Private Sub MenuItemQueryStatistics_OnClick(sender As Object, e As RoutedEventArgs)
        ShowQueryStatistics()
    End Sub

    Public Sub ShowQueryStatistics()
        Dim stats As String 

        Dim qs As QueryStatistics = QBuilder.QueryStatistics

        stats = "Used Objects (" & qs.UsedDatabaseObjects.Count & "):" & vbCr & vbLf
        stats = qs.UsedDatabaseObjects.Aggregate(stats, Function(current, t) current & (vbCr & vbLf & t.ObjectName.QualifiedName))

        stats += vbCr & vbLf & vbCr & vbLf & "Used Columns (" & qs.UsedDatabaseObjectFields.Count & "):" & vbCr & vbLf
        stats = qs.UsedDatabaseObjectFields.Aggregate(stats, Function(current, t) current & (vbCr & vbLf & t.ObjectName.QualifiedName))

        stats += vbCr & vbLf & vbCr & vbLf & "Output Expressions (" & qs.OutputColumns.Count & "):" & vbCr & vbLf
        stats = qs.OutputColumns.Aggregate(stats, Function(current, t) current & (vbCr & vbLf & t.Expression))

        Dim f As QueryStatisticsWindow = New QueryStatisticsWindow()
        f.textBox.Text = stats

        f.ShowDialog()
    End Sub

    Private Sub CommandOpen_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        ' open a saved query
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog() With {
            .DefaultExt = "sql",
            .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
        }

        If openFileDialog1.ShowDialog() <> True Then
            Return
        End If

        Dim sb As StringBuilder = New StringBuilder()

        Using sr As StreamReader = New StreamReader(openFileDialog1.FileName)
            Dim s As String

            While (InlineAssignHelper(s, sr.ReadLine())) IsNot Nothing
                sb.AppendLine(s)
            End While
        End Using

        If QBuilder.SQLContext Is Nothing Then
            CommandNew_OnExecuted(Nothing, Nothing)
        Else
            Try
                ' load query to the query builder by assigning its text to the SQL property
                QBuilder.SQL = sb.ToString()

            Finally
            End Try
        End If
    End Sub

    Private Sub CommandSave_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        SaveInFile()
    End Sub

    Private Sub CommandExit_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Close()
    End Sub

    Private Sub CommandNew_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim cf As DatabaseConnectionWindow = New DatabaseConnectionWindow(_showHintConnection) With {
            .Owner = Me
        }

        _showHintConnection = False

        If cf.ShowDialog() <> True Then
            Return
        End If
        _selectedConnection = cf.SelectedConnection

        InitializeSqlContext()
    End Sub

    Private Sub CommandUndo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        BoxSql.Undo()
    End Sub

    Private Sub CommandRedo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        BoxSql.Redo()
    End Sub

    Private Sub CommandCopy_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        BoxSql.Copy()
    End Sub

    Private Sub CommandPaste_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        BoxSql.Paste()
    End Sub

    Private Sub CommandCut_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        BoxSql.Cut()
    End Sub

    Private Sub CommandSelectAll_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        BoxSql.SelectAll()
    End Sub

    Private Sub MenuItemQueryAddDerived_OnClick(sender As Object, e As RoutedEventArgs)
        AddDerivedTable()
    End Sub

    Private Sub AddDerivedTable()
        Using New UpdateRegion(QBuilder.ActiveUnionSubQuery.FromClause)
            Dim sqlContext As SQLContext = QBuilder.ActiveUnionSubQuery.SQLContext

            Dim fq As SQLFromQuery = New SQLFromQuery(sqlContext) With {
                .[Alias] = New SQLAliasObjectAlias(sqlContext) With {
                    .[Alias] = QBuilder.ActiveUnionSubQuery.QueryRoot.CreateUniqueSubQueryName()
                },
                .SubQuery = New SQLSubSelectStatement(sqlContext)
            }

            Dim sqse As SQLSubQuerySelectExpression = New SQLSubQuerySelectExpression(sqlContext)
            fq.SubQuery.Add(sqse)
            sqse.SelectItems = New SQLSelectItems(sqlContext)
            sqse.From = New SQLFromClause(sqlContext)

            QBuilder.QueryNavigationBar.Query.AddObject(QBuilder.ActiveUnionSubQuery, fq, GetType(DataSourceQuery))
        End Using
    End Sub

    Private Sub MenuItemCopyUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        CopyUnionSubQuery()
    End Sub

    Private Sub CopyUnionSubQuery()
        ' add an empty UnionSubQuery
        Dim usq As UnionSubQuery = QBuilder.ActiveUnionSubQuery.ParentGroup.Add()

        ' copy the content of existing union sub-query to a new one
        Dim usqAst As SQLSubQuerySelectExpression = QBuilder.ActiveUnionSubQuery.ResultQueryAST
        usqAst.RestoreColumnPrefixRecursive(True)

        Dim lCte As List(Of SQLWithClauseItem) = New List(Of SQLWithClauseItem)()
        Dim lFromObj As List(Of SQLFromSource) = New List(Of SQLFromSource)()
        QBuilder.ActiveUnionSubQuery.GatherPrepareAndFixupContext(lCte, lFromObj, False)
        usqAst.PrepareAndFixupRecursive(lCte, lFromObj)

        usq.LoadFromAST(usqAst)
        QBuilder.ActiveUnionSubQuery = usq
    End Sub

    Private Sub MenuItemAddUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        AddUnionSubQuery()
    End Sub

    Private Sub AddUnionSubQuery()
        QBuilder.ActiveUnionSubQuery = QBuilder.ActiveUnionSubQuery.ParentGroup.Add()
    End Sub

    Private Sub MenuItemProp_OnClick(sender As Object, e As RoutedEventArgs)
        PropertiesQuery()
    End Sub

    Private Sub PropertiesQuery()
        QBuilder.ShowActiveUnionSubQueryProperties()
    End Sub

    Private Sub MenuItemAddObject_OnClick(sender As Object, e As RoutedEventArgs)
        AddObject()
    End Sub

    Private Sub AddObject()
        If QBuilder.AddObjectDialog IsNot Nothing Then
            QBuilder.AddObjectDialog.ShowModal()
        End If
    End Sub

    Private Sub MenuItemProperties_OnClick(sender As Object, e As RoutedEventArgs)
        Dim propWindow As QueryPropertiesWindow = New QueryPropertiesWindow(QBuilder.SQLContext, _sqlFormattingOptions)
        propWindow.ShowDialog()
    End Sub

    Private Sub LanguageMenuItemChecked(sender As Object, e As RoutedEventArgs)
        e.Handled = True

        Dim menuItem As MenuItem = DirectCast(sender, MenuItem)
        Dim lng As String = menuItem.Tag.ToString()
        Language = XmlLanguage.GetLanguage(lng)
    End Sub

    Private Sub MenuItem_OfflineMode_OnChecked(sender As Object, e As RoutedEventArgs)
        Dim menuItem As MenuItem = DirectCast(sender, MenuItem)

        If menuItem.IsChecked Then
            Try
                Cursor = Cursors.Wait

                QBuilder.MetadataContainer.LoadAll(True)
            Finally
                Cursor = Cursors.Arrow
            End Try
        End If

        QBuilder.MetadataContainer.LoadingOptions.OfflineMode = menuItem.IsChecked
    End Sub

    Private Sub MenuItem_RefreshMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        If QBuilder.SQLContext.MetadataProvider IsNot Nothing AndAlso QBuilder.SQLContext.MetadataProvider.Connected Then
            ' Force the query builder to refresh metadata from current connection
            ' to refresh metadata, just clear MetadataContainer and reinitialize metadata tree
            QBuilder.MetadataContainer.Clear()
            QBuilder.InitializeDatabaseSchemaTree()
        End If
    End Sub

    Private Sub MenuItem_ClearMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        QBuilder.MetadataContainer.Clear()
    End Sub

    Private Sub MenuItem_LoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog As OpenFileDialog = New OpenFileDialog() With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        }

        If fileDialog.ShowDialog() <> True Then
            Return
        End If

        QBuilder.MetadataContainer.LoadingOptions.OfflineMode = True
        QBuilder.MetadataContainer.ImportFromXML(fileDialog.FileName)

        ' Instruct the query builder to fill the database schema tree
        QBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_SaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog = New SaveFileDialog() With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
            .FileName = "Metadata.xml"
        }

        If fileDialog.ShowDialog() <> True Then
            Return
        End If

        QBuilder.MetadataContainer.LoadAll(True)
        QBuilder.MetadataContainer.ExportToXML(fileDialog.FileName)
    End Sub

    Private Sub MenuItem_About_OnClick(sender As Object, e As RoutedEventArgs)
        Dim f = New AboutForm() With {
            .Owner = Me
        }

        f.ShowDialog()
    End Sub

    Private Sub SaveInFile()
        ' Save the query text to file
        If QBuilder.ActiveUnionSubQuery Is Nothing Then
            Return
        End If

        Dim saveFileDialog1 As SaveFileDialog = New SaveFileDialog() With {
            .DefaultExt = "sql",
            .FileName = "query",
            .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
        }

        If saveFileDialog1.ShowDialog() <> True Then
            Return
        End If

        Using sw As StreamWriter = New StreamWriter(saveFileDialog1.FileName)
            sw.Write(FormattedSQLBuilder.GetSQL(QBuilder.ActiveUnionSubQuery.QueryRoot, _sqlFormattingOptions))
        End Using
    End Sub

    Private Sub BoxSql_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QBuilder.SQL = GetTextRichTextBox(BoxSql)

            ' Hide error banner if any
            ErrorBox.Message = String.Empty
        Catch ex As SQLParsingException
            ' Show banner with error text
            ErrorBox.Message = ex.Message
        End Try
    End Sub

    Private Sub QBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
        SetTextRichTextBox(QBuilder.FormattedSQL, BoxSql)

        SetSqlTextCurrentSubQuery()

        If Not TabItemFastResult.IsSelected OrElse CheckBoxAutoRefresh.IsChecked = False Then
            Return
        End If

        _timerStartingExecuteSql.Change(600, Timeout.Infinite)
    End Sub

    Private Sub TimerStartingExecuteSql_Elapsed(state As Object)
        Dispatcher.BeginInvoke(DirectCast(AddressOf FillFastResult, Action))
    End Sub

    Private Shared Sub SetTextRichTextBox(text As String, editor As RichTextBox)
        editor.Document.Blocks.Clear()
        editor.Document.Blocks.Add(New Paragraph(New Run(text)))
    End Sub

    Private Shared Function GetTextRichTextBox(editor As RichTextBox) As String
        Return New TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text
    End Function

    Private Sub TabControl_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ' Execute a query on switching to the Data tab
        If e.AddedItems.Count = 0 OrElse _selectedConnection Is Nothing Then
            Return
        End If

        If QBuilder.SyntaxProvider Is Nothing OrElse Not Equals(e.AddedItems(0), TabItemData) Then
            Return
        End If

        ResetPagination()

        SetTextRichTextBox(GetTextRichTextBox(BoxSql), BoxSqlTransformer)

        If Equals(TabItemData, TabControl.SelectedItem) Then
            BorderBlockPagination.Visibility = Visibility.Visible
            DataGridResult.FillDataGrid(GetTextRichTextBox(BoxSql))
        End If
    End Sub

    Private Sub ResetPagination()
        PaginationPanel.Reset()
        CBuilder.QueryTransformer.Skip("")
        CBuilder.QueryTransformer.Take("")
    End Sub

    Private Sub PaginationPanel_OnCurrentPageChanged(sender As Object, e As RoutedEventArgs)
        If PaginationPanel.CurrentPage = 1 Then
            CBuilder.QueryTransformer.Skip("")
            Return
        End If

        ' Select next n records
        CBuilder.QueryTransformer.Skip((PaginationPanel.PageSize * (PaginationPanel.CurrentPage - 1)).ToString())
    End Sub

    Private Sub PaginationPanel_OnEnabledPaginationChanged(sender As Object, e As RoutedEventArgs)
        ' Turn paging on and off
        If PaginationPanel.IsEnabled Then
            CBuilder.QueryTransformer.Take(PaginationPanel.PageSize.ToString())
        Else
            ResetPagination()
        End If
    End Sub

    Private Sub PaginationPanel_OnPageSizeChanged(sender As Object, e As RoutedEventArgs)
        CBuilder.QueryTransformer.Take(PaginationPanel.PageSize.ToString())
    End Sub

    Private Sub BoxSql_OnTextChanged(sender As Object, e As TextChangedEventArgs)
        ErrorBox.Message = String.Empty
    End Sub

    Private Sub QBuilder_OnActiveUnionSubQueryChanged(sender As Object, e As EventArgs)
        SetSqlTextCurrentSubQuery()
    End Sub

    Private Sub SetSqlTextCurrentSubQuery()
        BorderErrorFast.Visibility = Visibility.Collapsed
        If _transformerSql Is Nothing Then
            Return
        End If

        If QBuilder.ActiveUnionSubQuery Is Nothing Or QBuilder.SleepMode Then
            SetTextRichTextBox("", BoxSqlCurrentSubQuery)
            _transformerSql.Query = Nothing
            Return
        End If


        _transformerSql.Query = New SQLQuery(QBuilder.ActiveUnionSubQuery.SQLContext) With {
            .SQL = QBuilder.ActiveUnionSubQuery.ParentSubQuery.GetSqlForDataPreview()
        }

        Dim sql As String = QBuilder.ActiveUnionSubQuery.ParentSubQuery.GetResultSQL(_sqlFormattingOptions)
        SetTextRichTextBox(sql, BoxSqlCurrentSubQuery)
    End Sub

    Private Sub FillFastResult()
        Dim result As QueryTransformer = _transformerSql.Take("10")

        Try
            Dim dv As DataView = Common.Helpers.ExecuteSql(result.SQL, DirectCast(_transformerSql.Query, SQLQuery))

            ListViewFastResultSql.ItemsSource = dv
            BorderError.Visibility = Visibility.Collapsed
        Catch exception As Exception
            BorderError.Visibility = Visibility.Visible
            LabelError.Text = exception.Message

            ListViewFastResultSql.ItemsSource = Nothing
        End Try
    End Sub

    Private Sub TabControlSql_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If ErrorBox.Visibility = Visibility.Visible Then
            ErrorBox.Visibility = Visibility.Collapsed
        End If
        If Not e.AddedItems.Contains(TabItemFastResult) OrElse _transformerSql.Query Is Nothing OrElse CheckBoxAutoRefresh.IsChecked = False Then
            Return
        End If

        FillFastResult()
    End Sub

    Private Sub ButtonRefreshFastResult_OnClick(sender As Object, e As RoutedEventArgs)
        FillFastResult()
    End Sub

    Private Sub CheckBoxAutoRefresh_OnChecked(sender As Object, e As RoutedEventArgs)
        If ButtonRefreshFastResult Is Nothing OrElse CheckBoxAutoRefresh Is Nothing Then
            Return
        End If
        ButtonRefreshFastResult.IsEnabled = CheckBoxAutoRefresh.IsChecked = False
    End Sub

    Private Sub CloseImage_OnMouseUp(sender As Object, e As MouseButtonEventArgs)
        BorderError.Visibility = Visibility.Collapsed
    End Sub

    Private Sub DataGridResult_OnRowsLoaded(sender As Object, e As EventArgs)
        If Not PaginationPanel.IsEnabled Then
            PaginationPanel.CountRows = DataGridResult.CountRows
        End If
        BorderBlockPagination.Visibility = Visibility.Collapsed
    End Sub

    Private Sub BoxSqlCurrentSubQuery_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        If QBuilder.ActiveUnionSubQuery Is Nothing Then
            Return
        End If
        Try
            BorderErrorFast.Visibility = Visibility.Collapsed

            QBuilder.ActiveUnionSubQuery.ParentSubQuery.SQL = GetTextRichTextBox(DirectCast(sender, RichTextBox))

            Dim sql As String = QBuilder.ActiveUnionSubQuery.ParentSubQuery.GetResultSQL(_sqlFormattingOptions)

            _transformerSql.Query = New SQLQuery(QBuilder.ActiveUnionSubQuery.SQLContext) With {
                .SQL = sql
            }
        Catch ex As Exception
            LabelErrorFast.Text = ex.Message
            BorderErrorFast.Visibility = Visibility.Visible
        End Try
    End Sub
    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function

    Private Sub MenuItemEditMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        QueryBuilder.EditMetadataContainer(QBuilder.SQLContext)
    End Sub
End Class
