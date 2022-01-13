''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports ActiveQueryBuilder.Core.QueryTransformer
Imports Microsoft.Win32
Imports System.Drawing.Imaging
Imports System.Globalization
Imports System.IO
Imports System.Threading
Imports System.Windows.Markup
Imports Helpers = ActiveQueryBuilder.Core.Helpers
Imports BuildInfo = ActiveQueryBuilder.Core.BuildInfo
Imports Brushes = System.Windows.Media.Brushes
Imports GeneralAssembly
Imports GeneralAssembly.Common
Imports GeneralAssembly.Windows.QueryInformationWindows
Imports GeneralAssembly.QueryBuilderProperties
Imports GeneralAssembly.Windows

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private _selectedConnection As ConnectionInfo
    'private readonly SQLFormattingOptions _sqlFormattingOptions;
    'private readonly SQLGenerationOptions _sqlGenerationOptions;
    Private _showHintConnection As Boolean = True
    Private ReadOnly _transformerSql As QueryTransformer
    Private ReadOnly _timerStartingExecuteSql As Timer

    Public Sub New()
        InitializeComponent()
        ' Options to present the formatted SQL query text to end-user
        ' Use names of virtual objects, do not replace them with appropriate derived tables
        QBuilder.SQLFormattingOptions = New SQLFormattingOptions With {.ExpandVirtualObjects = False}

        ' Options to generate the SQL query text for execution against a database server
        ' Replace virtual objects with derived tables
        QBuilder.SQLGenerationOptions = New SQLGenerationOptions With {.ExpandVirtualObjects = True}



        AddHandler Closing, AddressOf MainWindow_Closing
        AddHandler Dispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive

        Dim currentLang = Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName

        LoadLanguage()

        Dim defLang = "en"

        If Helpers.Localizer.Languages.Contains(currentLang.ToLower()) Then
            Language = XmlLanguage.GetLanguage(currentLang)
            defLang = currentLang.ToLower()
        End If

        Dim menuItem = MenuItemLanguage.Items.Cast(Of MenuItem)().First(Function(item) CStr(item.Tag) = defLang)
        menuItem.IsChecked = True

        QBuilder.SyntaxProvider = New GenericSyntaxProvider()

        _transformerSql = New QueryTransformer()

        _timerStartingExecuteSql = New Timer(AddressOf TimerStartingExecuteSql_Elapsed)

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

            Dim button = New Button With {
                .Background = Brushes.Transparent,
                .Padding = New Thickness(0),
                .BorderThickness = New Thickness(0),
                .Cursor = Cursors.Hand,
                .Margin = New Thickness(0, 0, 5, 0),
                .HorizontalAlignment = HorizontalAlignment.Right,
                .VerticalAlignment = VerticalAlignment.Center,
                .Content = New System.Windows.Controls.Image With {
                    .Source = ActiveQueryBuilder.View.WPF.Helpers.GetImageSource(My.Resources.cancel, ImageFormat.Png),
                    .Stretch = Stretch.None
                }
            }

            AddHandler button.Click, Sub()
                                         GridRoot.Visibility = Visibility.Collapsed
                                     End Sub

            trialNoticePanel.Child = label
            GridRoot.Children.Add(trialNoticePanel)
            GridRoot.Children.Add(button)
        End If

        QBuilder.SQLQuery.QueryRoot.AllowSleepMode = True

        AddHandler QBuilder.SleepModeChanged, AddressOf SqlQuery_SleepModeChanged
        AddHandler QBuilder.QueryAwake, AddressOf SqlQuery_QueryAwake
    End Sub

    Private _shown As Boolean
    Private _errorPosition As Integer = -1
    Private _lastValidText As String
    Private _lastValidText1 As String
    Private _errorPosition1 As Integer = -1

    Protected Overrides Sub OnContentRendered(e As EventArgs)
        MyBase.OnContentRendered(e)

        If _shown Then
            Return
        End If

        _shown = True

        CommandNew_OnExecuted(Me, Nothing)
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
        'INSTANT VB NOTE: The variable language was renamed since Visual Basic does not handle local variables named the same as class members well:
        For Each language_Renamed In Helpers.Localizer.Languages
            If language_Renamed.ToLower() = "auto" OrElse language_Renamed.ToLower() = "default" Then
                Continue For
            End If

            Dim culture = New CultureInfo(language_Renamed)

            Dim stroke = String.Format("{0}", culture.DisplayName)

            Dim menuItem = New MenuItem With {
                .Header = stroke,
                .Tag = language_Renamed,
                .IsCheckable = True
            }

            MenuItemLanguage.Items.Add(menuItem)
            menuItem.SetValue(GroupedMenuBehavior.OptionGroupNameProperty, "group")
        Next language_Renamed
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs)
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
        MenuItemCopy.IsEnabled = BoxSql.SelectionLength > 0
        MenuItemCopyIco.IsEnabled = MenuItemCopy.IsEnabled
        MenuItemPaste.IsEnabled = Clipboard.ContainsText()
        MenuItemPasteIco.IsEnabled = MenuItemPaste.IsEnabled
        MenuItemCut.IsEnabled = BoxSql.SelectionLength > 0
        MenuItemCutIco.IsEnabled = MenuItemCut.IsEnabled
        MenuItemSelectAll.IsEnabled = True

        MenuItemQueryAddDerived.IsEnabled = CanAddDerivedTable()

        MenuItemCopyUnionSq.IsEnabled = CanCopyUnionSubQuery()
        MenuItemAddUnionSq.IsEnabled = CanAddUnionSubQuery()
        MenuItemProp.IsEnabled = CanShowProperties()
        MenuItemAddObject.IsEnabled = CanAddObject()
        MenuItemProperties.IsEnabled = (QBuilder.SQLFormattingOptions IsNot Nothing AndAlso QBuilder.SQLContext IsNot Nothing)
        MenuItemUserExpression.IsEnabled = _selectedConnection IsNot Nothing
        For Each item In MetadataItemMenu.Items.Cast(Of FrameworkElement)().Where(Function(x) TypeOf x Is MenuItem).ToList()
            item.IsEnabled = QBuilder.SQLContext IsNot Nothing
        Next item
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

            Dim metadataProvaider As BaseMetadataProvider = Nothing

            ' create new SqlConnection object using the connections string from the connection form
            If Not _selectedConnection.IsXmlFile Then
                metadataProvaider = _selectedConnection.ConnectionDescriptor?.MetadataProvider
            End If

            ' setup the query builder with metadata and syntax providers
            QBuilder.SQLContext.MetadataContainer.Clear()
            QBuilder.MetadataProvider = metadataProvaider
            QBuilder.SyntaxProvider = _selectedConnection.ConnectionDescriptor.SyntaxProvider
            QBuilder.MetadataLoadingOptions.OfflineMode = metadataProvaider Is Nothing

            If metadataProvaider Is Nothing Then
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

            CBuilder.QueryTransformer = New QueryTransformer With {
                .Query = QBuilder.SQLQuery,
                .SQLGenerationOptions = QBuilder.SQLGenerationOptions
            }

            AddHandler CBuilder.QueryTransformer.SQLUpdated, AddressOf QueryTransformer_SQLUpdated

            DataGridResult.QueryTransformer = CBuilder.QueryTransformer
            DataGridResult.SqlQuery = QBuilder.SQLQuery
        End Try
    End Sub

    Private Sub QueryTransformer_SQLUpdated(sender As Object, e As EventArgs)
        ' Handle the event raised by Query Transformer object that the text of SQL query is changed
        ' update the text box
        If DataGridResult.QueryTransformer Is Nothing OrElse Not Equals(TabItemData, TabControl.SelectedItem) Then
            Return
        End If

        DataGridResult.FillData(DataGridResult.QueryTransformer.SQL)
        BoxSqlTransformer.Text = DataGridResult.QueryTransformer.SQL
    End Sub

    Private Sub MenuItemQueryStatistics_OnClick(sender As Object, e As RoutedEventArgs)
        ShowQueryStatistics()
    End Sub

    Public Sub ShowQueryStatistics()
        Dim qs = QBuilder.QueryStatistics

        Dim stats = "Used Objects (" & qs.UsedDatabaseObjects.Count & "):" & ControlChars.CrLf
        stats = qs.UsedDatabaseObjects.Aggregate(stats, Function(current, t) current + (ControlChars.CrLf & t.ObjectName.QualifiedName))

        stats &= ControlChars.CrLf & ControlChars.CrLf & "Used Columns (" & qs.UsedDatabaseObjectFields.Count & "):" & ControlChars.CrLf
        stats = qs.UsedDatabaseObjectFields.Aggregate(stats, Function(current, t) current + (ControlChars.CrLf & t.ObjectName.QualifiedName))

        stats &= ControlChars.CrLf & ControlChars.CrLf & "Output Expressions (" & qs.OutputColumns.Count & "):" & ControlChars.CrLf
        stats = qs.OutputColumns.Aggregate(stats, Function(current, t) current + (ControlChars.CrLf & t.Expression))

        Dim f = New QueryStatisticsWindow(stats)

        f.ShowDialog()
    End Sub

    Private Sub CommandOpen_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        ' open a saved query
        Dim openFileDialog1 = New OpenFileDialog With {
            .DefaultExt = "sql",
            .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
        }

        If Not openFileDialog1.ShowDialog().Equals(True) Then
            Return
        End If

        Dim sb = New StringBuilder()

        Using sr = New StreamReader(openFileDialog1.FileName)
            Dim s As String

            'INSTANT VB WARNING: An assignment within expression was extracted from the following statement:
            'ORIGINAL LINE: while ((s = sr.ReadLine()) != null)
            s = sr.ReadLine()
            Do While s IsNot Nothing
                sb.AppendLine(s)
                s = sr.ReadLine()
            Loop
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
        Dim cf = New DatabaseConnectionWindow(_showHintConnection) With {.Owner = Me}

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
        Using tempVar As New UpdateRegion(QBuilder.ActiveUnionSubQuery.FromClause)
            Dim sqlContext = QBuilder.ActiveUnionSubQuery.SQLContext

            Dim fq = New SQLFromQuery(sqlContext) With {
                .Alias = New SQLAliasObjectAlias(sqlContext) With {.Alias = QBuilder.ActiveUnionSubQuery.QueryRoot.CreateUniqueSubQueryName()},
                .SubQuery = New SQLSubSelectStatement(sqlContext)
            }

            Dim sqse = New SQLSubQuerySelectExpression(sqlContext)
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
        Dim usq = QBuilder.ActiveUnionSubQuery.ParentGroup.Add()

        ' copy the content of existing union sub-query to a new one
        Dim usqAst = QBuilder.ActiveUnionSubQuery.ResultQueryAST
        usqAst.RestoreColumnPrefixRecursive(True)

        Dim lCte = New List(Of SQLWithClauseItem)()
        Dim lFromObj = New List(Of SQLFromSource)()
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
        Dim propWindow = New QueryBuilderPropertiesWindow(QBuilder)
        propWindow.ShowDialog()
    End Sub

    Private Sub LanguageMenuItemChecked(sender As Object, e As RoutedEventArgs)
        e.Handled = True

        Dim menuItem = DirectCast(sender, MenuItem)
        Dim lng = menuItem.Tag.ToString()
        Language = XmlLanguage.GetLanguage(lng)
    End Sub

    Private Sub MenuItem_OfflineMode_OnChecked(sender As Object, e As RoutedEventArgs)
        Dim menuItem = DirectCast(sender, MenuItem)

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
        If QBuilder.SQLContext Is Nothing OrElse QBuilder.SQLContext.MetadataProvider Is Nothing OrElse Not QBuilder.SQLContext.MetadataProvider.Connected Then
            Return
        End If
        ' Force the query builder to refresh metadata from current connection
        ' to refresh metadata, just clear MetadataContainer and reinitialize metadata tree
        QBuilder.MetadataContainer.Clear()
        QBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_ClearMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        QBuilder.MetadataContainer.Clear()
    End Sub

    Private Sub MenuItem_LoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog = New OpenFileDialog With {.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"}

        If Not fileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        QBuilder.MetadataContainer.LoadingOptions.OfflineMode = True
        QBuilder.MetadataContainer.ImportFromXML(fileDialog.FileName)

        ' Instruct the query builder to fill the database schema tree
        QBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_SaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog = New SaveFileDialog With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
            .FileName = "Metadata.xml"
        }

        If Not fileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        QBuilder.MetadataContainer.LoadAll(True)
        QBuilder.MetadataContainer.ExportToXML(fileDialog.FileName)
    End Sub

    Private Sub MenuItem_About_OnClick(sender As Object, e As RoutedEventArgs)
        Dim f = New AboutWindow With {.Owner = Me}

        f.ShowDialog()
    End Sub

    Private Sub SaveInFile()
        ' Save the query text to file
        If QBuilder.ActiveUnionSubQuery Is Nothing Then
            Return
        End If

        Dim saveFileDialog1 = New SaveFileDialog() With {
            .DefaultExt = "sql",
            .FileName = "query",
            .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
        }

        If saveFileDialog1.ShowDialog() <> True Then
            Return
        End If

        Using sw = New StreamWriter(saveFileDialog1.FileName)
            sw.Write(FormattedSQLBuilder.GetSQL(QBuilder.ActiveUnionSubQuery.QueryRoot, QBuilder.SQLFormattingOptions))
        End Using
    End Sub

    Private Sub BoxSql_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QBuilder.SQL = BoxSql.Text

            ' Hide error banner if any
            ErrorBox.Show(Nothing, QBuilder.SyntaxProvider)
        Catch ex As SQLParsingException
            ' Show banner with error text
            ErrorBox.Show(ex.Message, QBuilder.SyntaxProvider)
            _errorPosition = ex.ErrorPos.pos
        End Try
    End Sub

    Private Sub QBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
        BoxSql.Text = QBuilder.FormattedSQL
        _lastValidText = BoxSql.Text
        SetSqlTextCurrentSubQuery()

        If Not TabItemFastResult.IsSelected OrElse CheckBoxAutoRefreash.IsChecked = False Then
            Return
        End If

        _timerStartingExecuteSql.Change(600, Timeout.Infinite)
    End Sub

    Private Sub TimerStartingExecuteSql_Elapsed(state As Object)
        Dispatcher.BeginInvoke(CType(AddressOf FillFastResult, Action))
    End Sub

    Private Sub TabControl_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ' Execute a query on switching to the Data tab
        If e.AddedItems.Count = 0 OrElse _selectedConnection Is Nothing Then
            Return
        End If

        If QBuilder.SyntaxProvider Is Nothing OrElse Not Equals(e.AddedItems(0), TabItemData) Then
            Return
        End If

        CBuilder.Clear()
        BoxSqlTransformer.Text = BoxSql.Text

        If Equals(TabItemData, TabControl.SelectedItem) Then
            DataGridResult.FillData(BoxSql.Text)
        End If
    End Sub

    Private Sub BoxSql_OnTextChanged(sender As Object, e As TextChangedEventArgs)
        ErrorBox.Show(Nothing, QBuilder.SyntaxProvider)
    End Sub

    Private Sub QBuilder_OnActiveUnionSubQueryChanged(sender As Object, e As EventArgs)
        SetSqlTextCurrentSubQuery()
    End Sub

    Private Sub SetSqlTextCurrentSubQuery()
        ErrorBox2.Show(Nothing, QBuilder.SyntaxProvider)

        If _transformerSql Is Nothing Then
            Return
        End If

        Dim activeUnionSubQuery = QBuilder.ActiveUnionSubQuery
        If activeUnionSubQuery Is Nothing OrElse QBuilder.SleepMode Then
            BoxSqlCurrentSubQuery.Text = ""
            _transformerSql.Query = Nothing
            Return
        End If

        Dim parentSubQuery = activeUnionSubQuery.ParentSubQuery

        Dim sqlForDataPreview = parentSubQuery.GetSqlForDataPreview()
        _transformerSql.Query = New SQLQuery(activeUnionSubQuery.SQLContext) With {.SQL = sqlForDataPreview}

        Dim sql = parentSubQuery.GetResultSQL(QBuilder.SQLFormattingOptions)
        BoxSqlCurrentSubQuery.Text = sql
    End Sub

    Private Sub FillFastResult()
        Dim result = _transformerSql.Take("10")
        ListViewFastResultSql.FillData(result.SQL, DirectCast(_transformerSql.Query, SQLQuery))
    End Sub

    Private Sub TabControlSql_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If ErrorBox.Visibility = Visibility.Visible Then
            ErrorBox.Visibility = Visibility.Collapsed
        End If
        If Not e.AddedItems.Contains(TabItemFastResult) OrElse _transformerSql.Query Is Nothing OrElse CheckBoxAutoRefreash.IsChecked = False Then
            Return
        End If

        FillFastResult()
    End Sub

    Private Sub ButtonRefreshFastResult_OnClick(sender As Object, e As RoutedEventArgs)
        FillFastResult()
    End Sub

    Private Sub CheckBoxAutoRefresh_OnChecked(sender As Object, e As RoutedEventArgs)
        If ButtonRefreashFastResult Is Nothing OrElse CheckBoxAutoRefreash Is Nothing Then
            Return
        End If
        Dim isChecked As Boolean? = CheckBoxAutoRefreash.IsChecked
        ButtonRefreashFastResult.IsEnabled = isChecked.HasValue AndAlso isChecked.Value = False
    End Sub

    Private Sub BoxSqlCurrentSubQuery_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        If QBuilder.ActiveUnionSubQuery Is Nothing Then
            Return
        End If
        Try
            ErrorBox.Show(Nothing, QBuilder.SyntaxProvider)

            QBuilder.ActiveUnionSubQuery.ParentSubQuery.SQL = DirectCast(sender, TextBox).Text

            Dim sql = QBuilder.ActiveUnionSubQuery.ParentSubQuery.GetResultSQL(QBuilder.SQLFormattingOptions)

            _transformerSql.Query = New SQLQuery(QBuilder.ActiveUnionSubQuery.SQLContext) With {.SQL = sql}
            _lastValidText1 = DirectCast(sender, TextBox).Text
        Catch ex As SQLParsingException
            ErrorBox2.Show(ex.Message, QBuilder.SyntaxProvider)
            _errorPosition1 = ex.ErrorPos.pos
        End Try
    End Sub

    Private Sub MenuItemEditMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        QueryBuilder.EditMetadataContainer(QBuilder.SQLContext)
    End Sub

    Private Sub ErrorBox_OnGoToErrorPosition(sender As Object, e As EventArgs)
        BoxSql.Focus()

        If _errorPosition = -1 Then
            Return
        End If

        If BoxSql.LineCount <> 1 Then
            BoxSql.ScrollToLine(BoxSql.GetLineIndexFromCharacterIndex(_errorPosition))
        End If
        BoxSql.CaretIndex = _errorPosition
    End Sub

    Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
        BoxSql.Text = _lastValidText
        BoxSql.Focus()
    End Sub

    Private Sub ErrorBox_OnGoToErrorPositionEvent(sender As Object, e As EventArgs)
        BoxSqlCurrentSubQuery.Focus()

        If _errorPosition1 = -1 Then
            Return
        End If

        If BoxSqlCurrentSubQuery.LineCount <> 1 Then
            BoxSqlCurrentSubQuery.ScrollToLine(BoxSqlCurrentSubQuery.GetLineIndexFromCharacterIndex(_errorPosition1))
        End If
        BoxSqlCurrentSubQuery.CaretIndex = _errorPosition1
    End Sub

    Private Sub ErrorBox_OnRevertValidTextEvent(sender As Object, e As EventArgs)
        BoxSqlCurrentSubQuery.Text = _lastValidText1
        BoxSqlCurrentSubQuery.Focus()
    End Sub

    Private Sub MenuItemUserExpression_OnClick(sender As Object, e As RoutedEventArgs)
        Dim window = New EditUserExpressionWindow With {
            .Owner = Me,
            .WindowStartupLocation = WindowStartupLocation.CenterOwner
        }
        window.Load(QBuilder.QueryView)
        window.ShowDialog()
    End Sub
End Class
