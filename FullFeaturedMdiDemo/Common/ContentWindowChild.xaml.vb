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
Imports System.ComponentModel
Imports System.Data
Imports System.Linq
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.QueryTransformer
Imports ActiveQueryBuilder.View.WPF.ExpressionEditor
Imports ActiveQueryBuilder.View.WPF.QueryView

Namespace Common
    ''' <summary>
    ''' Interaction logic for ContentWindowChild.xaml
    ''' </summary>
    Partial Public Class ContentWindowChild
        Public Event SaveQueryEvent As EventHandler
        Public Event SaveAsInFileEvent As EventHandler
        Public Event SaveAsNewUserQueryEvent As EventHandler

        Private ReadOnly _transformerSql As QueryTransformer
        Private ReadOnly _timerStartingExecuteSql As Timer

        Private _lastValidSql As String = String.Empty
        Private _lastValidSqlCurrent As String = String.Empty
        Private _errorPosition As Integer = -1
        Private _errorPositionCurrent As Integer = -1

        Private _sql As String

        Public Property IsModified As Boolean

        Public Property UserMetadataStructureItem As MetadataStructureItem

        Public Property FileSourceUrl As String

        Public Property SqlContext As SQLContext

        Public Property QueryText() As String
            Get
                Return SqlQuery.SQL
            End Get
            Set
                SqlQuery.SQL = Value
                _sql = Value
                IsModified = False
                UpdateStateButtons()
            End Set
        End Property

        Public Property SqlFormattingOptions As SQLFormattingOptions


        Public Property SqlGenerationOptions As SQLGenerationOptions

        Public ReadOnly Property QueryView() As QueryView
            Get
                Return QView
            End Get
        End Property

        Public Property IsNeedClose As Boolean

        Public Property SqlQuery As SQLQuery

        Public Property SqlSourceType As Helpers.SourceType

        Public ReadOnly Property FormattedQueryText() As String
            Get
                Return If(Not SqlQuery.QueryRoot.SleepMode, FormattedSQLBuilder.GetSQL(SqlQuery.QueryRoot, SqlFormattingOptions), SqlQuery.SQL)
            End Get
        End Property

        Public Sub New(sqlContext1 As SQLContext)
            Init()

            SqlContext = sqlContext1

            SqlSourceType = Helpers.SourceType.[New]

            SqlQuery = New SQLQuery(SqlContext)

            AddHandler SqlQuery.SQLUpdated, AddressOf SqlQuery_SQLUpdated
            SqlQuery.QueryRoot.AllowSleepMode = True
            AddHandler SqlQuery.QueryAwake, AddressOf SqlQueryOnQueryAwake
            AddHandler SqlQuery.SleepModeChanged, AddressOf SqlQuery_SleepModeChanged

            QueryView.Query = SqlQuery
            BoxSql.Query = SqlQuery
            BoxSqlCurrentSubQuery.Query = SqlQuery

            BoxSql.ExpressionContext = QueryView.ActiveUnionSubQuery
            BoxSqlCurrentSubQuery.ExpressionContext = QueryView.ActiveUnionSubQuery

            AddHandler QueryView.ActiveUnionSubQueryChanged, AddressOf ActiveUnionSubQueryChanged

            _transformerSql = New QueryTransformer()

            _timerStartingExecuteSql = New Timer(AddressOf TimerStartingExecuteSql_Elapsed)

            ' Options to present the formatted SQL query text to end-user
            ' Use names of virtual objects, do not replace them with appropriate derived tables
            SqlFormattingOptions = New SQLFormattingOptions() With {
                .ExpandVirtualObjects = False
            }

            ' Options to generate the SQL query text for execution against a database server
            ' Replace virtual objects with derived tables
            SqlGenerationOptions = New SQLGenerationOptions() With {
                .ExpandVirtualObjects = True
            }

            NavigationBar.QueryView = QueryView
            NavigationBar.Query = SqlQuery

            CBuilder.QueryTransformer = New QueryTransformer() With {
                .Query = SqlQuery,
                .SQLGenerationOptions = SqlFormattingOptions
            }

            AddHandler CBuilder.QueryTransformer.SQLUpdated, AddressOf QueryTransformer_SQLUpdated

            DataGridResult.QueryTransformer = CBuilder.QueryTransformer
            DataGridResult.SqlQuery = SqlQuery

            ' The pagination panel is displayed if the current SyntaxProvider has support for pagination
            PaginationPanel.Visibility = If((CBuilder.QueryTransformer.IsSupportLimitCount OrElse CBuilder.QueryTransformer.IsSupportLimitOffset) AndAlso SqlContext.SyntaxProvider IsNot Nothing, Visibility.Visible, Visibility.Collapsed)


            PaginationPanel.IsSupportLimitCount = CBuilder.QueryTransformer.IsSupportLimitCount
            PaginationPanel.IsSupportLimitOffset = CBuilder.QueryTransformer.IsSupportLimitOffset

            UpdateStateButtons()
        End Sub

        Private Sub ActiveUnionSubQueryChanged()
            BoxSql.ExpressionContext = QueryView.ActiveUnionSubQuery
            BoxSqlCurrentSubQuery.ExpressionContext = QueryView.ActiveUnionSubQuery
            ' ReSharper disable once VBWarnings::BC42105,BC42106,BC42107
        End Sub

        Private Sub SqlQuery_SQLUpdated(sender As Object, e As EventArgs)
            IsModified = _sql <> QueryText

            BoxSql.Text = FormattedQueryText

            SetSqlTextCurrentSubQuery()

            UpdateStateButtons()

            If Not TabItemFastResult.IsSelected OrElse CheckBoxAutoRefreash.IsChecked = False Then
                Return
            End If

            _timerStartingExecuteSql.Change(600, Timeout.Infinite)


        End Sub

        Public Sub OpenExecuteTab()
            TabItemData.IsSelected = True
        End Sub

        Public Sub ShowQueryStatistics()
            Dim qs As QueryStatistics = QueryView.Query.QueryStatistics

            Dim stats As String = "Used Objects (" & qs.UsedDatabaseObjects.Count & "):" & vbCr & vbLf
            stats = qs.UsedDatabaseObjects.Aggregate(stats, Function(current, t) current & (vbCr & vbLf & t.ObjectName.QualifiedName))

            stats += vbCr & vbLf & vbCr & vbLf & "Used Columns (" & qs.UsedDatabaseObjectFields.Count & "):" & vbCr & vbLf
            stats = qs.UsedDatabaseObjectFields.Aggregate(stats, Function(current, t) current & (vbCr & vbLf & t.ObjectName.QualifiedName))

            stats += vbCr & vbLf & vbCr & vbLf & "Output Expressions (" & qs.OutputColumns.Count & "):" & vbCr & vbLf
            stats = qs.OutputColumns.Aggregate(stats, Function(current, t) current & (vbCr & vbLf & t.Expression))

            Dim f As QueryStatisticsWindow = New QueryStatisticsWindow()
            f.textBox.Text = stats

            f.ShowDialog()
        End Sub

        Public Function CanAddObject() As Boolean
            Return QueryView.AddObjectDialog IsNot Nothing
        End Function

        Public Function CanShowProperties() As Boolean
            Return QueryView.ActiveUnionSubQuery IsNot Nothing
        End Function

        Public Function CanAddUnionSubQuery() As Boolean
            If SqlContext.SyntaxProvider Is Nothing Then
                Return False
            End If

            If QueryView.ActiveUnionSubQuery IsNot Nothing AndAlso QueryView.ActiveUnionSubQuery.QueryRoot.IsSubQuery Then
                Return SqlContext.SyntaxProvider.IsSupportSubQueryUnions()
            End If

            Return SqlContext.SyntaxProvider.IsSupportUnions()
        End Function

        Public Function CanCopyUnionSubQuery() As Boolean
            Return CanAddUnionSubQuery()
        End Function

        Public Function CanAddDerivedTable() As Boolean
            If SqlContext.SyntaxProvider Is Nothing Then
                Return False
            End If

            If QueryView.ActiveUnionSubQuery IsNot Nothing AndAlso QueryView.ActiveUnionSubQuery.QueryRoot.IsMainQuery Then
                Return SqlContext.SyntaxProvider.IsSupportDerivedTables()
            End If

            Return SqlContext.SyntaxProvider.IsSupportSubQueryDerivedTables()
        End Function

        Public Function CanCopy() As Boolean
            If Not GetCurrentEditor().IsFocused Then
                Return False
            End If
            Return GetCurrentEditor().SelectionLength > 0
        End Function

        Public Function CanCut() As Boolean
            If Not GetCurrentEditor().IsFocused Then
                Return False
            End If

            Return Not String.IsNullOrEmpty(GetCurrentEditor().SelectedText)
        End Function

        Public Function CanPaste() As Boolean
            Return GetCurrentEditor().IsFocused AndAlso Clipboard.ContainsText()
        End Function

        Public Function CanUndo() As Boolean
            Return GetCurrentEditor().IsFocused AndAlso GetCurrentEditor().CanUndo
        End Function

        Public Function CanRedo() As Boolean
            Return GetCurrentEditor().IsFocused AndAlso GetCurrentEditor().CanRedo
        End Function

        Public Function CanSelectAll() As Boolean
            Return GetCurrentEditor().IsFocused AndAlso GetCurrentEditor().Text.Length > 0
        End Function

        Public Sub Undo()
            GetCurrentEditor().Undo()
        End Sub

        Public Sub Redo()
            GetCurrentEditor().Undo()
        End Sub

        Public Sub Cut()
            GetCurrentEditor().Cut()
        End Sub

        Public Sub Copy()
            GetCurrentEditor().Copy()
        End Sub

        Public Sub Paste()
            GetCurrentEditor().Paste()
        End Sub

        Public Sub SelectAll()
            GetCurrentEditor().SelectAll()
        End Sub

        Private Function GetCurrentEditor() As SqlTextEditor
            If TabItemQueryText.IsSelected Then
                Return BoxSql
            End If

            If TabItemQueryText.IsSelected Then
                Return BoxSqlCurrentSubQuery
            End If


            Return BoxSql
        End Function

        Public Sub AddDerivedTable()
            Using New UpdateRegion(QueryView.ActiveUnionSubQuery.FromClause)
                Dim sqlContext1 As SQLContext = QueryView.ActiveUnionSubQuery.SQLContext

                Dim fq As SQLFromQuery = New SQLFromQuery(sqlContext1) With {
                    .[Alias] = New SQLAliasObjectAlias(sqlContext1) With {
                        .[Alias] = QueryView.ActiveUnionSubQuery.QueryRoot.CreateUniqueSubQueryName()
                    },
                    .SubQuery = New SQLSubSelectStatement(sqlContext1)
                }

                Dim sqse As SQLSubQuerySelectExpression = New SQLSubQuerySelectExpression(sqlContext1)
                fq.SubQuery.Add(sqse)
                sqse.SelectItems = New SQLSelectItems(sqlContext1)
                sqse.From = New SQLFromClause(sqlContext1)

                NavigationBar.Query.AddObject(QueryView.ActiveUnionSubQuery, fq, GetType(DataSourceQuery))
            End Using
        End Sub

        Public Sub CopyUnionSubQuery()
            ' add an empty UnionSubQuery
            Dim usq As UnionSubQuery = QueryView.ActiveUnionSubQuery.ParentGroup.Add()

            ' copy the content of existing union sub-query to a new one
            Dim usqAst As SQLSubQuerySelectExpression = QueryView.ActiveUnionSubQuery.ResultQueryAST
            usqAst.RestoreColumnPrefixRecursive(True)

            Dim lCte As List(Of SQLWithClauseItem) = New List(Of SQLWithClauseItem)()
            Dim lFromObj As List(Of SQLFromSource) = New List(Of SQLFromSource)()
            QueryView.ActiveUnionSubQuery.GatherPrepareAndFixupContext(lCte, lFromObj, False)
            usqAst.PrepareAndFixupRecursive(lCte, lFromObj)

            usq.LoadFromAST(usqAst)
            QueryView.ActiveUnionSubQuery = usq
        End Sub

        Public Sub AddUnionSubQuery()
            QueryView.ActiveUnionSubQuery = QueryView.ActiveUnionSubQuery.ParentGroup.Add()
        End Sub

        Public Sub PropertiesQuery()
            QueryView.ShowActiveUnionSubQueryProperties()
        End Sub

        Public Sub AddObject()
            If QueryView.AddObjectDialog IsNot Nothing Then
                QueryView.AddObjectDialog.ShowModal()
            End If
        End Sub

        Private Sub BoxSql_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            Try
                ' Update the query builder with manually edited query text:
                SqlQuery.SQL = BoxSql.Text

                ' Hide error banner if any
                ErrorBox.Visibility = Visibility.Collapsed
                _lastValidSql = SqlQuery.SQL
            Catch ex As SQLParsingException
                ' Show banner with error text
                _errorPosition = ex.ErrorPos.pos
                ErrorBox.Show(ex.Message, SqlContext.SyntaxProvider)
            End Try
        End Sub


        Private Sub TimerStartingExecuteSql_Elapsed(state As Object)
            Dispatcher.BeginInvoke(DirectCast(AddressOf FillFastResult, Action))
        End Sub

        Private Shared Sub SqlQueryOnQueryAwake(sender As QueryRoot, ByRef abort As Boolean)
            If MessageBox.Show("You had typed something that is not a SELECT statement in the text editor and continued with visual query building." & "Whatever the text in the editor is, it will be replaced with the SQL generated by the component. Is it right?", "Active Query Builder .NET Demo", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                abort = True
            End If
        End Sub

        Private Sub SqlQuery_SleepModeChanged(sender As Object, e As EventArgs)
            BorderSleepMode.Visibility = If(SqlQuery.SleepMode, Visibility.Visible, Visibility.Collapsed)
            TabItemFastResult.Visibility = If(SqlQuery.SleepMode, Visibility.Collapsed, Visibility.Visible)
            TabItemCurrentSubQuery.Visibility = If(SqlQuery.SleepMode, Visibility.Collapsed, Visibility.Visible)
            TabItemData.IsEnabled = Not SqlQuery.SleepMode
        End Sub

        Private Sub UpdateStateButtons()
            Dim value As Boolean = SqlQuery IsNot Nothing AndAlso SqlQuery.SQL.Length > 0

            ButtonSave.IsEnabled = value AndAlso IsModified
            ButtonSaveCurrentSubQuery.IsEnabled = value AndAlso IsModified
            ButtonSaveToFileAs.IsEnabled = value
        End Sub

        Public Sub New()
            Init()
        End Sub

        Private Sub Init()
            InitializeComponent()

            AddHandler Loaded, AddressOf ContentWindowChild_Loaded
            Dim langProperty As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(LanguageProperty, GetType(ContentWindowChild))
            langProperty.AddValueChanged(Me, AddressOf LanguageChanged)
        End Sub

        Private Sub QueryTransformer_SQLUpdated(sender As Object, e As EventArgs)
            ' Handle the event raised by Query Transformer object that the text of SQL query is changed
            ' update the text box
            If DataGridResult.QueryTransformer Is Nothing OrElse Not TabItemData.IsSelected OrElse BoxSqlTransformer.Text = DataGridResult.QueryTransformer.SQL Then
                Return
            End If

            BoxSqlTransformer.Text = DataGridResult.QueryTransformer.SQL

            DataGridResult.FillDataGrid(DataGridResult.QueryTransformer.SQL)

        End Sub

        Private Sub LanguageChanged(sender As Object, e As EventArgs)
            DockPanelPropertiesBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strPropertiesBarCaption",
                                                                                                   ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language),
                                                                                                   LocalizableConstantsUI.strPropertiesBarCaption)
            DockPanelSubQueryNavBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strSubQueryStructureBarCaption",
                                                                                                    ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language),
                                                                                                    LocalizableConstantsUI.strSubQueryStructureBarCaption)
        End Sub

        Private Sub ContentWindowChild_Loaded(sender As Object, e As RoutedEventArgs)
            DockPanelPropertiesBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strPropertiesBarCaption",
                                                                                                   ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language),
                                                                                                   LocalizableConstantsUI.strPropertiesBarCaption)
            DockPanelSubQueryNavBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strSubQueryStructureBarCaption",
                                                                                                    ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language),
                                                                                                    LocalizableConstantsUI.strSubQueryStructureBarCaption)
        End Sub

#Region "Menu buttons event method"

        Private Sub ButtonProperties_OnClick(sender As Object, e As RoutedEventArgs)
            PropertiesQuery()
        End Sub

        Private Sub ButtonAddObject_OnClick(sender As Object, e As RoutedEventArgs)
            AddObject()
        End Sub

        Private Sub ButtonAddDerivedTable_OnClick(sender As Object, e As RoutedEventArgs)
            AddDerivedTable()
        End Sub

        Private Sub ButtonNewUnionSubQuery_OnClick(sender As Object, e As RoutedEventArgs)
            AddUnionSubQuery()
        End Sub

        Private Sub ButtonCopyUnionSubQuery_OnClick(sender As Object, e As RoutedEventArgs)
            CopyUnionSubQuery()
        End Sub

        Private Sub ButtonSave_OnClick(sender As Object, e As RoutedEventArgs)
            OnSaveQueryEvent()
            UpdateStateButtons()
        End Sub

        Private Sub ButtonSaveToFileAs_OnClick(sender As Object, e As RoutedEventArgs)
            OnSaveAsInFileEvent()
            UpdateStateButtons()
        End Sub

        Private Sub ButtonSaveCurrentSubQuery_OnClick(sender As Object, e As RoutedEventArgs)
            OnSaveAsNewUserQueryEvent()
            UpdateStateButtons()
        End Sub
#End Region

        Private Sub BoxSql_OnTextChanged(sender As Object, eventArgs As EventArgs)
            ErrorBox.Visibility = Visibility.Collapsed
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

        Private Sub SetSqlTextCurrentSubQuery()
            ErrorBoxCurrent.Visibility = Visibility.Collapsed
            If _transformerSql Is Nothing Then
                Return
            End If

            If QueryView.ActiveUnionSubQuery Is Nothing Or SqlQuery.SleepMode Then
                BoxSqlCurrentSubQuery.Text = String.Empty
                _transformerSql.Query = Nothing
                Return
            End If

            _transformerSql.Query = New SQLQuery(QueryView.ActiveUnionSubQuery.SQLContext) With {
                .SQL = QueryView.ActiveUnionSubQuery.ParentSubQuery.GetSqlForDataPreview()
            }

            Dim sql As String = QueryView.ActiveUnionSubQuery.ParentSubQuery.GetResultSQL(SqlFormattingOptions)
            BoxSqlCurrentSubQuery.Text = sql
            _lastValidSqlCurrent = sql
        End Sub

        Private Sub FillFastResult()
            Dim result As QueryTransformer = _transformerSql.Take("10")

            Try
                Dim dv As DataView = Helpers.ExecuteSql(result.SQL, DirectCast(_transformerSql.Query, SQLQuery))

                ListViewFastResultSql.ItemsSource = dv
                BorderError.Visibility = Visibility.Collapsed
            Catch exception As Exception
                BorderError.Visibility = Visibility.Visible
                LabelError.Text = exception.Message

                ListViewFastResultSql.ItemsSource = Nothing
            End Try
        End Sub

        Private Sub ButtonRefreshFastResult_OnClick(sender As Object, e As RoutedEventArgs)
            FillFastResult()
        End Sub

        Private Sub CheckBoxAutoRefresh_OnChecked(sender As Object, e As RoutedEventArgs)
            If ButtonRefreashFastResult Is Nothing OrElse CheckBoxAutoRefreash Is Nothing Then
                Return
            End If
            ButtonRefreashFastResult.IsEnabled = CType((CheckBoxAutoRefreash.IsChecked = False), Boolean)
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
            If QueryView.ActiveUnionSubQuery Is Nothing Then
                Return
            End If
            Try
                ErrorBoxCurrent.Visibility = Visibility.Collapsed

                QView.ActiveUnionSubQuery.ParentSubQuery.SQL = DirectCast(sender, SqlTextEditor).Text

                Dim sql As String = QueryView.ActiveUnionSubQuery.ParentSubQuery.GetResultSQL(SqlFormattingOptions)

                _transformerSql.Query = New SQLQuery(QueryView.ActiveUnionSubQuery.SQLContext) With {
                    .SQL = sql
                }
            Catch ex As SQLParsingException
                ErrorBoxCurrent.Show(ex.Message, SqlContext.SyntaxProvider)
                _errorPositionCurrent = ex.ErrorPos.pos
            End Try
        End Sub

        Protected Friend Overridable Sub OnSaveQueryEvent()
            RaiseEvent SaveQueryEvent(Me, EventArgs.Empty)
        End Sub

        Protected Friend Overridable Sub OnSaveAsInFileEvent()
            RaiseEvent SaveAsInFileEvent(Me, EventArgs.Empty)
        End Sub

        Protected Friend Overridable Sub OnSaveAsNewUserQueryEvent()
            RaiseEvent SaveAsNewUserQueryEvent(Me, EventArgs.Empty)
        End Sub

        Private Sub TabControlMain_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            ' Execute a query on switching to the Data tab
            If SqlContext Is Nothing OrElse (SqlContext.SyntaxProvider Is Nothing OrElse Not TabItemData.IsSelected) Then
                DataGridResult.StopFill()
                Return
            End If

            ResetPagination()
            BoxSqlTransformer.Text = BoxSql.Text

            If Not TabItemData.IsSelected Then
                Return
            End If

            BorderBlockPagination.Visibility = Visibility.Visible
            DataGridResult.FillDataGrid(BoxSql.Text)
        End Sub

        Private Sub TabControlSqlText_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If ErrorBox.Visibility = Visibility.Visible Then
                ErrorBox.Visibility = Visibility.Collapsed
            End If
            If Not e.AddedItems.Contains(TabItemFastResult) OrElse _transformerSql.Query Is Nothing OrElse CheckBoxAutoRefreash.IsChecked = False Then
                Return
            End If

            FillFastResult()
        End Sub

        Private Sub ErrorBox_OnGoToErrorPosition(sender As Object, e As EventArgs)
            BoxSql.Focus()

            If _errorPosition <> -1 Then
                BoxSql.ScrollToPosition(_errorPosition)
                BoxSql.CaretOffset = _errorPosition
            End If
        End Sub

        Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
            BoxSql.Text = _lastValidSql
            BoxSql.Focus()
        End Sub

        Private Sub ErrorBoxCurrent_OnGoToErrorPosition(sender As Object, e As EventArgs)
            BoxSqlCurrentSubQuery.Focus()

            If _errorPositionCurrent <> -1 Then
                BoxSqlCurrentSubQuery.ScrollToPosition(_errorPositionCurrent)
                BoxSqlCurrentSubQuery.CaretOffset = _errorPositionCurrent
            End If
        End Sub

        Private Sub ErrorBoxCurrent_OnRevertValidText(sender As Object, e As EventArgs)
            BoxSqlCurrentSubQuery.Text = _lastValidSqlCurrent
            BoxSqlCurrentSubQuery.Focus()
        End Sub
    End Class
End Namespace
