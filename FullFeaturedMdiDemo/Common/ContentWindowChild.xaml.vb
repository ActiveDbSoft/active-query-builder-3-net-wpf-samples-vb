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
Imports ActiveQueryBuilder.Core.QueryTransformer
Imports ActiveQueryBuilder.View.QueryView
Imports ActiveQueryBuilder.View.WPF.ExpressionEditor
Imports ActiveQueryBuilder.View.WPF.QueryView
Imports FullFeaturedMdiDemo.CommonWindow
Imports FullFeaturedMdiDemo.Reports
Imports Microsoft.Win32

Imports SQLParsingException = ActiveQueryBuilder.Core.SQLParsingException
Imports GeneralAssembly
Imports GeneralAssembly.Windows.QueryInformationWindows

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

        Private _sql As String

        Public Property BehaviorOptions() As BehaviorOptions
            Get
                Return SqlQuery.BehaviorOptions
            End Get
            Set(value As BehaviorOptions)
                SqlQuery.BehaviorOptions = value
            End Set
        End Property

        Public Property DataSourceOptions() As DataSourceOptions
            Get
                Return CType(DPaneControl.DataSourceOptions, DataSourceOptions)
            End Get
            Set(value As DataSourceOptions)
                DPaneControl.DataSourceOptions = value
            End Set
        End Property

        Public Property DesignPaneOptions() As DesignPaneOptions
            Get
                Return DPaneControl.Options
            End Get
            Set(value As DesignPaneOptions)
                DPaneControl.Options = value
            End Set
        End Property

        Public Property QueryNavBarOptions() As QueryNavBarOptions
            Get
                Return NavigationBar.Options
            End Get
            Set(value As QueryNavBarOptions)
                RemoveHandler NavigationBar.Options.Updated, AddressOf QueryNavBarOptionsOnUpdated
                NavigationBar.Options.Assign(value)
                AddHandler NavigationBar.Options.Updated, AddressOf QueryNavBarOptionsOnUpdated
                SubQueryNavBarControl.Options.Assign(DirectCast(NavigationBar.Options, IQueryNavigationBarOptions))
            End Set
        End Property

        Private Sub QueryNavBarOptionsOnUpdated(sender As Object, e As EventArgs)
            SubQueryNavBarControl.Options.Assign(DirectCast(QueryNavBarOptions, IQueryNavigationBarOptions))
        End Sub

        Public Property MetadataLoadingOptions() As MetadataLoadingOptions
            Get
                Return SqlContext.LoadingOptions
            End Get
            Set(value As MetadataLoadingOptions)
                SqlContext.LoadingOptions = value
            End Set
        End Property

        Public Property MetadataStructureOptions() As MetadataStructureOptions
            Get
                Return SqlContext.MetadataStructureOptions
            End Get
            Set(value As MetadataStructureOptions)
                SqlContext.MetadataStructureOptions = value
            End Set
        End Property

        '        
        '        public AddObjectDialogOptions AddObjectDialogOptions
        '        {
        '            get { return QueryView; }
        '            set { QueryView.AddObjectDialog.Options = value; }
        '        }

        Public Property UserInterfaceOptions() As UserInterfaceOptions
            Get
                Return QView.UserInterfaceOptions
            End Get
            Set(value As UserInterfaceOptions)
                QView.UserInterfaceOptions = value
            End Set
        End Property

        Public Property ExpressionEditorOptions() As ExpressionEditorOptions
            Get
                Return ExpressionEditor.Options
            End Get
            Set(value As ExpressionEditorOptions)
                ExpressionEditor.Options = value
            End Set
        End Property

        Public Property TextEditorOptions() As TextEditorOptions
            Get
                Return BoxSql.Options
            End Get
            Set(value As TextEditorOptions)
                ExpressionEditor.TextEditorOptions = value
                BoxSql.Options = value
                BoxSqlCurrentSubQuery.Options = value
            End Set
        End Property

        Public Property TextEditorSqlOptions() As SqlTextEditorOptions
            Get
                Return BoxSql.SqlOptions
            End Get
            Set(value As SqlTextEditorOptions)
                ExpressionEditor.TextEditorSqlOptions = value
                BoxSql.SqlOptions = value
                BoxSqlCurrentSubQuery.SqlOptions = value
            End Set
        End Property

        Public Property VisualOptions() As VisualOptions
            Get
                Return DockManager.Options
            End Get
            Set(value As VisualOptions)
                DockManager.Options = value
            End Set
        End Property

        Public Property QueryColumnListOptions() As QueryColumnListOptions
            Get
                Return ColumnListControl.Options
            End Get
            Set(value As QueryColumnListOptions)
                ColumnListControl.Options = value
            End Set
        End Property

        Public Property IsModified() As Boolean

        Public Property UserMetadataStructureItem() As MetadataStructureItem
        Public Property FileSourceUrl() As String

        Private _privateSqlContext As SQLContext
        Public Property SqlContext() As SQLContext
            Get
                Return _privateSqlContext
            End Get
            Private Set(value As SQLContext)
                _privateSqlContext = value
            End Set
        End Property


        Public Property QueryText() As String
            Get
                Return SqlQuery.SQL
            End Get
            Set(value As String)
                SqlQuery.SQL = value
                _sql = value
                IsModified = False
                UpdateStateButtons()
            End Set
        End Property

        Private _sqlFormattingOptions As SQLFormattingOptions
        Private _errorPosition As Integer = -1
        Private _lastValidSql As String
        Private _lastValidSqlCurrentSubQuery As String
        Private _errorPositionCurrentSubQuery As Integer = -1

        Public Property SqlFormattingOptions() As SQLFormattingOptions
            Get
                Return _sqlFormattingOptions
            End Get
            Set(value As SQLFormattingOptions)
                If _sqlFormattingOptions IsNot Nothing Then
                    RemoveHandler _sqlFormattingOptions.Updated, AddressOf SqlFormattingOptionsOnUpdated
                End If

                _sqlFormattingOptions = value

                If _sqlFormattingOptions IsNot Nothing Then
                    AddHandler _sqlFormattingOptions.Updated, AddressOf SqlFormattingOptionsOnUpdated
                End If
                CBuilder.QueryTransformer.SQLGenerationOptions = value
            End Set
        End Property

        Private Sub SqlFormattingOptionsOnUpdated(sender As Object, eventArgs As EventArgs)
            BoxSql.Text = FormattedQueryText
        End Sub

        Public Property SqlGenerationOptions() As SQLGenerationOptions
            Get
                Return QueryView.SQLGenerationOptions
            End Get
            Set(value As SQLGenerationOptions)
                QueryView.SQLGenerationOptions = value
            End Set
        End Property

        Public ReadOnly Property QueryView() As QueryView
            Get
                Return QView
            End Get
        End Property

        Public Property IsNeedClose() As Boolean

        Private privateSqlQuery As SQLQuery
        Public Property SqlQuery() As SQLQuery
            Get
                Return privateSqlQuery
            End Get
            Private Set(value As SQLQuery)
                privateSqlQuery = value
            End Set
        End Property

        Public Property SqlSourceType() As SourceType

        Public ReadOnly Property FormattedQueryText() As String
            Get
                Return If(Not SqlQuery.QueryRoot.SleepMode, FormattedSQLBuilder.GetSQL(SqlQuery.QueryRoot, SqlFormattingOptions), SqlQuery.SQL)
            End Get
        End Property

        Public Property AddObjectDialogOptions() As AddObjectDialogOptions
            Get
                Return DirectCast(QueryView.AddObjectDialog, AddObjectDialog).Options
            End Get
            Set(value As AddObjectDialogOptions)
                DirectCast(QueryView.AddObjectDialog, AddObjectDialog).Options.Assign(value)
            End Set
        End Property

        Public Sub New(sqlContext As SQLContext)
            Init()

            Me.SqlContext = sqlContext

            SqlSourceType = SourceType.[New]

            SqlQuery = New SQLQuery(Me.SqlContext)

            AddHandler SqlQuery.SQLUpdated, AddressOf SqlQuery_SQLUpdated
            SqlQuery.QueryRoot.AllowSleepMode = True
            AddHandler SqlQuery.QueryAwake, AddressOf SqlQueryOnQueryAwake
            AddHandler SqlQuery.SleepModeChanged, AddressOf SqlQuery_SleepModeChanged
            NavigationBar.QueryView = QueryView
            QueryView.Query = SqlQuery

            AddHandler QueryView.ActiveUnionSubQueryChanged, AddressOf QueryView_ActiveUnionSubQueryChanged

            BoxSql.Query = SqlQuery
            BoxSqlCurrentSubQuery.Query = SqlQuery
            BoxSqlCurrentSubQuery.ExpressionContext = QueryView.ActiveUnionSubQuery

            AddHandler QueryView.ActiveUnionSubQueryChanged, Sub()
                                                                 BoxSqlCurrentSubQuery.ExpressionContext = QueryView.ActiveUnionSubQuery
                                                             End Sub

            _transformerSql = New QueryTransformer()

            _timerStartingExecuteSql = New Timer(AddressOf TimerStartingExecuteSql_Elapsed)

            CBuilder.QueryTransformer = New QueryTransformer With {.Query = SqlQuery}

            ' Options to present the formatted SQL query text to end-user
            ' Use names of virtual objects, do not replace them with appropriate derived tables
            SqlFormattingOptions = New SQLFormattingOptions With {.ExpandVirtualObjects = False}

            ' Options to generate the SQL query text for execution against a database server
            ' Replace virtual objects with derived tables
            SqlGenerationOptions = New SQLGenerationOptions With {.ExpandVirtualObjects = True}

            NavigationBar.QueryView = QueryView
            NavigationBar.Query = SqlQuery

            AddHandler CBuilder.QueryTransformer.SQLUpdated, AddressOf QueryTransformer_SQLUpdated

            DataViewerResult.QueryTransformer = CBuilder.QueryTransformer
            DataViewerResult.SqlQuery = SqlQuery

            UpdateStateButtons()
        End Sub

        Private Sub QueryView_ActiveUnionSubQueryChanged(sender As Object, e As EventArgs)
            SetSqlTextCurrentSubQuery()
        End Sub

        Private Sub SqlQuery_SQLUpdated(sender As Object, e As EventArgs)
            _lastValidSql = FormattedQueryText

            IsModified = _sql <> QueryText

            BoxSql.Text = FormattedQueryText

            SetSqlTextCurrentSubQuery()

            UpdateStateButtons()
            CheckParameters()

            Dim isEnable As Boolean = Not String.IsNullOrEmpty(FormattedQueryText) AndAlso Not IsNothing(SqlContext.MetadataProvider)
            ButtonCVS.IsEnabled = isEnable
            ButtonExcel.IsEnabled = isEnable
            ButtonReport.IsEnabled = isEnable

            If Not TabItemFastResult.IsSelected OrElse CheckBoxAutoRefreash.IsChecked = False Then
                Return
            End If

            _timerStartingExecuteSql.Change(600, Timeout.Infinite)
        End Sub

        Private Sub CheckParameters()
            If SqlHelpers.CheckParameters(SqlContext.MetadataProvider, SqlContext.SyntaxProvider, SqlQuery.QueryParameters) Then
                HideParametersErrorPanel()
            Else
                Dim acceptableFormats = SqlHelpers.GetAcceptableParametersFormats(SqlContext.MetadataProvider, SqlContext.SyntaxProvider)
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

        Public Sub OpenExecuteTab()
            TabItemData.IsSelected = True
        End Sub

        Public Sub ShowQueryStatistics()
            Dim qs = QueryView.Query.QueryStatistics

            Dim stats = "Used Objects (" & qs.UsedDatabaseObjects.Count & "):" & ControlChars.CrLf
            stats = qs.UsedDatabaseObjects.Aggregate(stats, Function(current, t) current + (ControlChars.CrLf & t.ObjectName.QualifiedName))

            stats &= ControlChars.CrLf & ControlChars.CrLf & "Used Columns (" & qs.UsedDatabaseObjectFields.Count & "):" & ControlChars.CrLf
            stats = qs.UsedDatabaseObjectFields.Aggregate(stats, Function(current, t) current + (ControlChars.CrLf & t.ObjectName.QualifiedName))

            stats &= ControlChars.CrLf & ControlChars.CrLf & "Output Expressions (" & qs.OutputColumns.Count & "):" & ControlChars.CrLf
            stats = qs.OutputColumns.Aggregate(stats, Function(current, t) current + (ControlChars.CrLf & t.Expression))

            Dim f = New QueryStatisticsWindow(stats)

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
            Using tempVar As New UpdateRegion(QueryView.ActiveUnionSubQuery.FromClause)
                'INSTANT VB NOTE: The variable sqlContext was renamed since Visual Basic does not handle local variables named the same as class members well:
                Dim sqlContext_Renamed = QueryView.ActiveUnionSubQuery.SQLContext

                Dim fq = New SQLFromQuery(sqlContext_Renamed) With {
                    .Alias = New SQLAliasObjectAlias(sqlContext_Renamed) With {.Alias = QueryView.ActiveUnionSubQuery.QueryRoot.CreateUniqueSubQueryName()},
                    .SubQuery = New SQLSubSelectStatement(sqlContext_Renamed)
                }

                Dim sqse = New SQLSubQuerySelectExpression(sqlContext_Renamed)
                fq.SubQuery.Add(sqse)
                sqse.SelectItems = New SQLSelectItems(sqlContext_Renamed)
                sqse.From = New SQLFromClause(sqlContext_Renamed)

                NavigationBar.Query.AddObject(QueryView.ActiveUnionSubQuery, fq, GetType(DataSourceQuery))
            End Using
        End Sub

        Public Sub CopyUnionSubQuery()
            ' add an empty UnionSubQuery
            Dim usq = QueryView.ActiveUnionSubQuery.ParentGroup.Add()

            ' copy the content of existing union sub-query to a new one
            Dim usqAst = QueryView.ActiveUnionSubQuery.ResultQueryAST
            usqAst.RestoreColumnPrefixRecursive(True)

            Dim lCte = New List(Of SQLWithClauseItem)()
            Dim lFromObj = New List(Of SQLFromSource)()
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
                ErrorBox.Show(Nothing, QView.Query.SQLContext.SyntaxProvider)
            Catch ex As SQLParsingException
                ' Show banner with error text
                ErrorBox.Show(ex.Message, QView.Query.SQLContext.SyntaxProvider)
                _errorPosition = ex.ErrorPos.pos
            End Try
        End Sub

        Private Sub TimerStartingExecuteSql_Elapsed(state As Object)
            If Dispatcher IsNot Nothing Then
                Dispatcher.BeginInvoke(CType(AddressOf FillFastResult, Action))
            End If
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
            Dim value = SqlQuery IsNot Nothing AndAlso SqlQuery.SQL.Length > 0

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
            Dim langProperty = DependencyPropertyDescriptor.FromProperty(LanguageProperty, GetType(ContentWindowChild))
            langProperty.AddValueChanged(Me, AddressOf LanguageChanged)

            AddHandler QueryNavBarOptions.Updated, AddressOf QueryNavBarOptionsOnUpdated
        End Sub

        Private Sub QueryTransformer_SQLUpdated(sender As Object, e As EventArgs)
            ' Handle the event raised by Query Transformer object that the text of SQL query is changed
            ' update the text box
            Try
                If DataViewerResult.QueryTransformer Is Nothing OrElse Not TabItemData.IsSelected OrElse BoxSqlTransformer.Text = DataViewerResult.QueryTransformer.SQL Then
                    Return
                End If

                BoxSqlTransformer.Text = DataViewerResult.QueryTransformer.SQL

                DataViewerResult.FillData(DataViewerResult.QueryTransformer.SQL)
            Catch ex As Exception
                'ignore
            End Try

        End Sub

        Private Sub LanguageChanged(sender As Object, e As EventArgs)
            DockPanelPropertiesBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strPropertiesBarCaption", ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language), LocalizableConstantsUI.strProperties)
            DockPanelSubQueryNavBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strSubQueryStructureBarCaption", ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language), LocalizableConstantsUI.strSubQueryStructureBarCaption)
        End Sub

        Private Sub ContentWindowChild_Loaded(sender As Object, e As RoutedEventArgs)
            DockPanelPropertiesBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strPropertiesBarCaption", ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language), LocalizableConstantsUI.strProperties)
            DockPanelSubQueryNavBar.Title = ActiveQueryBuilder.View.WPF.Helpers.Localizer.GetString("strSubQueryStructureBarCaption", ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(Language), LocalizableConstantsUI.strSubQueryStructureBarCaption)
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
            ErrorBox.Show(Nothing, SqlContext.SyntaxProvider)
        End Sub

        Private Sub SetSqlTextCurrentSubQuery()
            ErrorBoxCurrentSunQuery.Show(Nothing, QView.Query.SQLContext.SyntaxProvider)
            If _transformerSql Is Nothing Then
                Return
            End If

            If QueryView.ActiveUnionSubQuery Is Nothing OrElse SqlQuery.SleepMode Then
                BoxSqlCurrentSubQuery.Text = String.Empty
                _transformerSql.Query = Nothing
                Return
            End If


            Dim sqlForDataPreview = QueryView.ActiveUnionSubQuery.ParentSubQuery.GetSqlForDataPreview()
            _transformerSql.Query = New SQLQuery(QueryView.ActiveUnionSubQuery.SQLContext) With {.SQL = sqlForDataPreview}

            Dim sql = QueryView.ActiveUnionSubQuery.ParentSubQuery.GetResultSQL(SqlFormattingOptions)
            BoxSqlCurrentSubQuery.Text = sql
            _lastValidSqlCurrentSubQuery = sql
        End Sub

        Private Sub FillFastResult()
            Dim result = _transformerSql.Take("10")

            dataViewerFastResultSql.FillData(result.SQL, DirectCast(_transformerSql.Query, SQLQuery))
        End Sub

        Private Sub ButtonRefreshFastResult_OnClick(sender As Object, e As RoutedEventArgs)
            FillFastResult()
        End Sub

        Private Sub CheckBoxAutoRefresh_OnChecked(sender As Object, e As RoutedEventArgs)
            If ButtonRefreashFastResult Is Nothing OrElse CheckBoxAutoRefreash Is Nothing Then
                Return
            End If
            ButtonRefreashFastResult.IsEnabled = CheckBoxAutoRefreash.IsChecked.HasValue AndAlso CheckBoxAutoRefreash.IsChecked.Value = False
        End Sub

        Private Sub BoxSqlCurrentSubQuery_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
            If QueryView.ActiveUnionSubQuery Is Nothing Then
                Return
            End If
            Try
                ErrorBoxCurrentSunQuery.Show(Nothing, SqlContext.SyntaxProvider)

                QView.ActiveUnionSubQuery.ParentSubQuery.SQL = DirectCast(sender, SqlTextEditor).Text

                Dim sql = QueryView.ActiveUnionSubQuery.ParentSubQuery.GetResultSQL(SqlFormattingOptions)

                _transformerSql.Query = New SQLQuery(QueryView.ActiveUnionSubQuery.SQLContext) With {.SQL = sql}
            Catch ex As SQLParsingException
                ErrorBoxCurrentSunQuery.Show(ex.Message, SqlContext.SyntaxProvider)
                _errorPositionCurrentSubQuery = ex.ErrorPos.pos
            End Try
        End Sub

        Protected Friend Overridable Sub OnSaveQueryEvent()
            Dim handler = SaveQueryEventEvent
            If handler IsNot Nothing Then
                handler(Me, EventArgs.Empty)
            End If
        End Sub

        Protected Friend Overridable Sub OnSaveAsInFileEvent()
            Dim handler = SaveAsInFileEventEvent
            If handler IsNot Nothing Then
                handler(Me, EventArgs.Empty)
            End If
        End Sub

        Protected Friend Overridable Sub OnSaveAsNewUserQueryEvent()
            Dim handler = SaveAsNewUserQueryEventEvent
            If handler IsNot Nothing Then
                handler(Me, EventArgs.Empty)
            End If
        End Sub

        Private Sub TabControlMain_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            ' Execute a query on switching to the Data tab
            If SqlContext Is Nothing OrElse (SqlContext.SyntaxProvider Is Nothing OrElse Not TabItemData.IsSelected) Then
                Return
            End If

            BoxSqlTransformer.Text = BoxSql.Text

            If Not TabItemData.IsSelected Then
                Return
            End If

            DataViewerResult.FillData(CBuilder.SQL)
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

            If _errorPosition = -1 Then
                Return
            End If

            BoxSql.ScrollToPosition((_errorPosition))
            BoxSql.CaretOffset = _errorPosition
        End Sub

        Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
            BoxSql.Text = _lastValidSql
            BoxSql.Focus()
        End Sub

        Private Sub ErrorBoxCurrentSunQuery_OnGoToErrorPosition(sender As Object, e As EventArgs)
            BoxSqlCurrentSubQuery.Focus()

            If _errorPositionCurrentSubQuery = -1 Then
                Return
            End If

            BoxSqlCurrentSubQuery.ScrollToPosition((_errorPositionCurrentSubQuery))
            BoxSqlCurrentSubQuery.CaretOffset = _errorPositionCurrentSubQuery
        End Sub

        Private Sub ErrorBoxCurrentSunQuery_OnRevertValidText(sender As Object, e As EventArgs)
            BoxSqlCurrentSubQuery.Text = _lastValidSql
            BoxSqlCurrentSubQuery.Focus()
        End Sub

        Private Sub CreateFastReport(dataTable As DataTable)
            If IsNothing(dataTable) Then
                Throw New ArgumentException("Argument cannot be null or empty.", "DataTable")
            End If
            Dim reportWindow As FastReportWindow = New FastReportWindow(dataTable) With {
                .Owner = ActiveQueryBuilder.View.WPF.Helpers.FindVisualParent(Of Window)(Me)
            }
            reportWindow.ShowDialog()
        End Sub

        Private Sub CreateStimulsoftReport(dataTable As DataTable)
            If IsNothing(dataTable) Then
                Throw New ArgumentException("Argument cannot be null or empty.", "DataTable")
            End If
#If ENABLE_REPORTSNET_SUPPORT Then
            Dim reportWindow As StimulsoftWindow = New StimulsoftWindow(dataTable) With {
                .Owner = ActiveQueryBuilder.View.WPF.Helpers.FindVisualParent(Of Window)(Me)
            }
            reportWindow.ShowDialog()
#Else
            MessageBox.Show("To test the integration with Stimulsoft Reports.NET, please open the ""Directory.Build.props"" file in the demo projects installation directory (usually ""%USERPROFILE%\Documents\Active Query Builder x.x .NET Examples"") with a text editor and set the ""EnableSupportReportsNet"" flag to true. Then, open the Active Query Builder Demos solution with your IDE, compile and run the Full-featured MDI demo." & vbCrLf & vbCrLf & "
You may also need to activate the trial version of Reports.NET report on the Stimulsoft website.", "Reports.NET support", MessageBoxButton.OK, MessageBoxImage.Information)
#End If
        End Sub
        Private Sub CreateActiveReport(dataTable As DataTable)
            If IsNothing(dataTable) Then
                Throw New ArgumentException("Argument cannot be null or empty.", "DataTable")
            End If
#If ENABLE_ACTIVEREPORTS_SUPPORT Then
            Dim reportWindow As ActiveReportsWindow = New ActiveReportsWindow(dataTable) With {
                .Owner = ActiveQueryBuilder.View.WPF.Helpers.FindVisualParent(Of Window)(Me),
                .ShowInTaskbar = False
            }
            reportWindow.ShowDialog()
#Else
            MessageBox.Show("To test the integration with GrapeCity ActiveReports, please open the ""Directory.Build.props"" file in the demo projects installation directory (usually ""%USERPROFILE%\Documents\Active Query Builder x.x .NET Examples"") with a text editor and set the ""EnableSupportActiveReports"" flag to true. Then, open the Active Query Builder Demos solution with your IDE, compile and run the Full-featured MDI demo." & vbCrLf & vbCrLf & "
You may also need to activate the trial version of ActiveReports on the GrapeCity website.", "ActiveReports support", MessageBoxButton.OK, MessageBoxImage.Information)
#End If

        End Sub
        Private Sub GenerateReport_OnClick(sender As Object, e As RoutedEventArgs)
            Dim window As CreateReportWindow = New CreateReportWindow() With {
                .Owner = ActiveQueryBuilder.View.WPF.Helpers.FindVisualParent(Of Window)(Me)
            }
            Dim result As Boolean? = window.ShowDialog()
            If result <> True OrElse IsNothing(window.SelectedReportType) Then
                Return
            End If
            Dim dataTable As DataTable = SqlHelpers.GetDataTable(CBuilder.QueryTransformer.ResultAST.GetSQL(SqlGenerationOptions), SqlQuery)
            Select Case window.SelectedReportType
                Case ReportType.ActiveReports14
                    CreateActiveReport(dataTable)
                    Exit Select
                Case ReportType.Stimulsoft
                    CreateStimulsoftReport(dataTable)
                    Exit Select
                Case ReportType.FastReport
                    CreateFastReport(dataTable)
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select
        End Sub
        Private Sub ExportToExcel_OnClick(sender As Object, e As RoutedEventArgs)
            Dim dt As DataTable = SqlHelpers.GetDataTable(CBuilder.QueryTransformer.ResultAST.GetSQL(SqlGenerationOptions), SqlQuery)
            Dim saveDialog As SaveFileDialog = New SaveFileDialog() With {
                .AddExtension = True,
                .DefaultExt = "xlsx",
                .FileName = "Export.xlsx"
            }
            If saveDialog.ShowDialog(ActiveQueryBuilder.View.WPF.Helpers.FindVisualParent(Of Window)(Me)) <> True Then
                Return
            End If
            ExportHelpers.ExportToExcel(dt, saveDialog.FileName)
        End Sub
        Private Sub ExportToCSV_OnClick(sender As Object, e As RoutedEventArgs)
            Dim saveDialog As SaveFileDialog = New SaveFileDialog() With {
                .AddExtension = True,
                .DefaultExt = "csv",
                .FileName = "Data.csv"
            }
            Dim result As Boolean? = saveDialog.ShowDialog(ActiveQueryBuilder.View.WPF.Helpers.FindVisualParent(Of Window)(Me))
            If result <> True Then
                Return
            End If
            Dim dt As DataTable = SqlHelpers.GetDataTable(CBuilder.QueryTransformer.ResultAST.GetSQL(SqlGenerationOptions), SqlQuery)
            ExportHelpers.ExportToCSV(dt, saveDialog.FileName)
        End Sub
    End Class
End Namespace
