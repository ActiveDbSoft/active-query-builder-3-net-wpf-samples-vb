''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports ActiveQueryBuilder.View.ExpressionEditor
Imports ActiveQueryBuilder.View.QueryView
Imports ActiveQueryBuilder.View.WPF.DatabaseSchemaView
Imports ActiveQueryBuilder.View.WPF.ExpressionEditor
Imports ActiveQueryBuilder.View.WPF.QueryView
Imports FullFeaturedMdiDemo.Common
Imports FullFeaturedMdiDemo.MdiControl
Imports GeneralAssembly
Imports GeneralAssembly.Windows.SaveWindows

Public Class ChildWindow
    Inherits MdiChildWindow

    Public Custom Event SaveQueryEvent As EventHandler
        AddHandler(value As EventHandler)
            AddHandler ContentControl.SaveQueryEvent, value
        End AddHandler
        RemoveHandler(value As EventHandler)
            RemoveHandler ContentControl.SaveQueryEvent, value
        End RemoveHandler
        RaiseEvent(sender As System.Object, e As System.EventArgs)
        End RaiseEvent
    End Event
    Public Custom Event SaveAsInFileEvent As EventHandler
        AddHandler(value As EventHandler)
            AddHandler ContentControl.SaveAsInFileEvent, value
        End AddHandler
        RemoveHandler(value As EventHandler)
            RemoveHandler ContentControl.SaveAsInFileEvent, value
        End RemoveHandler
        RaiseEvent(sender As System.Object, e As System.EventArgs)
        End RaiseEvent
    End Event
    Public Custom Event SaveAsNewUserQueryEvent As EventHandler
        AddHandler(value As EventHandler)
            AddHandler ContentControl.SaveAsNewUserQueryEvent, value
        End AddHandler
        RemoveHandler(value As EventHandler)
            RemoveHandler ContentControl.SaveAsNewUserQueryEvent, value
        End RemoveHandler
        RaiseEvent(sender As System.Object, e As System.EventArgs)
        End RaiseEvent
    End Event

    Public ReadOnly Property QueryView() As IQueryView
        Get
            Return ContentControl.QueryView
        End Get
    End Property

    Public ReadOnly Property SqlQuery() As SQLQuery
        Get
            Return ContentControl.SqlQuery
        End Get
    End Property

    Public Property MetadataLoadingOptions() As MetadataLoadingOptions
        Get
            Return ContentControl.SqlContext.LoadingOptions
        End Get
        Set(value As MetadataLoadingOptions)
            ContentControl.SqlContext.LoadingOptions.Assign(value)
        End Set
    End Property

    Public Property MetadataStructureOptions() As MetadataStructureOptions
        Get
            Return ContentControl.SqlContext.MetadataStructureOptions
        End Get
        Set(value As MetadataStructureOptions)
            ContentControl.SqlContext.MetadataStructureOptions.Assign(value)
        End Set
    End Property

    Public Property SqlFormattingOptions() As SQLFormattingOptions
        Set(value As SQLFormattingOptions)
            If _sqlFormattingOptions IsNot Nothing Then
                RemoveHandler _sqlFormattingOptions.Updated, AddressOf _sqlFormattingOptions_Updated
            End If

            If _sqlFormattingOptions Is Nothing Then
                _sqlFormattingOptions = value
            Else
                _sqlFormattingOptions.Assign(value)
            End If

            If _sqlFormattingOptions Is Nothing Then
                Return
            End If
            AddHandler _sqlFormattingOptions.Updated, AddressOf _sqlFormattingOptions_Updated
            ContentControl.CBuilder.QueryTransformer.SQLGenerationOptions = _sqlFormattingOptions
        End Set
        Get
            Return _sqlFormattingOptions
        End Get
    End Property

    Private Sub _sqlFormattingOptions_Updated(sender As Object, e As EventArgs)
        ContentControl.BoxSql.Text = FormattedQueryText
    End Sub

    Public Property SqlGenerationOptions() As SQLGenerationOptions
        Get
            Return QueryView.SQLGenerationOptions
        End Get
        Set(value As SQLGenerationOptions)
            QueryView.SQLGenerationOptions.Assign(value)
        End Set
    End Property

    Public Property BehaviorOptions() As BehaviorOptions
        Get
            Return SqlQuery.BehaviorOptions
        End Get
        Set(value As BehaviorOptions)
            SqlQuery.BehaviorOptions.Assign(value)
        End Set
    End Property

    Public Property ExpressionEditorOptions() As ExpressionEditorOptions
        Get
            Return ContentControl.ExpressionEditorOptions
        End Get
        Set(value As ExpressionEditorOptions)
            ContentControl.ExpressionEditorOptions.Assign(value)
        End Set
    End Property

    Public Property TextEditorOptions() As TextEditorOptions
        Get
            Return ContentControl.BoxSql.Options
        End Get
        Set(value As TextEditorOptions)
            ContentControl.TextEditorOptions.Assign(CType(value, ITextEditorOptions))
            ContentControl.BoxSql.Options.Assign(CType(value, ITextEditorOptions))
            ContentControl.BoxSqlCurrentSubQuery.Options.Assign(CType(value, ITextEditorOptions))
        End Set
    End Property

    Public Property TextEditorSqlOptions() As SqlTextEditorOptions
        Get
            Return ContentControl.BoxSql.SqlOptions
        End Get
        Set(value As SqlTextEditorOptions)
            ContentControl.TextEditorSqlOptions.Assign(CType(value, ISqlTextEditorOptions))
            ContentControl.BoxSql.SqlOptions.Assign(CType(value, ISqlTextEditorOptions))
            ContentControl.BoxSqlCurrentSubQuery.SqlOptions.Assign(CType(value, ISqlTextEditorOptions))
        End Set
    End Property

    Public Property DataSourceOptions() As DataSourceOptions
        Get
            Return ContentControl.DataSourceOptions
        End Get
        Set(value As DataSourceOptions)
            ContentControl.DataSourceOptions.Assign(value)
        End Set
    End Property

    Public Property DesignPaneOptions() As DesignPaneOptions
        Get
            Return ContentControl.DesignPaneOptions
        End Get
        Set(value As DesignPaneOptions)
            ContentControl.DesignPaneOptions.Assign(value)
        End Set
    End Property

    Public Property QueryNavBarOptions() As QueryNavBarOptions
        Get
            Return ContentControl.QueryNavBarOptions
        End Get
        Set(value As QueryNavBarOptions)
            ContentControl.QueryNavBarOptions.Assign(value)
        End Set
    End Property

    Public Property AddObjectDialogOptions() As AddObjectDialogOptions
        Get
            Return ContentControl.AddObjectDialogOptions
        End Get
        Set(value As AddObjectDialogOptions)
            ContentControl.AddObjectDialogOptions.Assign(value)
        End Set
    End Property

    Public Property UserInterfaceOptions() As UserInterfaceOptions
        Get
            Return ContentControl.QView.UserInterfaceOptions
        End Get
        Set(value As UserInterfaceOptions)
            ContentControl.QView.UserInterfaceOptions.Assign(value)
        End Set
    End Property

    Public Property VisualOptions() As VisualOptions
        Get
            Return ContentControl.DockManager.Options
        End Get
        Set(value As VisualOptions)
            ContentControl.DockManager.Options.Assign(value)
        End Set
    End Property

    Public Property QueryColumnListOptions() As QueryColumnListOptions
        Get
            Return ContentControl.QueryColumnListOptions
        End Get
        Set(value As QueryColumnListOptions)
            ContentControl.QueryColumnListOptions.Assign(value)
        End Set
    End Property

    Public Property FileSourceUrl() As String
        Get
            Return ContentControl.FileSourceUrl
        End Get
        Set(value As String)
            ContentControl.FileSourceUrl = value
        End Set
    End Property

    Public Property QueryText() As String
        Get
            Return ContentControl.QueryText
        End Get
        Set(value As String)
            ContentControl.QueryText = value
        End Set
    End Property

    Public Property SqlSourceType() As SourceType
        Get
            Return ContentControl.SqlSourceType
        End Get
        Set(value As SourceType)
            ContentControl.SqlSourceType = value
        End Set
    End Property

    Public Property UserMetadataStructureItem() As MetadataStructureItem
        Get
            Return ContentControl.UserMetadataStructureItem
        End Get
        Set(value As MetadataStructureItem)
            ContentControl.UserMetadataStructureItem = value
        End Set
    End Property

    Public Property IsNeedClose() As Boolean
        Get
            Return ContentControl.IsNeedClose
        End Get
        Set(value As Boolean)
            ContentControl.IsNeedClose = value
        End Set
    End Property

    Public ReadOnly Property FormattedQueryText() As String
        Get
            Return ContentControl.FormattedQueryText
        End Get
    End Property

    Public Property IsModified() As Boolean
        Get
            Return ContentControl.IsModified
        End Get
        Set(value As Boolean)
            ContentControl.IsModified = value
        End Set
    End Property

    Private privateContentControl As ContentWindowChild
    Public Property ContentControl() As ContentWindowChild
        Get
            Return privateContentControl
        End Get
        Private Set(value As ContentWindowChild)
            privateContentControl = value
        End Set
    End Property

    Private ReadOnly _databaseSchemaView As DatabaseSchemaView
    Private _sqlFormattingOptions As SQLFormattingOptions

    Public Sub New(sqlContext As SQLContext, databaseSchemaView As DatabaseSchemaView)
        ContentControl = New ContentWindowChild(sqlContext)
        _sqlFormattingOptions = New SQLFormattingOptions With {.ExpandVirtualObjects = False}
        _databaseSchemaView = databaseSchemaView

        Children.Add(ContentControl)


        AddHandler Loaded, Sub()
                               If Double.IsNaN(Width) Then
                                   Width = ActualWidth
                               End If
                               If Double.IsNaN(Height) Then
                                   Height = ActualHeight
                               End If
                           End Sub
    End Sub

    Public Function CanRedo() As Boolean
        Return ContentControl.CanRedo()
    End Function

    Public Function CanCopy() As Boolean
        Return ContentControl.CanCopy()
    End Function

    Public Function CanPaste() As Boolean
        Return ContentControl.CanPaste()
    End Function

    Public Function CanCut() As Boolean
        Return ContentControl.CanCut()
    End Function

    Public Function CanSelectAll() As Boolean
        Return ContentControl.CanSelectAll()
    End Function

    Public Function CanAddDerivedTable() As Boolean
        Return ContentControl.CanAddDerivedTable()
    End Function

    Public Function CanCopyUnionSubQuery() As Boolean
        Return ContentControl.CanCopyUnionSubQuery()
    End Function

    Public Function CanAddUnionSubQuery() As Boolean
        Return ContentControl.CanAddUnionSubQuery()
    End Function

    Public Function CanShowProperties() As Boolean
        Return ContentControl.CanShowProperties()
    End Function

    Public Function CanAddObject() As Boolean
        Return ContentControl.CanAddObject()
    End Function

    Public Sub AddDerivedTable()
        ContentControl.AddDerivedTable()
    End Sub

    Public Sub CopyUnionSubQuery()
        ContentControl.CopyUnionSubQuery()
    End Sub

    Public Sub AddUnionSubQuery()
        ContentControl.AddUnionSubQuery()
    End Sub

    Public Sub PropertiesQuery()
        ContentControl.PropertiesQuery()
    End Sub

    Public Sub AddObject()
        ContentControl.AddObject()
    End Sub

    Public Sub ShowQueryStatistics()
        ContentControl.ShowQueryStatistics()
    End Sub

    Public Sub Undo()
        ContentControl.Undo()
    End Sub

    Public Sub Redo()
        ContentControl.Redo()
    End Sub

    Public Sub Copy()
        ContentControl.Copy()
    End Sub

    Public Sub Paste()
        ContentControl.Paste()
    End Sub

    Public Sub Cut()
        ContentControl.Cut()
    End Sub

    Public Sub SelectAll()
        ContentControl.SelectAll()
    End Sub

    Public Sub OpenExecuteTab()
        ContentControl.OpenExecuteTab()
    End Sub

    Public Sub ForceClose()
        MyBase.Close()
    End Sub

    Public Overrides Sub Close()
        If IsNeedClose OrElse Not IsModified Then
            MyBase.Close()
            Return
        End If

        Dim point = PointToScreen(New System.Windows.Point(0, 0))


        If SqlSourceType = SourceType.New Then
            IsNeedClose = True

            Dim dialog = New SaveAsWindowDialog(Title) With {
                .Left = point.X,
                .Top = point.Y
            }

            AddHandler dialog.Loaded, AddressOf Dialog_Loaded
            dialog.ShowDialog()

            Select Case dialog.Action
                Case SaveAsWindowDialog.ActionSave.UserQuery
                    ContentControl.OnSaveAsNewUserQueryEvent()
                Case SaveAsWindowDialog.ActionSave.File
                    ContentControl.OnSaveAsInFileEvent()
                Case SaveAsWindowDialog.ActionSave.NotSave
                    MyBase.Close()
                Case SaveAsWindowDialog.ActionSave.Continue
                    IsNeedClose = False
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select
        Else
            Dim saveDialog = New SaveExistQueryDialog With {
                .Left = point.X,
                .Top = point.Y
            }

            AddHandler saveDialog.Loaded, AddressOf Dialog_Loaded
            saveDialog.ShowDialog()

            If Not saveDialog.Result.HasValue Then
                Return
            End If

            IsNeedClose = True

            If saveDialog.Result = True Then
                ContentControl.OnSaveQueryEvent()
                Return
            End If
        End If

        MyBase.Close()
    End Sub

    Private Sub Dialog_Loaded(sender As Object, e As RoutedEventArgs)
        Dim dialog = TryCast(sender, Window)

        If dialog Is Nothing Then
            Return
        End If

        RemoveHandler dialog.Loaded, AddressOf Dialog_Loaded

        dialog.Left += (ActualWidth / 2 - dialog.ActualWidth / 2)
        dialog.Top += (ActualHeight / 2 - dialog.ActualHeight / 2)
    End Sub

    Public Function CanUndo() As Boolean
        Return ContentControl.CanUndo()
    End Function

    Public Sub SetOptions(options As Options)

        AddObjectDialogOptions.Assign(options.AddObjectDialogOptions)
        BehaviorOptions.Assign(options.BehaviorOptions)

        _databaseSchemaView.Options.Assign(options.DatabaseSchemaViewOptions)

        DataSourceOptions.Assign(options.DataSourceOptions)
        DesignPaneOptions.Assign(options.DesignPaneOptions)
        ExpressionEditorOptions.Assign(options.ExpressionEditorOptions)
        QueryColumnListOptions.Assign(options.QueryColumnListOptions)
        QueryNavBarOptions.Assign(options.QueryNavBarOptions)
        SqlFormattingOptions.Assign(options.SqlFormattingOptions)
        SqlGenerationOptions.Assign(options.SqlGenerationOptions)
        TextEditorOptions.Assign(CType(options.TextEditorOptions, ITextEditorOptions))
        TextEditorSqlOptions.Assign(CType(options.TextEditorSqlOptions, ISqlTextEditorOptions))
        UserInterfaceOptions.Assign(options.UserInterfaceOptions)
        VisualOptions.Assign(options.VisualOptions)

    End Sub

    Public Function GetOptions() As Options
        Return New Options With {
            .AddObjectDialogOptions = AddObjectDialogOptions,
            .BehaviorOptions = BehaviorOptions,
            .DatabaseSchemaViewOptions = _databaseSchemaView.Options,
            .DataSourceOptions = DataSourceOptions,
            .DesignPaneOptions = DesignPaneOptions,
            .ExpressionEditorOptions = ExpressionEditorOptions,
            .QueryColumnListOptions = QueryColumnListOptions,
            .QueryNavBarOptions = QueryNavBarOptions,
            .SqlFormattingOptions = SqlFormattingOptions,
            .SqlGenerationOptions = SqlGenerationOptions,
            .TextEditorOptions = TextEditorOptions,
            .TextEditorSqlOptions = TextEditorSqlOptions,
            .UserInterfaceOptions = UserInterfaceOptions,
            .VisualOptions = VisualOptions
        }
    End Function
End Class
