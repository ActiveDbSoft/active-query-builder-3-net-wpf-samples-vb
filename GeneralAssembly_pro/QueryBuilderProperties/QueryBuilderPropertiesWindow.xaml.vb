''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF

Namespace QueryBuilderProperties
    ''' <summary>
    ''' Interaction logic for QueryBuilderPropertiesWindow.xaml
    ''' </summary>
    Partial Public Class QueryBuilderPropertiesWindow

        Private _currentSelectedLink As TextBlock
        Private ReadOnly _queryBuilder As QueryBuilder
        Private ReadOnly _sqlSyntaxPage As SqlSyntaxPage
        Private ReadOnly _offlineModePage As OfflineModePage
        Private ReadOnly _panesVisibilityPage As PanesVisibilityPage
        Private ReadOnly _databaseSchemaViewPage As DatabaseSchemaViewPage
        Private ReadOnly _miscellaneousPage As MiscellaneousPage
        Private ReadOnly _generalPage As GeneralPage
        Private ReadOnly _mainQueryPage As SqlFormattingPage
        Private ReadOnly _derivedQueriesPage As SqlFormattingPage
        Private ReadOnly _expressionSubqueriesPage As SqlFormattingPage

        <DefaultValue(False), Browsable(False)>
        Public Property Modified() As Boolean
            Get
                Return _sqlSyntaxPage.Modified OrElse _offlineModePage.Modified OrElse _panesVisibilityPage.Modified OrElse _databaseSchemaViewPage.Modified OrElse _miscellaneousPage.Modified OrElse _generalPage.Modified OrElse _mainQueryPage.Modified OrElse _derivedQueriesPage.Modified OrElse _expressionSubqueriesPage.Modified
            End Get
            Set(value As Boolean)
                _sqlSyntaxPage.Modified = value
                _offlineModePage.Modified = value
                _panesVisibilityPage.Modified = value
                _databaseSchemaViewPage.Modified = value
                _miscellaneousPage.Modified = value
                _generalPage.Modified = value
                _mainQueryPage.Modified = value
                _derivedQueriesPage.Modified = value
                _expressionSubqueriesPage.Modified = value
            End Set
        End Property

        Public Sub New()
            InitializeComponent()
        End Sub

        Public Sub New(queryBuilder As QueryBuilder)
            Debug.Assert(queryBuilder IsNot Nothing)

            InitializeComponent()

            _queryBuilder = queryBuilder

            Dim syntaxProvider As BaseSyntaxProvider = If(queryBuilder.SyntaxProvider IsNot Nothing, queryBuilder.SyntaxProvider.Clone(), New GenericSyntaxProvider())

            _sqlSyntaxPage = New SqlSyntaxPage(_queryBuilder, syntaxProvider)
            _offlineModePage = New OfflineModePage(_queryBuilder.SQLContext)

            _panesVisibilityPage = New PanesVisibilityPage(_queryBuilder)
            _databaseSchemaViewPage = New DatabaseSchemaViewPage(_queryBuilder)
            _miscellaneousPage = New MiscellaneousPage(_queryBuilder)

            _generalPage = New GeneralPage(_queryBuilder)
            _mainQueryPage = New SqlFormattingPage(SqlBuilderOptionsPages.MainQuery, _queryBuilder)
            _derivedQueriesPage = New SqlFormattingPage(SqlBuilderOptionsPages.DerivedQueries, _queryBuilder)
            _expressionSubqueriesPage = New SqlFormattingPage(SqlBuilderOptionsPages.ExpressionSubqueries, _queryBuilder)

            ' Activate the first page
            UIElement_OnMouseLeftButtonUp(linkSqlSyntax, Nothing)
        End Sub

        Private Sub UIElement_OnMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            If _currentSelectedLink IsNot Nothing Then
                _currentSelectedLink.Foreground = Brushes.Black
            End If

            If Equals(sender, linkSqlSyntax) Then
                SwitchPage(_sqlSyntaxPage)
            ElseIf Equals(sender, linkOfflineMode) Then
                SwitchPage(_offlineModePage)
            ElseIf Equals(sender, linkPanesVisibility) Then
                SwitchPage(_panesVisibilityPage)
            ElseIf Equals(sender, linkMetadataTree) Then
                SwitchPage(_databaseSchemaViewPage)
            ElseIf Equals(sender, linkMiscellaneous) Then
                SwitchPage(_miscellaneousPage)
            ElseIf Equals(sender, linkGeneral) Then
                SwitchPage(_generalPage)
            ElseIf Equals(sender, linkMainQuery) Then
                SwitchPage(_mainQueryPage)
            ElseIf Equals(sender, linkDerievedQueries) Then
                SwitchPage(_derivedQueriesPage)
            ElseIf Equals(sender, linkExpressionSubqueries) Then
                SwitchPage(_expressionSubqueriesPage)
            End If

            _currentSelectedLink = DirectCast(sender, TextBlock)
            _currentSelectedLink.Foreground = Brushes.Blue
        End Sub

        Private Sub SwitchPage(page As UserControl)
            panelPages.Children.Clear()
            page.Margin = New Thickness(10, 10, 0, 0)
            panelPages.Children.Add(page)

        End Sub

        Public Sub ApplyChanges()
            _queryBuilder.BeginUpdate()

            Try
                _sqlSyntaxPage.ApplyChanges()
                _offlineModePage.ApplyChanges()
                _panesVisibilityPage.ApplyChanges()
                _databaseSchemaViewPage.ApplyChanges()
                _miscellaneousPage.ApplyChanges()
                _generalPage.ApplyChanges()
                _mainQueryPage.ApplyChanges()
                _derivedQueriesPage.ApplyChanges()
                _expressionSubqueriesPage.ApplyChanges()
            Finally
                _queryBuilder.EndUpdate()
            End Try

            If _databaseSchemaViewPage.Modified OrElse _offlineModePage.Modified Then
                _queryBuilder.InitializeDatabaseSchemaTree()
            End If

            Modified = False
        End Sub

        Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
            ApplyChanges()
            Close()
        End Sub

        Private Sub ButtonApply_OnClick(sender As Object, e As RoutedEventArgs)
            ApplyChanges()
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub
    End Class
End Namespace
