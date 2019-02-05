'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core

Namespace PropertiesForm
	''' <summary>
	''' Interaction logic for QueryPropertiesWindow.xaml
	''' </summary>
	Public Partial Class QueryPropertiesWindow

		Private _currentSelectedLink As TextBlock
		Private ReadOnly _sqlSyntaxPage As SqlSyntaxPage
		Private ReadOnly _offlineModePage As OfflineModePage
		Private ReadOnly _generalPage As GeneralPage
		Private ReadOnly _mainQueryPage As SqlFormattingPage
		Private ReadOnly _derievedQueriesPage As SqlFormattingPage
		Private ReadOnly _expressionSubqueriesPage As SqlFormattingPage

		<DefaultValue(False)> _
		<Browsable(False)> _
		Public Property Modified() As Boolean
			Get
				Return _sqlSyntaxPage.Modified OrElse _offlineModePage.Modified OrElse _generalPage.Modified OrElse _mainQueryPage.Modified OrElse _derievedQueriesPage.Modified OrElse _expressionSubqueriesPage.Modified
			End Get
			Set
				buttonApply.IsEnabled = value
				_sqlSyntaxPage.Modified = value
				_offlineModePage.Modified = value

				_generalPage.Modified = value
				_mainQueryPage.Modified = value
				_derievedQueriesPage.Modified = value
				_expressionSubqueriesPage.Modified = value
			End Set
		End Property

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(sqlContext As SQLContext, sqlFormattingOptions As SQLFormattingOptions)
			InitializeComponent()

			Dim syntaxProvider As BaseSyntaxProvider = If(sqlContext.SyntaxProvider IsNot Nothing, sqlContext.SyntaxProvider.Clone(), New GenericSyntaxProvider())

			_sqlSyntaxPage = New SqlSyntaxPage(sqlContext, syntaxProvider)
			_offlineModePage = New OfflineModePage(sqlContext)


			_generalPage = New GeneralPage(sqlFormattingOptions)
			_mainQueryPage = New SqlFormattingPage(SqlBuilderOptionsPages.MainQuery, sqlFormattingOptions)
			_derievedQueriesPage = New SqlFormattingPage(SqlBuilderOptionsPages.DerievedQueries, sqlFormattingOptions)
			_expressionSubqueriesPage = New SqlFormattingPage(SqlBuilderOptionsPages.ExpressionSubqueries, sqlFormattingOptions)

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
			ElseIf Equals(sender, linkGeneral) Then
				SwitchPage(_generalPage)
			ElseIf Equals(sender, linkMainQuery) Then
				SwitchPage(_mainQueryPage)
			ElseIf Equals(sender, linkDerievedQueries) Then
				SwitchPage(_derievedQueriesPage)
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
			_sqlSyntaxPage.ApplyChanges()
			_offlineModePage.ApplyChanges()
			_generalPage.ApplyChanges()
			_mainQueryPage.ApplyChanges()
			_derievedQueriesPage.ApplyChanges()
			_expressionSubqueriesPage.ApplyChanges()

			Modified = False
		End Sub

		Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
			ApplyChanges()
			DialogResult = False
		End Sub

		Private Sub ButtonApply_OnClick(sender As Object, e As RoutedEventArgs)
			ApplyChanges()
		End Sub

		Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
			DialogResult = False
		End Sub
	End Class
End Namespace
