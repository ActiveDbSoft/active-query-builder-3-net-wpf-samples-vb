'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF

Namespace QueryBuilderProperties
	Public Enum SqlBuilderOptionsPages
		MainQuery
		DerivedQueries
		ExpressionSubqueries
	End Enum
	''' <summary>
	''' Interaction logic for SqlFormattingPage.xaml
	''' </summary>
	Partial Public Class SqlFormattingPage

		Private ReadOnly _page As SqlBuilderOptionsPages = SqlBuilderOptionsPages.MainQuery
		Private ReadOnly _queryBuilder As QueryBuilder
		Private ReadOnly _format As SQLBuilderSelectFormat
		Public Property Modified() As Boolean

		Public Sub New(page As SqlBuilderOptionsPages, queryBuilder As QueryBuilder)
			Modified = False
			_page = page
			_queryBuilder = queryBuilder
			_format = New SQLBuilderSelectFormat(Nothing)

			If _page = SqlBuilderOptionsPages.MainQuery Then
				_format.Assign(_queryBuilder.SQLFormattingOptions.MainQueryFormat)
			ElseIf _page = SqlBuilderOptionsPages.DerivedQueries Then
				_format.Assign(_queryBuilder.SQLFormattingOptions.DerivedQueryFormat)
			ElseIf _page = SqlBuilderOptionsPages.ExpressionSubqueries Then
				_format.Assign(_queryBuilder.SQLFormattingOptions.ExpressionSubQueryFormat)
			End If

			InitializeComponent()

			cbPartsOnNewLines.IsChecked = _format.MainPartsFromNewLine
			cbNewLineAfterKeywords.IsChecked = _format.NewLineAfterPartKeywords
			updownGlobalIndent.Value = _format.IndentGlobal
			updownPartIndent.Value = _format.IndentInPart

			cbNewLineAfterSelectItem.IsChecked = _format.SelectListFormat.NewLineAfterItem

			cbNewLineAfterDatasource.IsChecked = _format.FromClauseFormat.NewLineAfterDatasource
			cbNewLineAfterJoin.IsChecked = _format.FromClauseFormat.NewLineAfterJoin

			cbNewLineWhereTop.IsChecked = (_format.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical OrElse _format.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostOr OrElse _format.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical)
			checkNewLineWhereTop_CheckedChanged(Nothing, New EventArgs())
			cbNewLineWhereRest.IsChecked = (_format.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical)
			checkNewLineWhereRest_CheckedChanged(Nothing, New EventArgs())
			updownWhereIndent.Value = _format.WhereFormat.IndentNestedConditions

			cbNewLineAfterGroupItem.IsChecked = _format.GroupByFormat.NewLineAfterItem

			cbNewLineHavingTop.IsChecked = (_format.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical OrElse _format.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostOr OrElse _format.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical)
			checkNewLineHavingTop_CheckedChanged(Nothing, New EventArgs())
			cbNewLineHavingRest.IsChecked = (_format.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical)
			checkNewLineHavingRest_CheckedChanged(Nothing, New EventArgs())
			updownHavingIndent.Value = _format.HavingFormat.IndentNestedConditions

			AddHandler updownHavingIndent.ValueChanged, AddressOf Changed
			AddHandler updownHavingIndent.ValueChanged, AddressOf Changed
			AddHandler cbNewLineHavingRest.Checked, AddressOf checkNewLineHavingRest_CheckedChanged
			AddHandler cbNewLineHavingRest.Unchecked, AddressOf checkNewLineHavingRest_CheckedChanged
			AddHandler cbNewLineHavingTop.Checked, AddressOf checkNewLineHavingTop_CheckedChanged
			AddHandler cbNewLineHavingTop.Unchecked, AddressOf checkNewLineHavingTop_CheckedChanged
			AddHandler cbNewLineAfterGroupItem.Checked, AddressOf Changed
			AddHandler cbNewLineAfterGroupItem.Unchecked, AddressOf Changed
			AddHandler updownWhereIndent.ValueChanged, AddressOf Changed
			'updownWhereIndent.TextChanged += Changed;
			AddHandler cbNewLineWhereRest.Checked, AddressOf checkNewLineWhereRest_CheckedChanged
			AddHandler cbNewLineWhereRest.Unchecked, AddressOf checkNewLineWhereRest_CheckedChanged
			AddHandler cbNewLineWhereTop.Checked, AddressOf checkNewLineWhereTop_CheckedChanged
			AddHandler cbNewLineWhereTop.Unchecked, AddressOf checkNewLineWhereTop_CheckedChanged
			AddHandler cbNewLineAfterJoin.Checked, AddressOf Changed
			AddHandler cbNewLineAfterJoin.Unchecked, AddressOf Changed
			AddHandler cbNewLineAfterDatasource.Checked, AddressOf Changed
			AddHandler cbNewLineAfterDatasource.Unchecked, AddressOf Changed
			AddHandler cbNewLineAfterSelectItem.Checked, AddressOf Changed
			AddHandler cbNewLineAfterSelectItem.Unchecked, AddressOf Changed
			AddHandler updownPartIndent.ValueChanged, AddressOf Changed
			AddHandler updownGlobalIndent.ValueChanged, AddressOf Changed
			AddHandler cbNewLineAfterKeywords.Checked, AddressOf Changed
			AddHandler cbNewLineAfterKeywords.Unchecked, AddressOf Changed
			AddHandler cbPartsOnNewLines.Checked, AddressOf Changed
			AddHandler cbPartsOnNewLines.Unloaded, AddressOf Changed
		End Sub

		Private Sub checkNewLineWhereTop_CheckedChanged(sender As Object, e As EventArgs)
			cbNewLineWhereRest.IsEnabled = cbNewLineWhereTop.IsChecked.HasValue AndAlso cbNewLineWhereTop.IsChecked.Value

			If Not (cbNewLineWhereTop.IsChecked.HasValue AndAlso cbNewLineWhereTop.IsChecked.Value) Then
				cbNewLineWhereRest.IsChecked = False
				checkNewLineWhereRest_CheckedChanged(cbNewLineWhereRest, New EventArgs())
			End If

			If sender IsNot Nothing Then
				Modified = True
			End If
		End Sub

		Private Sub checkNewLineWhereRest_CheckedChanged(sender As Object, e As EventArgs)
			updownWhereIndent.IsEnabled = cbNewLineWhereRest.IsChecked.HasValue AndAlso cbNewLineWhereRest.IsChecked.Value

			If sender IsNot Nothing Then
				Modified = True
			End If
		End Sub

		Private Sub checkNewLineHavingRest_CheckedChanged(sender As Object, e As EventArgs)
			updownHavingIndent.IsEnabled = cbNewLineHavingRest.IsChecked.HasValue AndAlso cbNewLineHavingRest.IsChecked.Value

			If sender IsNot Nothing Then
				Modified = True
			End If
		End Sub

		Private Sub checkNewLineHavingTop_CheckedChanged(sender As Object, e As EventArgs)
			cbNewLineHavingRest.IsEnabled = cbNewLineHavingTop.IsChecked.HasValue AndAlso cbNewLineHavingTop.IsChecked.Value

			If Not (cbNewLineHavingTop.IsChecked.HasValue AndAlso cbNewLineHavingTop.IsChecked.Value) Then
				cbNewLineHavingRest.IsChecked = False
				checkNewLineHavingRest_CheckedChanged(cbNewLineHavingRest, New EventArgs())
			End If

			If sender IsNot Nothing Then
				Modified = True
			End If
		End Sub

		Private Sub Changed(sender As Object, e As EventArgs)
			Modified = True
		End Sub

		Public Sub ApplyChanges()
			If Not Modified Then
				Return
			End If

			_format.MainPartsFromNewLine = cbPartsOnNewLines.IsChecked.HasValue AndAlso cbPartsOnNewLines.IsChecked.Value
			_format.NewLineAfterPartKeywords = cbNewLineAfterKeywords.IsChecked.HasValue AndAlso cbNewLineAfterKeywords.IsChecked.Value
			_format.IndentInPart = updownPartIndent.Value
			_format.IndentGlobal = updownGlobalIndent.Value

			_format.SelectListFormat.NewLineAfterItem = cbNewLineAfterSelectItem.IsChecked.HasValue AndAlso cbNewLineAfterSelectItem.IsChecked.Value

			_format.FromClauseFormat.NewLineAfterDatasource = cbNewLineAfterDatasource.IsChecked.HasValue AndAlso cbNewLineAfterDatasource.IsChecked.Value
			_format.FromClauseFormat.NewLineAfterJoin = cbNewLineAfterJoin.IsChecked.HasValue AndAlso cbNewLineAfterJoin.IsChecked.Value

			If cbNewLineWhereRest.IsChecked.HasValue AndAlso cbNewLineWhereRest.IsChecked.Value Then
				_format.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical
			ElseIf cbNewLineWhereTop.IsChecked.HasValue AndAlso cbNewLineWhereTop.IsChecked.Value Then
				_format.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical
			Else
				_format.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None
			End If

			_format.WhereFormat.IndentNestedConditions = updownWhereIndent.Value

			_format.GroupByFormat.NewLineAfterItem = cbNewLineAfterGroupItem.IsChecked.HasValue AndAlso cbNewLineAfterGroupItem.IsChecked.Value

			If cbNewLineHavingRest.IsChecked.HasValue AndAlso cbNewLineHavingRest.IsChecked.Value Then
				_format.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical
			ElseIf cbNewLineHavingTop.IsChecked.HasValue AndAlso cbNewLineHavingTop.IsChecked.Value Then
				_format.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical
			Else
				_format.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None
			End If

			_format.HavingFormat.IndentNestedConditions = updownHavingIndent.Value


			If _page = SqlBuilderOptionsPages.MainQuery Then
				_queryBuilder.SQLFormattingOptions.MainQueryFormat.Assign(_format)
			ElseIf _page = SqlBuilderOptionsPages.DerivedQueries Then
				_queryBuilder.SQLFormattingOptions.DerivedQueryFormat.Assign(_format)
			ElseIf _page = SqlBuilderOptionsPages.ExpressionSubqueries Then
				_queryBuilder.SQLFormattingOptions.ExpressionSubQueryFormat.Assign(_format)
			End If
		End Sub
	End Class
End Namespace
