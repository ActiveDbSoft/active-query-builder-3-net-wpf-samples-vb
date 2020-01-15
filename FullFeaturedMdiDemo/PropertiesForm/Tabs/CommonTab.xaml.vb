'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows
Imports ActiveQueryBuilder.Core

Namespace PropertiesForm.Tabs
	''' <summary>
	''' Interaction logic for CommonTab.xaml
	''' </summary>
	Public Partial Class CommonTab
		Public Property SelectFormat As SQLBuilderSelectFormat

        Public Property FormattingOptions As SQLFormattingOptions

        Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(sqlFormattingOptions As SQLFormattingOptions, sqlBuilderSelectFormat As SQLBuilderSelectFormat)
			InitializeComponent()

			SelectFormat = sqlBuilderSelectFormat
			FormattingOptions = sqlFormattingOptions

			LoadOptions()
		End Sub

		Public Sub LoadOptions()
			CheckBoxStartPartsNewLine.IsChecked = SelectFormat.MainPartsFromNewLine
			CheckBoxInsertNewLineAfterPart.IsChecked = SelectFormat.NewLineAfterPartKeywords
			UpDownIndent.Value = SelectFormat.IndentInPart
			CheckBoxStartSelectListNewLine.IsChecked = SelectFormat.SelectListFormat.NewLineAfterItem

			RadioButtonBeforeComma.IsChecked = SelectFormat.SelectListFormat.NewLineBeforeComma
			RadioButtonAfterComma.IsChecked = SelectFormat.SelectListFormat.NewLineAfterItem

			RadioButtonStartDataSource.IsChecked = SelectFormat.FromClauseFormat.NewLineAfterDatasource
			RadioButtonStartJoinKeywords.IsChecked = SelectFormat.FromClauseFormat.NewLineAfterJoin
			CheckBoxFromClauseNewLine.IsChecked = SelectFormat.FromClauseFormat.NewLineBeforeJoinExpression
		End Sub

		Private Sub CheckBoxStartPartsNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(SelectFormat)
				SelectFormat.MainPartsFromNewLine = CType((CheckBoxStartPartsNewLine.IsChecked = True), Boolean)
			End Using
		End Sub

		Private Sub CheckBoxInsertNewLineAfterPart_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(SelectFormat)
				SelectFormat.NewLineAfterPartKeywords = CType((CheckBoxInsertNewLineAfterPart.IsChecked = True), Boolean)
			End Using
		End Sub

		Private Sub CustomUpDown_OnValueChanged(sender As Object, e As EventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(SelectFormat)
				SelectFormat.IndentInPart = UpDownIndent.Value
			End Using
		End Sub

		Private Sub CheckBoxStartSelectListNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(SelectFormat)
				SelectFormat.SelectListFormat.NewLineAfterItem = CType((CheckBoxStartSelectListNewLine.IsChecked = True), Boolean)
			End Using
		End Sub

		Private Sub RadioButtonBeforeComma_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(FormattingOptions)
				SelectFormat.SelectListFormat.NewLineBeforeComma = CType((RadioButtonBeforeComma.IsChecked = True), Boolean)
				SelectFormat.OrderByFormat.NewLineBeforeComma = CType((RadioButtonBeforeComma.IsChecked = True), Boolean)
				SelectFormat.GroupByFormat.NewLineBeforeComma = CType((RadioButtonBeforeComma.IsChecked = True), Boolean)
			End Using
		End Sub

		Private Sub RadioButtonAfterComma_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(FormattingOptions)
				SelectFormat.SelectListFormat.NewLineAfterItem = CType((RadioButtonAfterComma.IsChecked = True), Boolean)
				SelectFormat.OrderByFormat.NewLineAfterItem = CType((RadioButtonAfterComma.IsChecked = True), Boolean)
				SelectFormat.GroupByFormat.NewLineAfterItem = CType((RadioButtonAfterComma.IsChecked = True), Boolean)
			End Using
		End Sub

		Private Sub RadioButtonStartJoinKeywords_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(SelectFormat)
				SelectFormat.FromClauseFormat.NewLineAfterJoin = CType((RadioButtonStartJoinKeywords.IsChecked = True), Boolean)
			End Using
		End Sub

		Private Sub RadioButtonStartDataSource_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(SelectFormat)
				SelectFormat.FromClauseFormat.NewLineAfterDatasource = CType((RadioButtonStartDataSource.IsChecked = True), Boolean)
			End Using
		End Sub

		Private Sub CheckBoxFromClauseNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(SelectFormat)
				SelectFormat.FromClauseFormat.NewLineBeforeJoinExpression = CType((CheckBoxFromClauseNewLine.IsChecked = True), Boolean)
			End Using
		End Sub
	End Class
End Namespace
