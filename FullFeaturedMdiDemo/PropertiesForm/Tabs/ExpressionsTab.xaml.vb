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
	''' Interaction logic for ExpressionsTab.xaml
	''' </summary>
	Public Partial Class ExpressionsTab
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
			If SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical AndAlso SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical Then
				CheckBoxStartUpperLevel.IsChecked = True


				CheckBoxStartAllLogicalExprNewLine.IsEnabled = True
			End If
			If SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None AndAlso SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None Then
				CheckBoxStartAllLogicalExprNewLine.IsChecked = False
				CheckBoxStartAllLogicalExprNewLine.IsEnabled = False
			End If

			If SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical AndAlso SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical Then
				CheckBoxStartAllLogicalExprNewLine.IsChecked = True
				UpDownIndentNested.IsEnabled = True
			Else
				If SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical AndAlso SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical Then
					CheckBoxStartUpperLevel.IsChecked = True
				End If
				If SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None AndAlso SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None Then
					UpDownIndentNested.IsEnabled = False
				End If
			End If


			If SelectFormat.WhereFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.TopmostLogical AndAlso SelectFormat.HavingFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.TopmostLogical Then
			End If
			If SelectFormat.WhereFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.AllLogical AndAlso SelectFormat.HavingFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.AllLogical Then
				RadioButtonStartLines.IsChecked = True
				CheckBoxStartUpperLevel.IsChecked = True
				CheckBoxStartAllLogicalExprNewLine.IsChecked = True
			End If

			If SelectFormat.WhereFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.None AndAlso SelectFormat.HavingFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.None Then
				RadioButtonEndLines.IsChecked = True
			End If

			UpDownIndentNested.Value = SelectFormat.WhereFormat.IndentNestedConditions
			UpDownIndentNested.Value = SelectFormat.HavingFormat.IndentNestedConditions

			UpDownIndentNested.Value = SelectFormat.FromClauseFormat.JoinConditionFormat.IndentNestedConditions

			CheckBoxBranchConditionKeywordNewLineWhen.IsChecked = SelectFormat.ConditionalOperatorsFormat.NewLineBeforeWhen

			CheckBoxBranchConditionExpressionNewLine.IsChecked = SelectFormat.ConditionalOperatorsFormat.NewLineAfterWhen

			CheckBoxBranchResultKeywordNewLineThen.IsChecked = SelectFormat.ConditionalOperatorsFormat.NewLineBeforeThen

			CheckBoxBranchResultExpressionNewLine.IsChecked = SelectFormat.ConditionalOperatorsFormat.NewLineAfterThen

			UpDownBranchIndent.Value = SelectFormat.ConditionalOperatorsFormat.IndentBranch

			UpDownExpressionIndent.Value = SelectFormat.ConditionalOperatorsFormat.IndentExpressions
		End Sub

		Private Sub CheckBoxStartUpperLevel_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(FormattingOptions)
				If CheckBoxStartUpperLevel.IsChecked = True Then
					SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical
					SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical

					CheckBoxStartAllLogicalExprNewLine.IsEnabled = True
				Else
					SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None
					SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None

					CheckBoxStartAllLogicalExprNewLine.IsChecked = False
					CheckBoxStartAllLogicalExprNewLine.IsEnabled = False
				End If
			End Using
		End Sub

		Private Sub CheckBoxStartAllLogicalExprNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(FormattingOptions)
				If CheckBoxStartAllLogicalExprNewLine.IsChecked = True Then
					SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical
					SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.AllLogical

					UpDownIndentNested.IsEnabled = True
				Else
					If CheckBoxStartUpperLevel.IsChecked = True Then
						SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical
						SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.TopmostLogical
					Else
						SelectFormat.WhereFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None
						SelectFormat.HavingFormat.NewLineAfter = SQLBuilderConditionFormatNewLine.None
					End If

					UpDownIndentNested.IsEnabled = False
				End If
			End Using
		End Sub

		Private Sub RadioButtonStartLines_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(FormattingOptions)
				If RadioButtonStartLines.IsChecked = True Then
					If CheckBoxStartUpperLevel.IsChecked = True AndAlso CheckBoxStartAllLogicalExprNewLine.IsChecked <> True Then
						SelectFormat.WhereFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.TopmostLogical
						SelectFormat.HavingFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.TopmostLogical
					End If
					If CheckBoxStartUpperLevel.IsChecked = True AndAlso CheckBoxStartAllLogicalExprNewLine.IsChecked = True Then
						SelectFormat.WhereFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.AllLogical
						SelectFormat.HavingFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.AllLogical
					End If
				End If
			End Using
		End Sub

		Private Sub RadioButtonEndLines_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(FormattingOptions)
				If RadioButtonEndLines.IsChecked = True Then
					SelectFormat.WhereFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.None
					SelectFormat.HavingFormat.NewLineBefore = SQLBuilderConditionFormatNewLine.None
				End If
			End Using
		End Sub

		Private Sub UpDownIndentNested_OnValueChanged(sender As Object, e As EventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using New UpdateRegion(FormattingOptions)
				SelectFormat.WhereFormat.IndentNestedConditions = UpDownIndentNested.Value
				SelectFormat.HavingFormat.IndentNestedConditions = UpDownIndentNested.Value

				SelectFormat.FromClauseFormat.JoinConditionFormat.IndentNestedConditions = UpDownIndentNested.Value
			End Using
		End Sub

		Private Sub CheckBoxBranchConditionKeywordNewLineWhen_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			SelectFormat.ConditionalOperatorsFormat.NewLineBeforeWhen = CType((CheckBoxBranchConditionKeywordNewLineWhen.IsChecked = True), Boolean)
		End Sub

		Private Sub CheckBoxBranchConditionExpressionNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			SelectFormat.ConditionalOperatorsFormat.NewLineAfterWhen = CType((CheckBoxBranchConditionExpressionNewLine.IsChecked = True), Boolean)
		End Sub

		Private Sub CheckBoxBranchResultKeywordNewLineThen_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			SelectFormat.ConditionalOperatorsFormat.NewLineBeforeThen = CType((CheckBoxBranchResultKeywordNewLineThen.IsChecked = True), Boolean)
		End Sub

		Private Sub CheckBoxBranchResultExpressionNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			SelectFormat.ConditionalOperatorsFormat.NewLineAfterThen = CType((CheckBoxBranchResultExpressionNewLine.IsChecked = True), Boolean)
		End Sub

		Private Sub UpDownBranchIndent_OnValueChanged(sender As Object, e As EventArgs)
			If Not IsInitialized Then
				Return
			End If

			SelectFormat.ConditionalOperatorsFormat.IndentBranch = CInt(UpDownBranchIndent.Value)
		End Sub

		Private Sub UpDownExpressionIndent_OnValueChanged(sender As Object, e As EventArgs)
			If Not IsInitialized Then
				Return
			End If

			SelectFormat.ConditionalOperatorsFormat.IndentExpressions = CInt(UpDownExpressionIndent.Value)
		End Sub
	End Class
End Namespace
