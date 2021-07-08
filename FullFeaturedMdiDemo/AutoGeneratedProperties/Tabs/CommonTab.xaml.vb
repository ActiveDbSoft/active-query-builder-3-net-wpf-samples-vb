''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Windows
Imports ActiveQueryBuilder.Core

Namespace AutoGeneratedProperties.Tabs
	''' <summary>
	''' Interaction logic for CommonTab.xaml
	''' </summary>
	Partial Public Class CommonTab
		Public Property SelectFormat() As SQLBuilderSelectFormat
		Public Property FormattingOptions() As SQLFormattingOptions

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

			Using tempVar As New UpdateRegion(SelectFormat)
				SelectFormat.MainPartsFromNewLine = CBool(CheckBoxStartPartsNewLine.IsChecked = True)
			End Using
		End Sub

		Private Sub CheckBoxInsertNewLineAfterPart_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(SelectFormat)
				SelectFormat.NewLineAfterPartKeywords = CBool(CheckBoxInsertNewLineAfterPart.IsChecked = True)
			End Using
		End Sub

		Private Sub CustomUpDown_OnValueChanged(sender As Object, e As EventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(SelectFormat)
				SelectFormat.IndentInPart = UpDownIndent.Value
			End Using
		End Sub

		Private Sub CheckBoxStartSelectListNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(SelectFormat)
				SelectFormat.SelectListFormat.NewLineAfterItem = CBool(CheckBoxStartSelectListNewLine.IsChecked = True)
			End Using
		End Sub

		Private Sub RadioButtonBeforeComma_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(FormattingOptions)
				SelectFormat.SelectListFormat.NewLineBeforeComma = CBool(RadioButtonBeforeComma.IsChecked = True)
				SelectFormat.OrderByFormat.NewLineBeforeComma = CBool(RadioButtonBeforeComma.IsChecked = True)
				SelectFormat.GroupByFormat.NewLineBeforeComma = CBool(RadioButtonBeforeComma.IsChecked = True)
			End Using
		End Sub

		Private Sub RadioButtonAfterComma_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(FormattingOptions)
				SelectFormat.SelectListFormat.NewLineAfterItem = CBool(RadioButtonAfterComma.IsChecked = True)
				SelectFormat.OrderByFormat.NewLineAfterItem = CBool(RadioButtonAfterComma.IsChecked = True)
				SelectFormat.GroupByFormat.NewLineAfterItem = CBool(RadioButtonAfterComma.IsChecked = True)
			End Using
		End Sub

		Private Sub RadioButtonStartJoinKeywords_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(SelectFormat)
				SelectFormat.FromClauseFormat.NewLineAfterJoin = CBool(RadioButtonStartJoinKeywords.IsChecked = True)
			End Using
		End Sub

		Private Sub RadioButtonStartDataSource_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(SelectFormat)
				SelectFormat.FromClauseFormat.NewLineAfterDatasource = CBool(RadioButtonStartDataSource.IsChecked = True)
			End Using
		End Sub

		Private Sub CheckBoxFromClauseNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Using tempVar As New UpdateRegion(SelectFormat)
				SelectFormat.FromClauseFormat.NewLineBeforeJoinExpression = CBool(CheckBoxFromClauseNewLine.IsChecked = True)
			End Using
		End Sub
	End Class
End Namespace
