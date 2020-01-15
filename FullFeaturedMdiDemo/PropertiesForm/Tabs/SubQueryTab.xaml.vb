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
	Public Enum SubQueryType
		Derived
		Cte
		Expression
	End Enum
	''' <summary>
	''' Interaction logic for SubQueryTab.xaml
	''' </summary>
	Public Partial Class SubQueryTab
		Public Property FormattingOptions As SQLFormattingOptions

        Public Property SelectFormat As SQLBuilderSelectFormat

        Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(formattingOptions1 As SQLFormattingOptions, Optional subQueryType2 As SubQueryType = SubQueryType.Expression)
			FormattingOptions = New SQLFormattingOptions()

			InitializeComponent()

			FormattingOptions.Dispose()

			FormattingOptions = formattingOptions1

			Select Case subQueryType2
				Case SubQueryType.Expression
					SelectFormat = formattingOptions1.ExpressionSubQueryFormat
					Exit Select
				Case SubQueryType.Derived
					GroupBoxControl.Header = "Derived tables format options"
					TextBlockCaptionUpDown.Text = "Derived tables indent:"
					CheckBoxStartSubQueriesNewLine.Content = "Start derived tables from new lines"
					TextBlockDescription.Text = "Derived Tables format options" & vbLf & "determine the layout of sub-queries" & vbLf & "used as data sources in the query."

					SelectFormat = formattingOptions1.DerivedQueryFormat
					Exit Select
				Case SubQueryType.Cte
					GroupBoxControl.Header = "Common table expressions format options"
					TextBlockCaptionUpDown.Text = "CTE indent:"
					CheckBoxStartSubQueriesNewLine.Content = "Start CTE from new lines"
					TextBlockDescription.Text = "CTE format options" & vbLf & "determine the layout of sub-queries" & vbLf & "used above the main query in the with clause."

					SelectFormat = formattingOptions1.CTESubQueryFormat
					Exit Select
			End Select

			LoadOptions()
		End Sub

		Public Sub LoadOptions()
			UpDownSubQueryIndent.Value = SelectFormat.IndentInPart
			CheckBoxStartSubQueriesNewLine.IsChecked = SelectFormat.SubQueryTextFromNewLine
		End Sub

		Private Sub UpDownSubQueryIndent_OnValueChanged(sender As Object, e As EventArgs)
			SelectFormat.IndentInPart = UpDownSubQueryIndent.Value
		End Sub

		Private Sub CheckBoxStartSubQueriesNewLine_OnChanged(sender As Object, e As RoutedEventArgs)
			SelectFormat.SubQueryTextFromNewLine = CType((CheckBoxStartSubQueriesNewLine.IsChecked = True), Boolean)
		End Sub

		Private Sub CheckBoxUseFormattingMainQuery_OnChanged(sender As Object, e As RoutedEventArgs)
			SelectFormat.Assign(FormattingOptions.MainQueryFormat)
		End Sub
	End Class
End Namespace
