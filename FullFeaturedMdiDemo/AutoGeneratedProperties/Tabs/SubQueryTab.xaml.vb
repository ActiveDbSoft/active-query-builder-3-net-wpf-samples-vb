''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Windows
Imports ActiveQueryBuilder.Core

Namespace AutoGeneratedProperties.Tabs
    Public Enum SubQueryType
        Derived
        Cte
        Expression
    End Enum
    ''' <summary>
    ''' Interaction logic for SubQueryTab.xaml
    ''' </summary>
    Partial Public Class SubQueryTab
        Public Property FormattingOptions() As SQLFormattingOptions
        Public Property SelectFormat() As SQLBuilderSelectFormat

        Public Sub New()
            InitializeComponent()
        End Sub

        Public Sub New(formattingOptions As SQLFormattingOptions, Optional subQueryType As SubQueryType = SubQueryType.Expression)
            Me.FormattingOptions = New SQLFormattingOptions()

            InitializeComponent()

            Me.FormattingOptions.Dispose()

            Me.FormattingOptions = formattingOptions

            Select Case subQueryType
                Case SubQueryType.Expression
                    SelectFormat = formattingOptions.ExpressionSubQueryFormat
                Case SubQueryType.Derived
                    GroupBoxControl.Header = "Derived tables format options"
                    TextBlockCaptionUpDown.Text = "Derived tables indent:"
                    CheckBoxStartSubQueriesNewLine.Content = "Start derived tables from new lines"
                    TextBlockDescription.Text = "Derived Tables format options" & ControlChars.Lf & "determine the layout of sub-queries" & ControlChars.Lf & "used as data sources in the query."

                    SelectFormat = formattingOptions.DerivedQueryFormat
                Case SubQueryType.Cte
                    GroupBoxControl.Header = "Common table expressions format options"
                    TextBlockCaptionUpDown.Text = "CTE indent:"
                    CheckBoxStartSubQueriesNewLine.Content = "Start CTE from new lines"
                    TextBlockDescription.Text = "CTE format options" & ControlChars.Lf & "determine the layout of sub-queries" & ControlChars.Lf & "used above the main query in the with clause."

                    SelectFormat = formattingOptions.CTESubQueryFormat
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
            SelectFormat.SubQueryTextFromNewLine = CBool(CheckBoxStartSubQueriesNewLine.IsChecked = True)
        End Sub

        Private Sub CheckBoxUseFormattingMainQuery_OnChanged(sender As Object, e As RoutedEventArgs)
            SelectFormat.Assign(FormattingOptions.MainQueryFormat)
        End Sub
    End Class
End Namespace
