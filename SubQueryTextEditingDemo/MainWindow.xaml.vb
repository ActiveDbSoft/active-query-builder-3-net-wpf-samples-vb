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


Friend Enum ModeEditor
    Entire
    SubQuery
    Expression
End Enum
''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Public Partial Class MainWindow
    Private _mode As ModeEditor
    Private _lastValidSql as String
    Private _errorPosition as Integer = -1

    Public Sub New()
        InitializeComponent()

        AddHandler Loaded, AddressOf MainWindow_Loaded
        _mode = ModeEditor.Entire
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

        Builder.SyntaxProvider = New MSSQLSyntaxProvider()
        Builder.MetadataProvider = New MSSQLMetadataProvider()

        Builder.MetadataContainer.LoadingOptions.OfflineMode = True
        Builder.MetadataContainer.ImportFromXML("Northwind.xml")

        Builder.InitializeDatabaseSchemaTree()

        TextEditor.QueryProvider = Builder

        Builder.SQL = "Select * From Customers"

        Breadcrumb.QueryView = Builder.QueryView

        AddHandler Builder.ActiveUnionSubQueryChanging, AddressOf Builder_ActiveUnionSubQueryChanging
        AddHandler Builder.ActiveUnionSubQueryChanged, AddressOf Builder_ActiveUnionSubQueryChanged

        AddHandler Breadcrumb.SizeChanged, AddressOf Breadcrumb_SizeChanged
    End Sub

    Private Sub Breadcrumb_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        BottomGrid.InvalidateVisual()
    End Sub

    Private Function CheckSql() As SQLParsingException
        Try
            Dim sql as String = TextEditor.Text.Trim()

            Select Case _mode
                Case ModeEditor.Entire
                    If Not String.IsNullOrEmpty(sql) Then
                        Builder.SQLContext.ParseSelect(sql)
                    End If
                    Exit Select
                Case ModeEditor.SubQuery
                    Builder.SQLContext.ParseSelect(sql)
                    Exit Select
                Case ModeEditor.Expression
                    Builder.SQLContext.ParseSelect(sql)
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select

            Return Nothing
        Catch ex As SQLParsingException
            Return ex
        End Try
    End Function

    Private Sub Builder_ActiveUnionSubQueryChanged(sender As Object, e As EventArgs)
        ApplyText()
    End Sub

    Private Sub ApplyText()
        Dim sqlFormattingOptions As SQLFormattingOptions = Builder.SQLFormattingOptions

        If TextEditor Is Nothing Then
            Return
        End If

        Select Case _mode
            Case ModeEditor.Entire
                TextEditor.Text = Builder.FormattedSQL
                Exit Select
            Case ModeEditor.SubQuery
                If Builder.ActiveUnionSubQuery Is Nothing Then
                    Exit Select
                End If
                Dim subQuery as SubQuery = Builder.ActiveUnionSubQuery.ParentSubQuery
                TextEditor.Text = FormattedSQLBuilder.GetSQL(subQuery, sqlFormattingOptions)
                Exit Select
            Case ModeEditor.Expression
                If Builder.ActiveUnionSubQuery Is Nothing Then
                    Exit Select
                End If
                Dim unionSubQuery as UnionSubQuery = Builder.ActiveUnionSubQuery
                TextEditor.Text = FormattedSQLBuilder.GetSQL(unionSubQuery, sqlFormattingOptions)
                Exit Select
            Case Else
                Throw New ArgumentOutOfRangeException()
        End Select
    End Sub

    Private Sub Builder_ActiveUnionSubQueryChanging(sender As Object, e As ActiveQueryBuilder.View.SubQueryChangingEventArgs)
        Dim exception As SQLParsingException = CheckSql()

        If exception Is Nothing Then
            Return
        End If

        e.Abort = True

        ErrorBox.Show(exception.Message, Builder.SyntaxProvider)
        _errorPosition = exception.ErrorPos.pos
    End Sub

    Private Sub Builder_OnSQLUpdated(sender As Object, e As EventArgs)
        ApplyText()
    End Sub
   
    Private Sub ToggleButton_OnChecked(sender As Object, e As RoutedEventArgs)
        Try
            PopupSwitch.IsOpen = False

            If Equals(sender, RadioButtonEntire) Then
                _mode = ModeEditor.Entire
                Return
            End If

            If Equals(sender, RadioButtonSubQuery) Then
                _mode = ModeEditor.SubQuery
                Return
            End If

            _mode = ModeEditor.Expression
        Finally
            ApplyText()
        End Try
    End Sub

    Private Sub ButtonSwitch_OnClick(sender As Object, e As RoutedEventArgs)
        PopupSwitch.IsOpen = True
    End Sub

    Private Sub TextEditor_OnPreviewLostFocus(sender As Object, e1 As RoutedEventArgs)
        Dim text as String = TextEditor.Text.Trim()

        Builder.BeginUpdate()

        Try
            If Not String.IsNullOrEmpty(text) Then
                Builder.SQLContext.ParseSelect(text)
            End If

            Select Case _mode
                Case ModeEditor.Entire
                    Builder.SQL = text
                    _lastValidSql  = text
                    Exit Select
                Case ModeEditor.SubQuery
                    Dim subQuery As SubQuery = Builder.ActiveUnionSubQuery.ParentSubQuery
                    subQuery.SQL = text
                    _lastValidSql = text
                    Exit Select
                Case ModeEditor.Expression
                    Dim unionSubQuery As UnionSubQuery = Builder.ActiveUnionSubQuery
                    unionSubQuery.SQL = text
                    _lastValidSql = text
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select
        Catch exception As SQLParsingException
          _errorPosition = exception.ErrorPos.pos
            ErrorBox.Show(exception.Message, Builder.SyntaxProvider)
        Finally
            Builder.EndUpdate()
        End Try
    End Sub

    Private Sub ErrorBox_OnGoToErrorPosition(sender As Object, e As EventArgs)
        TextEditor.Focus()

        If _errorPosition <> -1 Then
            TextEditor.CaretOffset = _errorPosition
            TextEditor.ScrollToPosition(_errorPosition)
        End If
    End Sub

    Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
        TextEditor.Text = _lastValidSql
        TextEditor.Focus()
    End Sub
End Class
