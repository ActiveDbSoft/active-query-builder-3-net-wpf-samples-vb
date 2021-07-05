//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System
Imports System.Net.Mime
Imports System.Windows
Imports System.Windows.Controls
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF

    Public Partial Class MainWindow
        Private _lastValidSql1 As String = String.Empty
        Private _lastValidSql2 As String = String.Empty
        Private _errorPosition1 As Integer = -1
        Private _errorPosition2 As Integer = -1

        Public Sub New()
            InitializeComponent()
            AddHandler Loaded, AddressOf MainWindow_Loaded
        End Sub

        Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            RemoveHandler Loaded, AddressOf MainWindow_Loaded
            QueryBuilder1.SyntaxProvider = New DB2SyntaxProvider()

            Try
                QueryBuilder1.MetadataLoadingOptions.OfflineMode = True
                QueryBuilder1.MetadataContainer.ImportFromXML("db2_sample_with_alt_names.xml")
                QueryBuilder1.InitializeDatabaseSchemaTree()
                QueryBuilder1.SQL = "Select DEPARTMENT.DEPTNO, DEPARTMENT.DEPTNAME, DEPARTMENT.MGRNO From DEPARTMENT"
            Catch ex As QueryBuilderException
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Private Sub QueryBuilder1_OnSQLUpdated(ByVal sender As Object, ByVal e As EventArgs)
            TextBox1.Text = QueryBuilder1.FormattedSQL
            TextBox2.Text = QueryBuilder1.SQL
        End Sub

        Private Sub TextBox1_OnLostKeyboardFocus(ByVal sender As Object, ByVal routedEventArgs As RoutedEventArgs)
            Try
                QueryBuilder1.SQL = TextBox1.Text
                ErrorBox1.Show(Nothing, QueryBuilder1.SyntaxProvider)
                _lastValidSql1 = TextBox1.Text
                _errorPosition1 = -1
            Catch ex As SQLParsingException
                TextBox1.SelectionStart = ex.ErrorPos.pos
                ErrorBox1.Show(ex.Message, QueryBuilder1.SyntaxProvider)
                _errorPosition1 = ex.ErrorPos.pos
            End Try
        End Sub

        Private Sub TextBox2_OnLostKeyboardFocus(ByVal sender As Object, ByVal routedEventArgs As RoutedEventArgs)
            Try
                QueryBuilder1.SQL = TextBox2.Text
                ErrorBox2.Show(Nothing, QueryBuilder1.SyntaxProvider)
                _lastValidSql2 = TextBox2.Text
                _errorPosition2 = -1
            Catch ex As SQLParsingException
                TextBox2.SelectionStart = ex.ErrorPos.pos
                ErrorBox2.Show(ex.Message, QueryBuilder1.SyntaxProvider)
                _errorPosition2 = ex.ErrorPos.pos
            End Try
        End Sub

        Private Sub MenuItemAbout_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
            QueryBuilder.ShowAboutDialog()
        End Sub

        Private Sub Selector_OnSelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
            ErrorBox1.Show(Nothing, QueryBuilder1.SyntaxProvider)
            ErrorBox2.Show(Nothing, QueryBuilder1.SyntaxProvider)
        End Sub

        Private Sub ErrorBox2_OnSyntaxProviderChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
            Dim oldSql = TextBox2.Text
            Dim caretPosition = TextBox2.CaretIndex
            QueryBuilder1.SyntaxProvider = CType(e.AddedItems(0), BaseSyntaxProvider)
            TextBox2.Text = oldSql
            TextBox2.Focus()
            TextBox2.CaretIndex = caretPosition
        End Sub

        Private Sub ErrorBox2_OnRevertValidTextEvent(ByVal sender As Object, ByVal e As EventArgs)
            TextBox2.Text = _lastValidSql2
            ErrorBox2.Show(Nothing, QueryBuilder1.SyntaxProvider)
            TextBox2.Focus()
        End Sub

        Private Sub ErrorBox2_OnGoToErrorPositionEvent(ByVal sender As Object, ByVal e As EventArgs)
            TextBox2.Focus()
            If _errorPosition2 = -1 Then Return
            If TextBox2.LineCount <> 1 Then TextBox2.ScrollToLine(TextBox1.GetLineIndexFromCharacterIndex(_errorPosition2))
            TextBox2.CaretIndex = _errorPosition2
        End Sub

        Private Sub ErrorBox1_OnSyntaxProviderChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
            Dim oldSql = TextBox1.Text
            Dim caretPosition = TextBox1.CaretIndex
            QueryBuilder1.SyntaxProvider = CType(e.AddedItems(0), BaseSyntaxProvider)
            TextBox1.Text = oldSql
            TextBox1.Focus()
            TextBox1.CaretIndex = caretPosition
        End Sub

        Private Sub ErrorBox1_OnRevertValidTextEvent(ByVal sender As Object, ByVal e As EventArgs)
            TextBox1.Text = _lastValidSql1
            TextBox1.Focus()
        End Sub

        Private Sub ErrorBox1_OnGoToErrorPositionEvent(ByVal sender As Object, ByVal e As EventArgs)
            TextBox1.Focus()

            If _errorPosition1 <> -1 Then
                If TextBox1.LineCount <> 1 Then TextBox1.ScrollToLine(TextBox1.GetLineIndexFromCharacterIndex(_errorPosition1))
                TextBox1.CaretIndex = _errorPosition1
            End If
        End Sub
    End Class
