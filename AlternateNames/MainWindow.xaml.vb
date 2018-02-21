'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF

Public Class MainWindow

	Public Sub New()
        InitializeComponent()

        AddHandler Loaded, AddressOf MainWindow_Loaded
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

        ' set required syntax provider
        QueryBuilder1.SyntaxProvider = New DB2SyntaxProvider()

        Try
            ' Load demo metadata from XML file
            QueryBuilder1.MetadataLoadingOptions.OfflineMode = True
            QueryBuilder1.MetadataContainer.ImportFromXML("db2_sample_with_alt_names.xml")
            QueryBuilder1.InitializeDatabaseSchemaTree()

            ' set example query text
            QueryBuilder1.SQL = "Select DEPARTMENT.DEPTNO, DEPARTMENT.DEPTNAME, DEPARTMENT.MGRNO From DEPARTMENT"
        Catch ex As QueryBuilderException
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub QueryBuilder1_OnSQLUpdated(sender As Object, e As EventArgs)
        ' Text of SQL query has been updated.

        ' To get the formatted query text with alternate object names just use FormattedSQL property
        TextBox1.Text = QueryBuilder1.FormattedSQL

        ' To get the query text, ready for execution on SQL server with real object names just use SQL property.
        TextBox2.Text = QueryBuilder1.SQL

        ' The query text containing in SQL property is unformatted. If you need the formatted text, but with real object names, 
        ' do the following:
        '			SQLFormattingOptions formattingOptions = new SQLFormattingOptions();
        '			formattingOptions.Assign(queryBuilder1.SQLFormattingOptions); // copy actual formatting options from the QueryBuilder
        '			formattingOptions.UseAltNames = false; // disable alt names
        '			TextBox2.Text = FormattedSQLBuilder.GetSQL(queryBuilder1.SQLQuery.QueryRoot, formattingOptions);
    End Sub

    Private Sub TextBox1_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QueryBuilder1.SQL = TextBox1.Text

            ' Hide error banner if any
            ShowErrorBanner(TextBox1, "")
        Catch ex As SQLParsingException
            ' Set caret to error position
            TextBox1.SelectionStart = ex.ErrorPos.pos

            ' Show banner with error text
            ShowErrorBanner(TextBox1, ex.Message)
        End Try
    End Sub

    Private Sub TextBox2_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QueryBuilder1.SQL = TextBox2.Text

            ' Hide error banner if any
            ShowErrorBanner(TextBox2, "")
        Catch ex As SQLParsingException
            ' Set caret to error position
            TextBox2.SelectionStart = ex.ErrorPos.pos

            ' Show banner with error text
            ShowErrorBanner(TextBox2, ex.Message)
        End Try
    End Sub

    Public Sub ShowErrorBanner(control As FrameworkElement, text As String)
        ' Show new banner if text is not empty
      If Equals(control, ErrorBox1)
            ErrorBox1.Message = text
      End If
        If Equals(control, ErrorBox2)
            ErrorBox2.Message = text
        End If
	End Sub

    Private Sub MenuItemAbout_OnClick(sender As Object, e As RoutedEventArgs)
		QueryBuilder.ShowAboutDialog()
	End Sub

	Private Sub TextBox1_OnTextChanged(sender As Object, e As TextChangedEventArgs)
	    ErrorBox1.Message = String.Empty
	End Sub

	Private Sub TextBox2_OnTextChanged(sender As Object, e As TextChangedEventArgs)
	    ErrorBox2.Message = String.Empty
	End Sub

	Private Sub Selector_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
	    ErrorBox1.Message = String.Empty
	    ErrorBox2.Message = String.Empty
	End Sub
End Class
