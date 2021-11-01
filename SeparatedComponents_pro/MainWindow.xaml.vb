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
Imports System.Text
Imports System.Windows
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF
Imports Microsoft.Win32


''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private ReadOnly _sqlQuery As SQLQuery
    Private ReadOnly _sqlContext As SQLContext

    Private _lastValidSql As String
    Private _errorPosition As Integer = -1

    Public Sub New()
        'sample query
        Dim builder = New StringBuilder()
        builder.AppendLine("Select Categories.CategoryID,")
        builder.AppendLine("Categories.CategoryName,")
        builder.AppendLine("Categories.Description,")
        builder.AppendLine("Categories.Picture,")
        builder.AppendLine("Products.CategoryID As CategoryID1,")
        builder.AppendLine("Products.QuantityPerUnit")
        builder.AppendLine("From Categories")
        builder.AppendLine("Inner Join Products On Categories.CategoryID = Products.CategoryID")

        InitializeComponent()

        _sqlContext = New SQLContext With {
            .SyntaxProvider = New MSSQLSyntaxProvider With {.ServerVersion = MSSQLServerVersion.MSSQL2012}
                        }
        _sqlContext.LoadingOptions.OfflineMode = True
        ' Load demo metadata from XML file
        _sqlContext.MetadataContainer.ImportFromXML("Northwind.xml")

        databaseSchemaView1.SQLContext = _sqlContext
        databaseSchemaView1.QueryView = QueryView1
        databaseSchemaView1.InitializeDatabaseSchemaTree()

        _sqlQuery = New SQLQuery(_sqlContext)
        AddHandler _sqlQuery.SQLUpdated, AddressOf sqlQuery_SQLUpdated
        ' set example query text
        _sqlQuery.SQL = builder.ToString()

        QueryView1.Query = _sqlQuery

        NavBar.Query = _sqlQuery
        NavBar.QueryView = QueryView1

        sqlTextEditor.Query = _sqlQuery
        sqlTextEditor.Text = builder.ToString()
    End Sub

    Private Sub sqlQuery_SQLUpdated(sender As Object, e As EventArgs)
        ' Text of SQL query has been updated.
        ' To get the query text, ready for execution on SQL server with real object names just use SQL property.
        sqlTextEditor.Text = FormattedSQLBuilder.GetSQL(_sqlQuery.QueryRoot, New SQLFormattingOptions())
        _lastValidSql = sqlTextEditor.Text
    End Sub

    Private Sub MenuItemLoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim openFileDialog = New OpenFileDialog With {.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"}

        ' Load metadata from XML file
        If Not openFileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        _sqlContext.LoadingOptions.OfflineMode = True
        _sqlContext.MetadataContainer.ImportFromXML(openFileDialog.FileName)
        databaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItemSaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim saveFileDialog = New SaveFileDialog With {
            .FileName = "Metadata.xml",
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        }

        ' Save metadata to XML file
        If Not saveFileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        _sqlContext.MetadataContainer.LoadAll(True)
        _sqlContext.MetadataContainer.ExportToXML(saveFileDialog.FileName)
    End Sub

    Private Sub MenuItemOpenQueryStatistic_OnClick(sender As Object, e As RoutedEventArgs)
        Dim builder = New StringBuilder()

        Dim queryStatistics = _sqlQuery.QueryStatistics

        builder.Append("Used Objects (").Append(queryStatistics.UsedDatabaseObjects.Count).AppendLine("):")
        builder.AppendLine()

        For i = 0 To queryStatistics.UsedDatabaseObjects.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjects(i).ObjectName.QualifiedName)
        Next i

        builder.AppendLine().AppendLine()
        builder.Append("Used Columns (").Append(queryStatistics.UsedDatabaseObjectFields.Count).AppendLine("):")
        builder.AppendLine()

        For i = 0 To queryStatistics.UsedDatabaseObjectFields.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjectFields(i).FullName.QualifiedName)
        Next i

        builder.AppendLine().AppendLine()
        builder.Append("Output Expressions (").Append(queryStatistics.OutputColumns.Count).AppendLine("):")
        builder.AppendLine()

        For i = 0 To queryStatistics.OutputColumns.Count - 1
            builder.AppendLine(queryStatistics.OutputColumns(i).Expression)
        Next i

        MessageBox.Show(builder.ToString())
    End Sub

    Private Sub MenuItemOpenAbout_OnClick(sender As Object, e As RoutedEventArgs)
        QueryBuilder.ShowAboutDialog()
    End Sub

    Private Sub SqlTextEditor_OnLostFocus(sender As Object, e As RoutedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            _sqlQuery.SQL = sqlTextEditor.Text
            ErrorBox.Show(Nothing, sqlTextEditor.SyntaxProvider)
        Catch ex As SQLParsingException
            ' Set caret to error position
            sqlTextEditor.SelectionStart = ex.ErrorPos.pos
            _errorPosition = sqlTextEditor.SelectionStart

            ' Show banner with error text
            ErrorBox.Show(ex.Message, sqlTextEditor.SyntaxProvider)
        End Try
    End Sub

    Private Sub SqlTextEditor_OnTextChanged(sender As Object, e As EventArgs)
        If _sqlContext Is Nothing Then
            Return
        End If
        ErrorBox.Show(Nothing, sqlTextEditor.SyntaxProvider)
    End Sub

    Private Sub ErrorBox_OnGoToErrorPosition(sender As Object, e As EventArgs)
        sqlTextEditor.Focus()

        If _errorPosition = -1 Then
            Return
        End If

        sqlTextEditor.ScrollToPosition(_errorPosition)
        sqlTextEditor.CaretOffset = _errorPosition
    End Sub

    Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
        sqlTextEditor.Text = _lastValidSql
        sqlTextEditor.Focus()
    End Sub
End Class
