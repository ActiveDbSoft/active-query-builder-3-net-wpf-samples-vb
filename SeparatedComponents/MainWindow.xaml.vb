﻿'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Text
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF
Imports Microsoft.Win32

Class MainWindow
 Private ReadOnly _sqlQuery As SQLQuery
    Private ReadOnly _sqlContext As SQLContext

    Public Sub New()
        Dim builder As StringBuilder = New StringBuilder()
        builder.AppendLine("Select Categories.CategoryID,")
        builder.AppendLine("Categories.CategoryName,")
        builder.AppendLine("Categories.Description,")
        builder.AppendLine("Categories.Picture,")
        builder.AppendLine("Products.CategoryID As CategoryID1,")
        builder.AppendLine("Products.QuantityPerUnit")
        builder.AppendLine("From Categories")
        builder.AppendLine("Inner Join Products On Categories.CategoryID = Products.CategoryID")

        InitializeComponent()

        _sqlContext = New SQLContext() With { _
            .SyntaxProvider = New MSSQLSyntaxProvider() With { _
                .ServerVersion = MSSQLServerVersion.MSSQL2012 _
            }
        }

        _sqlContext.LoadingOptions.OfflineMode = True
        _sqlContext.MetadataContainer.ImportFromXML("Northwind.xml")

        databaseSchemaView1.SQLContext = _sqlContext
        databaseSchemaView1.QueryView = QueryView1
        databaseSchemaView1.InitializeDatabaseSchemaTree()

        _sqlQuery = New SQLQuery(_sqlContext)
        AddHandler _sqlQuery.SQLUpdated, AddressOf sqlQuery_SQLUpdated
        _sqlQuery.SQL = builder.ToString()

        QueryView1.Query = _sqlQuery

        NavBar.Query = _sqlQuery
        NavBar.QueryView = QueryView1
    End Sub

    Private Sub sqlQuery_SQLUpdated(sender As Object, e As EventArgs)
        sqlTextEditor.Text = FormattedSQLBuilder.GetSQL(_sqlQuery.QueryRoot, New SQLFormattingOptions())
    End Sub

    Private Sub MenuItemLoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim openFileDialog = New OpenFileDialog() With { _
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*" _
        }

        ' Load metadata from XML file
        If openFileDialog.ShowDialog() <> True Then
            Return
        End If

        _sqlContext.LoadingOptions.OfflineMode = True
        _sqlContext.MetadataContainer.ImportFromXML(openFileDialog.FileName)
        databaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItemSaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim saveFileDialog As SaveFileDialog = New SaveFileDialog() With { _
            .FileName = "Metadata.xml", _
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*" _
        }

        ' Save metadata to XML file
        If saveFileDialog.ShowDialog() <> True Then
            Return
        End If

        _sqlContext.MetadataContainer.LoadAll(True)
        _sqlContext.MetadataContainer.ExportToXML(saveFileDialog.FileName)
    End Sub

    Private Sub MenuItemOpenQueryStatistic_OnClick(sender As Object, e As RoutedEventArgs)
        Dim builder As StringBuilder = New StringBuilder()

        Dim queryStatistics As QueryStatistics = _sqlQuery.QueryStatistics

        builder.Append("Used Objects (").Append(queryStatistics.UsedDatabaseObjects.Count).AppendLine("):")
        builder.AppendLine()

        For i As Int32 = 0 To queryStatistics.UsedDatabaseObjects.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjects(i).ObjectName.QualifiedName)
        Next

        builder.AppendLine().AppendLine()
        builder.Append("Used Columns (").Append(queryStatistics.UsedDatabaseObjectFields.Count).AppendLine("):")
        builder.AppendLine()

        For i As Int32 = 0 To queryStatistics.UsedDatabaseObjectFields.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjectFields(i).ObjectName.QualifiedName)
        Next

        builder.AppendLine().AppendLine()
        builder.Append("Output Expressions (").Append(queryStatistics.OutputColumns.Count).AppendLine("):")
        builder.AppendLine()

        For i As Int32 = 0 To queryStatistics.OutputColumns.Count - 1
            builder.AppendLine(queryStatistics.OutputColumns(i).Expression)
        Next

        MessageBox.Show(builder.ToString())
    End Sub

    Private Sub MenuItemOpenAbout_OnClick(sender As Object, e As RoutedEventArgs)
        QueryBuilder.ShowAboutDialog()
    End Sub

    Private Sub SqlTextEditor_OnLostFocus(sender As Object, e As RoutedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            _sqlQuery.SQL = sqlTextEditor.Text
            ErrorBox.Message = string.Empty
        Catch ex As SQLParsingException
            ' Set caret to error position
            sqlTextEditor.SelectionStart = ex.ErrorPos.pos

            ' Show banner with error text
            ErrorBox.Message = ex.Message
        End Try
    End Sub

    Private Sub TextBox1_OnTextChanged(sender As Object, e As EventArgs)
        If Equals(ErrorBox , Nothing) Then Return
        ErrorBox.Message = string.Empty
    End Sub
End Class
