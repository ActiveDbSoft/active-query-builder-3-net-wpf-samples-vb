''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.IO
Imports System.Windows
Imports System.Windows.Input
Imports ActiveQueryBuilder.Core
Imports Microsoft.Win32
Imports NPOI.HSSF.Record


Partial Public Class MainWindow
    Private _isCanUpdateSqlText As Boolean = True
    Private _defaultSql As String = "-- Any text for comment" & vbLf &
                                    "Select Categories.CategoryName," & vbLf &
                                    "  Products.UnitPrice" & vbLf & "From Categories" & vbLf &
                                    "  Inner Join Products On Categories.CategoryID = Products.CategoryID"

    Public Sub New()
        InitializeComponent()
        InitializeQueryBuilder()
        AddHandler Loaded, AddressOf OnLoaded
    End Sub

    Private Sub OnLoaded(sender As Object, e As RoutedEventArgs)
        TextBoxSql.Text = _defaultSql
    End Sub

    Private Sub InitializeQueryBuilder()
        QBuilder.MetadataContainer.LoadingOptions.OfflineMode = True
        QBuilder.SyntaxProvider = New MSSQLSyntaxProvider()
        QBuilder.MetadataContainer.ImportFromXML("Northwind.xml")
        QBuilder.InitializeDatabaseSchemaTree()
        QBuilder.SQL = _defaultSql
    End Sub

    Private Sub TextBoxSql_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        AssignSqlText(TextBoxSql.Text)
    End Sub

    Private Sub AssignSqlText(text As String)
        _isCanUpdateSqlText = False
        SqlErrorBox.Show("", QBuilder.SyntaxProvider)

        Try
            QBuilder.SQL = text
        Catch ex As Exception
            SqlErrorBox.Show(ex.Message, QBuilder.SyntaxProvider)
        End Try
        _isCanUpdateSqlText = True
    End Sub

    Private Sub QBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
        If Not _isCanUpdateSqlText Then Return
        TextBoxSql.Text = QBuilder.FormattedSQL
    End Sub

    Private Sub ButtonOpenSql_OnClick(sender As Object, e As RoutedEventArgs)
        Dim openDialog = New OpenFileDialog()

        If openDialog.ShowDialog(Me) = True Then
            If File.Exists(openDialog.FileName) Then
                Dim sqlText = String.Empty
                Using reader = New StreamReader(openDialog.FileName)
                    sqlText = reader.ReadToEnd()
                End Using

                TextBoxSql.Text = sqlText
                AssignSqlText(sqlText)
            End If
        End If
    End Sub
End Class
