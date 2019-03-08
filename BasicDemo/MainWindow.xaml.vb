'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports Oracle.ManagedDataAccess.Client
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Markup
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF
Imports ActiveQueryBuilder.View.WPF.ExpressionEditor
Imports BasicDemo.ConnectionWindow
Imports BasicDemo.PropertiesForm
Imports Microsoft.Win32
Imports Helpers = ActiveQueryBuilder.View.Helpers
Imports Timer = System.Timers.Timer

''' <summary>
''' Interaction logic for Window1.xaml
''' </summary>
Partial Public Class MainWindow
    Private ReadOnly _openFileDialog As New OpenFileDialog()
    Private ReadOnly _saveFileDialog As New SaveFileDialog()

    Private ReadOnly _mssqlMetadataProvider1 As MSSQLMetadataProvider
    Private ReadOnly _mssqlSyntaxProvider1 As MSSQLSyntaxProvider
    Private ReadOnly _oledbMetadataProvider1 As OLEDBMetadataProvider
    Private ReadOnly _msaccessSyntaxProvider1 As MSAccessSyntaxProvider
    Private ReadOnly _oracleMetadataProvider1 As OracleNativeMetadataProvider
    Private ReadOnly _oracleSyntaxProvider1 As OracleSyntaxProvider
    Private ReadOnly _genericSyntaxProvider1 As GenericSyntaxProvider
    Private ReadOnly _odbcMetadataProvider1 As ODBCMetadataProvider
    Private ReadOnly _errorPopup As Popup

    Public Sub New()
        InitializeComponent()

        _errorPopup = New Popup() With {
            .Placement = PlacementMode.Relative
        }
        _mssqlMetadataProvider1 = New MSSQLMetadataProvider()
        _mssqlSyntaxProvider1 = New MSSQLSyntaxProvider()
        _oledbMetadataProvider1 = New OLEDBMetadataProvider()
        _msaccessSyntaxProvider1 = New MSAccessSyntaxProvider()
        _oracleMetadataProvider1 = New OracleNativeMetadataProvider()
        _oracleSyntaxProvider1 = New OracleSyntaxProvider()
        _genericSyntaxProvider1 = New GenericSyntaxProvider()
        _odbcMetadataProvider1 = New ODBCMetadataProvider()

        _openFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        _saveFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"

        sqlTextEditor1.QueryProvider = queryBuilder
        sqlTextEditor1.ActiveUnionSubQuery = queryBuilder.ActiveUnionSubQuery
        AddHandler queryBuilder.ActiveUnionSubQueryChanged, AddressOf ActiveUnionSubQueryChanged

        ' DEMO WARNING
        Dim trialNoticePanel As Border = New Border() With {
            .SnapsToDevicePixels = True,
            .BorderBrush = Brushes.Black,
            .BorderThickness = New Thickness(1),
            .Background = Brushes.LightGreen,
            .Padding = New Thickness(5),
            .Margin = New Thickness(0, 0, 0, 2)
        }
        trialNoticePanel.SetValue(Grid.RowProperty, 1)

        Dim label As TextBlock = New TextBlock() With {
            .Text = "Generation of random aliases for the query output columns is the limitation of the trial version. The full version is free from this behavior.",
            .HorizontalAlignment = HorizontalAlignment.Left,
            .VerticalAlignment = VerticalAlignment.Top
        }

        trialNoticePanel.Child = label
        GridRoot.Children.Add(trialNoticePanel)
    End Sub

    Private Sub ActiveUnionSubQueryChanged(sender As Object, e As EventArgs)
        sqlTextEditor1.ActiveUnionSubQuery = queryBuilder.ActiveUnionSubQuery
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        queryBuilder.SyntaxProvider = _genericSyntaxProvider1

        Dim currentLang As String = Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName

        Language = XmlLanguage.GetLanguage(If(Helpers.Localizer.Languages.Contains(currentLang.ToLower()), currentLang, "en"))
    End Sub

    Private Sub menuItemAbout_Click(sender As Object, e As RoutedEventArgs)
        QueryBuilder.ShowAboutDialog()
    End Sub

    ' TextBox lost focus by mouse
    Private Sub SqlTextEditor_OnLostFocus(sender As Object, e As RoutedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            queryBuilder.SQL = sqlTextEditor1.Text
            ErrorBox.Message = string.Empty
        Catch ex As SQLParsingException
            ' Set caret to error position
            sqlTextEditor1.SelectionStart = ex.ErrorPos.pos
            ErrorBox.Message = ex.Message
        End Try
    End Sub

    ' TextBox lost focus by keyboard
    Private Sub SqlTextEditor_LostKeyboardFocus(sender As Object, e As System.Windows.Input.KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            queryBuilder.SQL = sqlTextEditor1.Text
            ShowErrorBanner(DirectCast(sender, FrameworkElement), "")
        Catch ex As SQLParsingException
            ' Set caret to error position
            sqlTextEditor1.SelectionStart = ex.ErrorPos.pos
            ' Report error
            ShowErrorBanner(DirectCast(sender, FrameworkElement), ex.Message)
        End Try
    End Sub

    Private Sub WarnAboutGenericSyntaxProvider()
        If TypeOf queryBuilder.SyntaxProvider Is GenericSyntaxProvider Then
            panel1.Visibility = Visibility.Visible

            ' setup the panel to hide automatically
            Dim timer As Timer = New Timer()
            AddHandler timer.Elapsed, AddressOf TimerEvent
            timer.Interval = 10000
            timer.Start()
        End If
    End Sub

    Private Sub TimerEvent(source As [Object], args As EventArgs)
        Dispatcher.Invoke(DirectCast(Sub() panel1.Visibility = Visibility.Collapsed, Action))

        DirectCast(source, Timer).[Stop]()
        DirectCast(source, Timer).Dispose()
    End Sub

    Private Sub RefreshMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Force the query builder to refresh metadata from current connection
        ' to refresh metadata, just clear MetadataContainer and reinitialize metadata tree

        If queryBuilder.MetadataProvider Is Nothing OrElse Not queryBuilder.MetadataProvider.Connected Then
            Return
        End If

        queryBuilder.ClearMetadata()
        queryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub ClearMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Clear the metadata

        If MessageBox.Show("Clear Metadata Container?", "Confirmation", MessageBoxButton.YesNo) = MessageBoxResult.Yes Then
            queryBuilder.ClearMetadata()
        End If
    End Sub

    Private Sub LoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Load metadata from XML file
        If _openFileDialog.ShowDialog() <> True Then
            Return
        End If

        queryBuilder.MetadataLoadingOptions.OfflineMode = True
        queryBuilder.MetadataContainer.ImportFromXML(_openFileDialog.FileName)
        queryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub SaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' Save metadata to XML file
        _saveFileDialog.FileName = "Metadata.xml"

        If _saveFileDialog.ShowDialog() <> True Then
            Return
        End If

        queryBuilder.MetadataContainer.LoadAll(True)
        queryBuilder.MetadataContainer.ExportToXML(_saveFileDialog.FileName)
    End Sub

    Private Sub QueryBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
        ' Handle the event raised by SQL Builder object that the text of SQL query is changed
        ' update the text box
        sqlTextEditor1.Text = queryBuilder.FormattedSQL
    End Sub


    Public Sub ShowErrorBanner(control As FrameworkElement, text As String)
        ' Show new banner if text is not empty
        If [String].IsNullOrEmpty(text) Then
            If _errorPopup.IsOpen Then
                _errorPopup.IsOpen = False
            End If
            Return
        End If

        Dim label As TextBlock = New TextBlock() With {
            .Text = text,
            .Background = Brushes.LightPink,
            .Padding = New Thickness(5)
        }

        _errorPopup.PlacementTarget = control

        _errorPopup.Child = label
        _errorPopup.IsOpen = True
        _errorPopup.HorizontalOffset = control.ActualWidth - label.ActualWidth - 2
        _errorPopup.VerticalOffset = control.ActualHeight - label.ActualHeight - 2
        Dim timer As System.Threading.Timer = New System.Threading.Timer(AddressOf CallBackPopup, Nothing, 3000, Timeout.Infinite)
    End Sub

    Private Sub CallBackPopup(state As Object)
        Dispatcher.BeginInvoke(DirectCast(Sub() If _errorPopup.IsOpen Then _errorPopup.IsOpen = False, Action))
    End Sub

    Public Sub ResetQueryBuilder()
        queryBuilder.ClearMetadata()
        queryBuilder.MetadataProvider = Nothing
        queryBuilder.SyntaxProvider = Nothing
        queryBuilder.MetadataLoadingOptions.OfflineMode = False
    End Sub

    Private Sub connectToMSSQLServer_OnClick(sender As Object, e As RoutedEventArgs)
        ' Connect to MS SQL Server

        ' show the connection form
        Dim f As MSSQLConnectionWindow = New MSSQLConnectionWindow() With {
            .Owner = Me
        }

        If f.ShowDialog() <> True Then
            Return
        End If

        ResetQueryBuilder()

        ' create new SqlConnection object using the connections string from the connection form
        _mssqlMetadataProvider1.Connection = New SqlConnection(f.ConnectionString)

        ' setup the query builder with metadata and syntax providers
        queryBuilder.MetadataProvider = _mssqlMetadataProvider1
        queryBuilder.SyntaxProvider = _mssqlSyntaxProvider1

        ' kick the query builder to fill metadata tree
        queryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub ConnectToAccess_OnClick(sender As Object, e As RoutedEventArgs)
        ' Connect to MS Access database using OLE DB provider

        ' show the connection form
        Dim f As AccessConnectionWindow = New AccessConnectionWindow() With {
            .Owner = Me
        }

        If f.ShowDialog() <> True Then
            Return
        End If

        ResetQueryBuilder()

        ' create new OleDbConnection object using the connections string from the connection form
        _oledbMetadataProvider1.Connection = New OleDbConnection(f.ConnectionString)

        ' setup the query builder with metadata and syntax providers
        queryBuilder.MetadataProvider = _oledbMetadataProvider1
        queryBuilder.SyntaxProvider = _msaccessSyntaxProvider1

        ' kick the query builder to fill metadata tree
        queryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub ConnectToOracle_OnClick(sender As Object, e As RoutedEventArgs)
        ' Connect to Oracle Server.
        ' Connect using a metadata provider based on the native Oracle Data Provider for .NET (Oracle.DataAccess.Client).

        ' show the connection form
        Dim f As OracleConnectionWindow = New OracleConnectionWindow() With {
            .Owner = Me
        }

        If f.ShowDialog() <> True Then
            Return
        End If

        ResetQueryBuilder()

        ' create new OracleConnection object using the connections string from the connection form
        _oracleMetadataProvider1.Connection = New OracleConnection(f.ConnectionString)

        ' setup the query builder with metadata and syntax providers
        queryBuilder.MetadataProvider = _oracleMetadataProvider1
        queryBuilder.SyntaxProvider = _oracleSyntaxProvider1

        ' kick the query builder to fill metadata tree
        queryBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub ConnectToOLEDB_OnClick(sender As Object, e As RoutedEventArgs)
        ' Connect to a database through the OLE DB provider

        ' show the connection form
        Dim f As OLEDBConnectionWindow = New OLEDBConnectionWindow() With {
            .Owner = Me
        }

        If f.ShowDialog() <> True Then
            Return
        End If

        ResetQueryBuilder()

        ' create new OleDbConnection object using the connections string from the connection form
        _oledbMetadataProvider1.Connection = New OleDbConnection(f.ConnectionString)

        ' setup the query builder with metadata and syntax providers
        queryBuilder.MetadataProvider = _oledbMetadataProvider1
        queryBuilder.SyntaxProvider = _genericSyntaxProvider1

        ' kick the query builder to fill metadata tree
        queryBuilder.InitializeDatabaseSchemaTree()

        WarnAboutGenericSyntaxProvider()
        ' show warning (just for demonstration purposes)
    End Sub

    Private Sub ConnectToODBC_OnClick(sender As Object, e As RoutedEventArgs)
        ' Connect to a database through the ODBC provider

        ' show the connection form
        Dim f As ODBCConnectionWindow = New ODBCConnectionWindow() With {
            .Owner = Me
        }

        If f.ShowDialog() <> True Then
            Return
        End If

        ResetQueryBuilder()

        ' create new OdbcConnection object using the connections string from the connection form
        _odbcMetadataProvider1.Connection = New OdbcConnection(f.ConnectionString)

        ' setup the query builder with metadata and syntax providers
        queryBuilder.MetadataProvider = _odbcMetadataProvider1
        queryBuilder.SyntaxProvider = _genericSyntaxProvider1

        ' kick the query builder to fill metadata tree
        queryBuilder.InitializeDatabaseSchemaTree()

        WarnAboutGenericSyntaxProvider()
        ' show warning (just for demonstration purposes)
    End Sub

    Private Sub FillProgrammatically_OnClick(sender As Object, e As RoutedEventArgs)
        ResetQueryBuilder()

        ' Fill the query builder metadata programmatically

        ' setup the query builder with metadata and syntax providers
        queryBuilder.SyntaxProvider = _genericSyntaxProvider1
        queryBuilder.MetadataLoadingOptions.OfflineMode = True
        ' prevent querying obejects from database
        ' create database and schema
        Dim database As MetadataNamespace = queryBuilder.MetadataContainer.AddDatabase("MyDB")
        database.[Default] = True
        Dim schema As MetadataNamespace = database.AddSchema("MySchema")
        schema.[Default] = True

        ' create table
        Dim tableOrders As MetadataObject = schema.AddTable("Orders")
        tableOrders.AddField("OrderID")
        tableOrders.AddField("OrderDate")
        tableOrders.AddField("CustomerID")
        tableOrders.AddField("ResellerID")

        ' create another table
        Dim tableCustomers As MetadataObject = schema.AddTable("Customers")
        tableCustomers.AddField("CustomerID")
        tableCustomers.AddField("CustomerName")
        tableCustomers.AddField("CustomerAddress")

        ' add a relation between these two tables
        Dim relation As MetadataForeignKey = tableCustomers.AddForeignKey("FK_CustomerID")
        relation.Fields.Add("CustomerID")
        relation.ReferencedObjectName = tableOrders.GetQualifiedName()
        relation.ReferencedFields.Add("CustomerID")

        'create view
        Dim viewResellers As MetadataObject = schema.AddView("Resellers")
        viewResellers.AddField("ResellerID")
        viewResellers.AddField("ResellerName")

        ' kick the query builder to fill metadata tree
        queryBuilder.InitializeDatabaseSchemaTree()

        WarnAboutGenericSyntaxProvider()
        ' show warning (just for demonstration purposes)
    End Sub

    Private Sub QueryStatistic_OnClick(sender As Object, e As RoutedEventArgs)
        Dim queryStatistics As QueryStatistics = queryBuilder.QueryStatistics
        Dim builder As New StringBuilder()

        builder.Append("Used Objects (").Append(queryStatistics.UsedDatabaseObjects.Count).AppendLine("):")
        builder.AppendLine()

        For i As Integer = 0 To queryStatistics.UsedDatabaseObjects.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjects(i).ObjectName.QualifiedName)
        Next

        builder.AppendLine().AppendLine()
        builder.Append("Used Columns (").Append(queryStatistics.UsedDatabaseObjectFields.Count).AppendLine("):")
        builder.AppendLine()

        For i As Integer = 0 To queryStatistics.UsedDatabaseObjectFields.Count - 1
            builder.AppendLine(queryStatistics.UsedDatabaseObjectFields(i).FullName.QualifiedName)
        Next

        builder.AppendLine().AppendLine()
        builder.Append("Output Expressions (").Append(queryStatistics.OutputColumns.Count).AppendLine("):")
        builder.AppendLine()

        For i As Integer = 0 To queryStatistics.OutputColumns.Count - 1
            builder.AppendLine(queryStatistics.OutputColumns(i).Expression)
        Next

        Dim f As QueryStatisticsWindow = New QueryStatisticsWindow() With {
            .Owner = Me
        }
        f.textBox.Text = builder.ToString()
        f.ShowDialog()
    End Sub

    Private Sub Properties_OnClick(sender As Object, e As RoutedEventArgs)
        ' Show Properties form
        Dim f As QueryBuilderPropertiesWindow = New QueryBuilderPropertiesWindow(queryBuilder) With {
            .Owner = Me
        }

        f.ShowDialog()


        WarnAboutGenericSyntaxProvider()
        ' show warning (just for demonstration purposes)
    End Sub

    Private Sub Selector_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If e.AddedItems.Count = 0 Then
            Return
        End If
        Dim tab As TabItem = TryCast(e.AddedItems(0), TabItem)
        If tab Is Nothing Then
            Return
        End If

        If tab.Header.ToString() <> "Data" Then
            Return
        End If

        dataGridView1.ItemsSource = Nothing

        If queryBuilder.MetadataProvider IsNot Nothing AndAlso queryBuilder.MetadataProvider.Connected Then
            If TypeOf queryBuilder.MetadataProvider Is MSSQLMetadataProvider Then
                Dim command As SqlCommand = DirectCast(queryBuilder.MetadataProvider.Connection.CreateCommand(), SqlCommand)
                command.CommandText = queryBuilder.SQL

                ' handle the query parameters
                If queryBuilder.Parameters.Count > 0 Then
                    For i As Integer = 0 To queryBuilder.Parameters.Count - 1
                        If Not command.Parameters.Contains(queryBuilder.Parameters(i).FullName) Then
                            Dim parameter As New SqlParameter()
                            parameter.ParameterName = queryBuilder.Parameters(i).FullName
                            parameter.DbType = queryBuilder.Parameters(i).DataType
                            command.Parameters.Add(parameter)
                        End If
                    Next

                    Dim qpf As New QueryParametersWindow(command)
                    qpf.ShowDialog()
                End If

                Dim adapter As New SqlDataAdapter(command)
                Dim dataset As New DataSet()

                Try
                    adapter.Fill(dataset, "QueryResult")
                    dataGridView1.ItemsSource = dataset.Tables("QueryResult").DefaultView
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "SQL query error")
                End Try
            ElseIf TypeOf queryBuilder.MetadataProvider Is OracleNativeMetadataProvider Then
                Dim command As OracleCommand = DirectCast(queryBuilder.MetadataProvider.Connection.CreateCommand(), OracleCommand)
                command.CommandText = queryBuilder.SQL

                ' handle the query parameters
                If queryBuilder.Parameters.Count > 0 Then
                    For Each t As Parameter In queryBuilder.Parameters
                        If command.Parameters.Contains(t.FullName) Then
                            Continue For
                        End If

                        Dim parameter As OracleParameter = New OracleParameter()
                        parameter.ParameterName = t.FullName
                        parameter.DbType = t.DataType
                        command.Parameters.Add(parameter)
                    Next

                    Dim qpf As New QueryParametersWindow(command)

                    qpf.ShowDialog()
                End If

                Dim adapter As OracleDataAdapter = New OracleDataAdapter(command)
                Dim dataset As DataSet = New DataSet()

                Try
                    adapter.Fill(dataset, "QueryResult")
                    dataGridView1.ItemsSource = dataset.Tables("QueryResult").DefaultView
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "SQL query error")
                End Try
            ElseIf TypeOf queryBuilder.MetadataProvider Is OLEDBMetadataProvider Then
                Dim command As OleDbCommand = DirectCast(queryBuilder.MetadataProvider.Connection.CreateCommand(), OleDbCommand)
                command.CommandText = queryBuilder.SQL

                ' handle the query parameters
                If queryBuilder.Parameters.Count > 0 Then
                    For i As Integer = 0 To queryBuilder.Parameters.Count - 1
                        If Not command.Parameters.Contains(queryBuilder.Parameters(i).FullName) Then
                            Dim parameter As New OleDbParameter()
                            parameter.ParameterName = queryBuilder.Parameters(i).FullName
                            parameter.DbType = queryBuilder.Parameters(i).DataType
                            command.Parameters.Add(parameter)
                        End If
                    Next

                    Dim qpf As New QueryParametersWindow(command)


                    qpf.ShowDialog()
                End If

                Dim adapter As New OleDbDataAdapter(command)
                Dim dataset As New DataSet()

                Try
                    adapter.Fill(dataset, "QueryResult")
                    dataGridView1.ItemsSource = dataset.Tables("QueryResult").DefaultView
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "SQL query error")
                End Try
            ElseIf TypeOf queryBuilder.MetadataProvider Is ODBCMetadataProvider Then
                Dim command As OdbcCommand = DirectCast(queryBuilder.MetadataProvider.Connection.CreateCommand(), OdbcCommand)
                command.CommandText = queryBuilder.SQL

                ' handle the query parameters
                If queryBuilder.Parameters.Count > 0 Then
                    For i As Integer = 0 To queryBuilder.Parameters.Count - 1
                        If Not command.Parameters.Contains(queryBuilder.Parameters(i).FullName) Then
                            Dim parameter As New OdbcParameter()
                            parameter.ParameterName = queryBuilder.Parameters(i).FullName
                            parameter.DbType = queryBuilder.Parameters(i).DataType
                            command.Parameters.Add(parameter)
                        End If
                    Next

                    Dim qpf As New QueryParametersWindow(command)

                    qpf.ShowDialog()
                End If

                Dim adapter As New OdbcDataAdapter(command)
                Dim dataset As New DataSet()

                Try
                    adapter.Fill(dataset, "QueryResult")
                    dataGridView1.ItemsSource = dataset.Tables("QueryResult").DefaultView
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "SQL query error")
                End Try
            End If
        End If
    End Sub

    Private Sub SqlTextEditor1_OnTextChanged(sender As Object, e As EventArgs)
        ErrorBox.Message = string.Empty
    End Sub
End Class
