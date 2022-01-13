''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports LoadMetadataDemo.ConnectionWindows

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Public Partial Class MainWindow
    Inherits Window
    Private _dbConnection As IDbConnection
    Private _way3EventMetadataProvider As EventMetadataProvider
    Private _lastValidSql As String
    Private _errorPosition As Integer = -1

    Public Sub New()
        InitializeComponent()
    End Sub

    '''///////////////////////////////////////////////////////////////////////
    ''' 1st way:
    ''' This method demonstrates the direct access to the internal metadata 
    ''' objects collection (MetadataContainer).
    '''///////////////////////////////////////////////////////////////////////
    Private Sub btn1Way_Click(sender As Object, e As EventArgs)
        ' prevent QueryBuilder to request metadata
        QBuilder.MetadataLoadingOptions.OfflineMode = True

        QBuilder.MetadataProvider = Nothing

        Dim metadataContainer As MetadataContainer = QBuilder.MetadataContainer
        metadataContainer.BeginUpdate()

        Try
            metadataContainer.Clear()

            Dim schemaDbo As MetadataNamespace = metadataContainer.AddSchema("dbo")

            ' prepare metadata for table "Orders"
            Dim orders As MetadataObject = schemaDbo.AddTable("Orders")
            ' fields
            orders.AddField("OrderId")
            orders.AddField("CustomerId")

            ' prepare metadata for table "Order Details"
            Dim orderDetails As MetadataObject = schemaDbo.AddTable("Order Details")
            ' fields
            orderDetails.AddField("OrderId")
            orderDetails.AddField("ProductId")
            ' foreign keys
            Dim foreignKey As MetadataForeignKey = orderDetails.AddForeignKey("OrderDetailsToOrders")

            Using referencedName As New MetadataQualifiedName()
                referencedName.Add("Orders")
                referencedName.Add("dbo")

                foreignKey.ReferencedObjectName = referencedName
            End Using

            foreignKey.Fields.Add("OrderId")
            foreignKey.ReferencedFields.Add("OrderId")
        Finally
            metadataContainer.EndUpdate()
        End Try

        QBuilder.InitializeDatabaseSchemaTree()
    End Sub

    '''///////////////////////////////////////////////////////////////////////
    ''' 2rd way:
    ''' This method demonstrates on-demand manual filling of metadata structure using 
    ''' corresponding MetadataContainer.ItemMetadataLoading event
    '''///////////////////////////////////////////////////////////////////////
    Private Sub btn2Way_Click(sender As Object, e As EventArgs)
        ' allow QueryBuilder to request metadata
        QBuilder.MetadataLoadingOptions.OfflineMode = False

        QBuilder.MetadataProvider = Nothing
        AddHandler QBuilder.MetadataContainer.ItemMetadataLoading, AddressOf way2ItemMetadataLoading
        QBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub way2ItemMetadataLoading(sender As Object, item As MetadataItem, types As MetadataType)
        Select Case item.Type
            Case MetadataType.Root
                If (types And MetadataType.Schema) > 0 Then
                    item.AddSchema("dbo")
                End If
                Exit Select

            Case MetadataType.Schema
                If (item.Name = "dbo") AndAlso (types And MetadataType.Table) > 0 Then
                    item.AddTable("Orders")
                    item.AddTable("Order Details")
                End If
                Exit Select

            Case MetadataType.Table
                If item.Name = "Orders" Then
                    If (types And MetadataType.Field) > 0 Then
                        item.AddField("OrderId")
                        item.AddField("CustomerId")
                    End If
                ElseIf item.Name = "Order Details" Then
                    If (types And MetadataType.Field) > 0 Then
                        item.AddField("OrderId")
                        item.AddField("ProductId")
                    End If

                    If (types And MetadataType.ForeignKey) > 0 Then
                        Dim foreignKey As MetadataForeignKey = item.AddForeignKey("OrderDetailsToOrder")
                        foreignKey.Fields.Add("OrderId")
                        foreignKey.ReferencedFields.Add("OrderId")
                        Using name As New MetadataQualifiedName()
                            name.Add("Orders")
                            name.Add("dbo")

                            foreignKey.ReferencedObjectName = name
                        End Using
                    End If
                End If
                Exit Select
        End Select

        item.Items.SetLoaded(types, True)
    End Sub

    '''///////////////////////////////////////////////////////////////////////
    ''' 3rd way:
    '''
    ''' This method demonstrates loading of metadata through .NET data providers 
    ''' unsupported by our QueryBuilder component. If such data provider is able 
    ''' to execute SQL queries, you can use our EventMetadataProvider with handling 
    ''' it's ExecSQL event. In this event the EventMetadataProvider will provide 
    ''' you SQL queries it use for the metadata retrieval. You have to execute 
    ''' a query and return resulting data reader object.
    '''
    ''' Note: In this sample we are using GenericSyntaxProvider that tries 
    ''' to detect the the server type automatically. In some conditions it's unable 
    ''' to detect the server type and it also has limited syntax parsing abilities. 
    ''' For this reason you have to use specific syntax providers in your application, 
    ''' e.g. MySQLSyntaxProver, OracleSyntaxProvider, etc.
    '''///////////////////////////////////////////////////////////////////////
    Private Sub btn3Way_Click(sender As Object, e As EventArgs)
        If _dbConnection IsNot Nothing Then
            Try
                _dbConnection.Close()
                _dbConnection.Open()

                ' allow QueryBuilder to request metadata
                QBuilder.MetadataLoadingOptions.OfflineMode = False

                ResetQueryBuilderMetadata()

                QBuilder.MetadataProvider = _way3EventMetadataProvider
                QBuilder.InitializeDatabaseSchemaTree()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "btn3Way_Click()")
            End Try
        Else
            MessageBox.Show("Please setup a database connection by clicking on the ""Connect"" menu item before testing this method.")
        End If
    End Sub

    Private Sub way3EventMetadataProvider_ExecSQL(metadataProvider As BaseMetadataProvider, sql As String, schemaOnly As Boolean, ByRef dataReader As IDataReader)
        dataReader = Nothing

        If _dbConnection IsNot Nothing Then
            Dim command As IDbCommand = _dbConnection.CreateCommand()
            command.CommandText = sql
            dataReader = command.ExecuteReader()
        End If
    End Sub

    '''///////////////////////////////////////////////////////////////////////
    ''' 4th way:
    ''' This method demonstrates manual filling of metadata structure from 
    ''' stored DataSet.
    '''///////////////////////////////////////////////////////////////////////
    Private Sub btn4Way_Click(sender As Object, e As EventArgs)
        QBuilder.MetadataProvider = Nothing
        QBuilder.MetadataLoadingOptions.OfflineMode = True
        ' prevent QueryBuilder to request metadata from connection
        Dim dataSet As New DataSet()

        ' Load sample dataset created in the Visual Studio with Dataset Designer
        ' and exported to XML using WriteXmlSchema() method.
        dataSet.ReadXmlSchema("StoredDataSetSchema.xml")

        QBuilder.MetadataContainer.BeginUpdate()

        Try
            QBuilder.ClearMetadata()

            ' add tables
            For Each table As DataTable In dataSet.Tables
                ' add new metadata table
                Dim metadataTable As MetadataObject = QBuilder.MetadataContainer.AddTable(table.TableName)

                ' add metadata fields (columns)
                For Each column As DataColumn In table.Columns
                    ' create new field
                    Dim metadataField As MetadataField = metadataTable.AddField(column.ColumnName)
                    ' setup field
                    metadataField.FieldType = TypeToDbType(column.DataType)
                    metadataField.Nullable = column.AllowDBNull
                    metadataField.[ReadOnly] = column.[ReadOnly]

                    If column.MaxLength <> -1 Then
                        metadataField.Size = column.MaxLength
                    End If

                    ' detect the field is primary key
                    For Each pkColumn As DataColumn In table.PrimaryKey
                        If column Is pkColumn Then
                            metadataField.PrimaryKey = True
                        End If
                    Next
                Next

                ' add relations
                For Each relation As DataRelation In table.ParentRelations
                    ' create new relation on the parent table
                    Dim metadataRelation As MetadataForeignKey = metadataTable.AddForeignKey(relation.RelationName)

                    ' set referenced table
                    Using referencedName As New MetadataQualifiedName()
                        referencedName.Add(relation.ParentTable.TableName)

                        metadataRelation.ReferencedObjectName = referencedName
                    End Using

                    ' set referenced fields
                    For Each parentColumn As DataColumn In relation.ParentColumns
                        metadataRelation.ReferencedFields.Add(parentColumn.ColumnName)
                    Next

                    ' set fields
                    For Each childColumn As DataColumn In relation.ChildColumns
                        metadataRelation.Fields.Add(childColumn.ColumnName)
                    Next
                Next
            Next
        Finally
            QBuilder.MetadataContainer.EndUpdate()
        End Try

        QBuilder.InitializeDatabaseSchemaTree()
    End Sub


    ' =============================================================================


    Private Sub aboutMenuItem_Click(sender As Object, e As EventArgs)
        QueryBuilder.ShowAboutDialog()
    End Sub

    Private Sub QBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
        If SqlTextBox Is Nothing Then
            Return
        End If
        ' Handle the event raised by SQL builder object that the text of SQL query is changed

        ' Hide error banner if any
        ErrorBox.Visibility = Visibility.Collapsed

        ' update the text box
        SqlTextBox.Text = QBuilder.FormattedSQL
        _lastValidSql = QBuilder.FormattedSQL
    End Sub



    Private Sub connectToMSSQLServerMenuItem_Click(sender As Object, e As EventArgs)
        ' Connect to MS SQL Server

        Dim f As New MSSQLConnectionWindow()

        If f.ShowDialog() = True Then
            If _dbConnection IsNot Nothing Then
                _dbConnection.Close()
                _dbConnection.Dispose()
            End If

            _dbConnection = New SqlConnection(f.ConnectionString)
        End If

    End Sub

    Private Sub connectToAccessDatabaseMenuItem_Click(sender As Object, e As EventArgs)
        ' Connect to MS Access database using OLE DB provider

        Dim f As New AccessConnectionWindow()

        If f.ShowDialog() = True Then
            If _dbConnection IsNot Nothing Then
                _dbConnection.Close()
                _dbConnection.Dispose()
            End If

            _dbConnection = New OleDbConnection(f.ConnectionString)
        End If
    End Sub

    Private Sub connectOleDbMenuItem_Click(sender As Object, e As EventArgs)
        ' Connect to a database through the OLE DB provider

        Dim f As New OLEDBConnectionWindow()

        If f.ShowDialog() = True Then
            If _dbConnection IsNot Nothing Then
                _dbConnection.Close()
                _dbConnection.Dispose()
            End If

            _dbConnection = New OleDbConnection(f.ConnectionString)
        End If
    End Sub

    Private Sub connectODBCMenuItem_Click(sender As Object, e As EventArgs)
        ' Connect to a database through the ODBC provider

        Dim f As New ODBCConnectionWindow()

        If f.ShowDialog() = True Then
            If _dbConnection IsNot Nothing Then
                _dbConnection.Close()
                _dbConnection.Dispose()
            End If

            _dbConnection = New OdbcConnection(f.ConnectionString)
        End If
    End Sub

    Private Shared Function TypeToDbType(type As Type) As DbType
        If type Is GetType(String) Then
            Return DbType.[String]
        End If
        If type Is GetType(Short) Then
            Return DbType.Int16
        End If
        If type Is GetType(Integer) Then
            Return DbType.Int32
        End If
        If type Is GetType(Long) Then
            Return DbType.Int64
        End If
        If type Is GetType(UShort) Then
            Return DbType.UInt16
        End If
        If type Is GetType(UInteger) Then
            Return DbType.UInt32
        End If
        If type Is GetType(ULong) Then
            Return DbType.UInt64
        End If
        If type Is GetType(Boolean) Then
            Return DbType.[Boolean]
        End If
        If type Is GetType(Single) Then
            Return DbType.[Single]
        End If
        If type Is GetType(Double) Then
            Return DbType.[Double]
        End If
        If type Is GetType(Decimal) Then
            Return DbType.[Decimal]
        End If
        If type Is GetType(DateTime) Then
            Return DbType.DateTime
        End If
        If type Is GetType(TimeSpan) Then
            Return DbType.Time
        End If
        If type Is GetType(Byte) Then
            Return DbType.[Byte]
        End If
        If type Is GetType(SByte) Then
            Return DbType.[SByte]
        End If
        If type Is GetType(Char) Then
            Return DbType.[String]
        End If
        If type Is GetType(Byte()) Then
            Return DbType.Binary
        End If
        If type Is GetType(Guid) Then
            Return DbType.Guid
        End If
        Return DbType.[Object]
    End Function

    Private Sub ResetQueryBuilderMetadata()
        QBuilder.MetadataProvider = Nothing
        QBuilder.ClearMetadata()
        RemoveHandler QBuilder.MetadataContainer.ItemMetadataLoading, AddressOf way2ItemMetadataLoading
    End Sub

    Private Sub SqlTextBox_OnLostFocus(sender As Object, e As RoutedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QBuilder.SQL = SqlTextBox.Text

            ' Hide error banner if any
            ErrorBox.Visibility = Visibility.Collapsed
            _lastValidSql = SqlTextBox.Text
        Catch ex As SQLParsingException
            ' Set caret to error position
            SqlTextBox.SelectionStart = ex.ErrorPos.pos

            ' Show banner with error text
            ErrorBox.Show(ex.Message, QBuilder.SyntaxProvider)
            _errorPosition = ex.ErrorPos.pos
        End Try
    End Sub


    Private Sub ErrorBox_OnGoToErrorPosition(sender As Object, e As EventArgs)
        SqlTextBox.Focus()

        If _errorPosition <> -1 Then
            If SqlTextBox.LineCount <> 1 Then SqlTextBox.ScrollToLine(SqlTextBox.GetLineIndexFromCharacterIndex(_errorPosition))
            SqlTextBox.CaretIndex = _errorPosition
        End If
    End Sub

    Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
        SqlTextBox.Text = _lastValidSql
        SqlTextBox.Focus()
    End Sub
End Class
