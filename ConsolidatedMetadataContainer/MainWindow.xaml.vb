'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data.SqlClient
Imports ActiveQueryBuilder.Core

Class MainWindow
    Private ReadOnly _connections As New Dictionary(Of String, SQLContext)()
    Private ReadOnly _consolidatedToInner As New Dictionary(Of MetadataItem, MetadataItem)()
    Private ReadOnly _innerToConsolidated As New Dictionary(Of MetadataItem, MetadataItem)()

    Public Sub New()
        InitializeComponent()

        ' first connection
        Dim xmlNorthwind As SQLContext = New SQLContext() With { 
                .SyntaxProvider = New MSSQLSyntaxProvider() _
                }
        xmlNorthwind.MetadataContainer.ImportFromXML("northwind.xml")
        _connections.Add("xml", xmlNorthwind)

        ' second connection
        Dim mssqlAdventureWorks As SQLContext = New SQLContext() With { 
                .SyntaxProvider = New MSSQLSyntaxProvider(),
                .MetadataProvider = New MSSQLMetadataProvider() With { 
                .Connection = New SqlConnection("Server=sql2014;Database=AdventureWorks;User Id=sa;Password=********;")
                } _
                }
        _connections.Add("live", mssqlAdventureWorks)

        ' QueryBuilder with consolidated metadata
        AddHandler QBuilder.MetadataContainer.ItemMetadataLoading, AddressOf MetadataContainerOnItemMetadataLoading

        QBuilder.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MetadataContainerOnItemMetadataLoading(sender As Object, item As MetadataItem, loadTypes As MetadataType)
        ' root of consolidated metadata contains connections
        If item Is QBuilder.MetadataContainer AndAlso loadTypes.Contains(MetadataType.Connection) Then
            ' add connections (as virtual "Connection" objects)
            For Each connectionDescription As KeyValuePair(Of string, SQLContext) In _connections
                Dim connectionName as string= connectionDescription.Key
                Dim connection as SQLContext = connectionDescription.Value
                Dim innerItem As MetadataContainer = connection.MetadataContainer

                If _innerToConsolidated.ContainsKey(innerItem) Then
                    Continue For
                End If

                Dim newItem as MetadataNamespace = item.AddConnection(connectionName)
                newItem.Items = innerItem.Items

                MapConsolidatedToInnerRecursive(newItem, innerItem)
            Next

            Return
        End If

        ' find "inner" item, load it's children and copy them to consolidated container
        If True Then
            Dim innerItem As MetadataItem = _consolidatedToInner(item)
            innerItem.Items.Load(loadTypes, False)

            For Each childItem As MetadataItem In innerItem.Items
                If Not loadTypes.Contains(childItem.Type) Then
                    Continue For
                End If

                If _innerToConsolidated.ContainsKey(childItem) Then
                    Continue For
                End If

                Dim newItem as MetadataItem= childItem.Clone(item.Items)
                item.Items.Add(newItem)

                MapConsolidatedToInnerRecursive(newItem, childItem)
            Next
        End If
    End Sub

    Private Sub MapConsolidatedToInnerRecursive(consolidated As MetadataItem, inner As MetadataItem)
        _consolidatedToInner.Add(consolidated, inner)
        _innerToConsolidated.Add(inner, consolidated)

        For i = 0 To inner.Items.Count - 1
            MapConsolidatedToInnerRecursive(consolidated.Items(i), inner.Items(i))
        Next
    End Sub
End Class
