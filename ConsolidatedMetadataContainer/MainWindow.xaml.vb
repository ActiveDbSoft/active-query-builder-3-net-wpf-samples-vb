//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System.Data.SqlClient
Imports ActiveQueryBuilder.Core

Class MainWindow
    Private Shared Function InitializeConnections() As Dictionary(Of String, SQLContext)
        Dim result As new Dictionary(Of String, SQLContext)()

        ' first connection
        Dim innerXml = New SQLContext() With {
            .SyntaxProvider = new MSSQLSyntaxProvider()
        }
        innerXml.MetadataContainer.ImportFromXML("northwind.xml")
        result.Add("xml", innerXml)

        ' second connection
        Dim innerMsSql = New SQLContext() With {
            .SyntaxProvider = New MSSQLSyntaxProvider(),
            .MetadataProvider = New MSSQLMetadataProvider() With {
                .Connection = New SqlConnection("Server=sql2014;Database=AdventureWorks;User Id=sa;Password=********;")
            }
        }
        result.Add("live", innerMsSql)

        Return result
    End Function

    Private ReadOnly _connections As Dictionary(Of String, SQLContext) = InitializeConnections()

    Public Sub New()
        InitializeComponent()

        ' add connections
        Dim metadataContainer = QBuilder.MetadataContainer
        For Each pair As KeyValuePair(Of String,SQLContext) In _connections
            Dim connectionName = pair.Key
            Dim innerContext = pair.Value
            metadataContainer.AddConnection(connectionName, innerContext)
        Next

        QBuilder.InitializeDatabaseSchemaTree()
    End Sub
End Class
