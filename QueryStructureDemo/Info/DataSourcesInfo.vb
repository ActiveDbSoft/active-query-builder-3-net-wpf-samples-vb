''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Collections.Generic
Imports System.Text
Imports ActiveQueryBuilder.Core

Namespace Info
    Public Class DataSourcesInfo
        Public Shared Function GetDataSourceList(ByVal unionSubQuery As UnionSubQuery) As List(Of DataSource)
            Dim list = New List(Of DataSource)()

            unionSubQuery.FromClause.GetDatasourceByClass(list)

            Return list
        End Function

        Private Shared Sub DumpDataSourceInfo(ByVal stringBuilder As StringBuilder, ByVal dataSource As DataSource)
            ' write full sql fragment
            stringBuilder.AppendLine(dataSource.GetResultSQL())

            ' write alias
            stringBuilder.AppendLine("  alias: " & dataSource.Alias)

            ' write referenced MetadataObject (if any)
            If dataSource.MetadataObject IsNot Nothing Then
                stringBuilder.AppendLine("  ref: " & dataSource.MetadataObject.Name)
            End If

            ' write subquery (if datasource is actually a derived table)
            If TypeOf dataSource Is DataSourceQuery Then
                stringBuilder.AppendLine("  subquery sql: " & CType(dataSource, DataSourceQuery).GetResultSQL())
            End If

            ' write fields
            Dim fields = String.Empty

            For i = 0 To dataSource.Metadata.Count - 1
                If fields.Length > 0 Then
                    fields &= ", "
                End If

                fields &= dataSource.Metadata(i).Name
            Next i

            stringBuilder.AppendLine("  fields (" & dataSource.Metadata.Count & "): " & fields)
        End Sub

        Private Shared Sub DumpDataSourcesInfo(ByVal stringBuilder As StringBuilder, ByVal dataSources As List(Of DataSource))
            For i = 0 To dataSources.Count - 1
                If stringBuilder.Length > 0 Then
                    stringBuilder.AppendLine()
                End If

                DumpDataSourceInfo(stringBuilder, dataSources(i))
            Next i
        End Sub

        Public Shared Sub DumpDataSourcesInfoFromUnionSubQuery(ByVal stringBuilder As StringBuilder, ByVal unionSubQuery As UnionSubQuery)
            DumpDataSourcesInfo(stringBuilder, GetDataSourceList(unionSubQuery))
        End Sub
    End Class
End Namespace
