''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Text
Imports ActiveQueryBuilder.Core

Namespace Info
    Public Class SubQueryStructureInfo
        Private Shared Sub DumpUnionGroupInfo(ByVal stringBuilder As StringBuilder, ByVal indent As String, ByVal unionGroup As UnionGroup)
            Dim children = GetUnionChildren(unionGroup)

            For Each child In children
                If stringBuilder.Length > 0 Then
                    stringBuilder.AppendLine()
                End If

                If TypeOf child Is UnionSubQuery Then
                    ' UnionSubQuery is a leaf node of query structure.
                    ' It represent a single SELECT statement in the tree of unions
                    DumpUnionSubQueryInfo(stringBuilder, indent, CType(child, UnionSubQuery))
                Else
                    ' UnionGroup is a tree node.
                    ' It contains one or more leafs of other tree nodes.
                    ' It represent a root of the subquery of the union tree or a
                    ' parentheses in the union tree.
                    Debug.Assert(TypeOf child Is UnionGroup)

                    unionGroup = CType(child, UnionGroup)

                    stringBuilder.AppendLine(indent & unionGroup.UnionOperatorFull & "group: [")
                    DumpUnionGroupInfo(stringBuilder, indent & "    ", unionGroup)
                    stringBuilder.AppendLine(indent & "]")
                End If
            Next child
        End Sub

        Private Shared Sub DumpUnionSubQueryInfo(ByVal stringBuilder As StringBuilder, ByVal indent As String, ByVal unionSubQuery As UnionSubQuery)
            Dim sql As String = unionSubQuery.GetResultSQL()

            stringBuilder.AppendLine(indent & unionSubQuery.UnionOperatorFull & ": " & sql)
        End Sub

        Private Shared Function GetUnionChildren(ByVal unionGroup As UnionGroup) As IEnumerable(Of QueryBase)
            Dim result = New ArrayList()

            For i = 0 To unionGroup.Count - 1
                result.Add(unionGroup(i))
            Next i

            Return DirectCast(result.ToArray(GetType(QueryBase)), QueryBase())
        End Function

        Public Shared Sub DumpQueryStructureInfo(ByVal stringBuilder As StringBuilder, ByVal subQuery As SubQuery)
            DumpUnionGroupInfo(stringBuilder, "", subQuery)
        End Sub
    End Class
End Namespace
