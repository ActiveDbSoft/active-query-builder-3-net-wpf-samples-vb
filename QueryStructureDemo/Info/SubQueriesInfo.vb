''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Text
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF

Namespace Info
	Public Class SubQueriesInfo
		Private Shared Sub DumpSubQueryInfo(ByVal stringBuilder As StringBuilder, ByVal index As Integer, ByVal subQuery As SubQuery)
			Dim sql = subQuery.GetResultSQL()

			stringBuilder.AppendLine(index & ": " & sql)
		End Sub

		Public Shared Sub DumpSubQueriesInfo(ByVal stringBuilder As StringBuilder, ByVal queryBuilder As QueryBuilder)
			For i = 0 To queryBuilder.ActiveUnionSubQuery.QueryRoot.SubQueryCount - 1
				If stringBuilder.Length > 0 Then
					stringBuilder.AppendLine()
				End If

				DumpSubQueryInfo(stringBuilder, i, queryBuilder.ActiveUnionSubQuery.QueryRoot.SubQueries(i))
				DumpSubQueryStatistics(stringBuilder, queryBuilder.ActiveUnionSubQuery.QueryRoot.SubQueries(i))
			Next i
		End Sub

		Private Shared Sub DumpSubQueryStatistics(ByVal stringBuilder As StringBuilder, ByVal subQuery As SubQuery)
			stringBuilder.AppendLine()
			stringBuilder.AppendLine("Subquery statistic:")
			StatisticsInfo.DumpQueryStatisticsInfo(stringBuilder, subQuery.QueryStatistics)
		End Sub
	End Class
End Namespace
