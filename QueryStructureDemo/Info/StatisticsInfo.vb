''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Text
Imports ActiveQueryBuilder.Core

Namespace Info
    Public Class StatisticsInfo
        Public Shared Sub DumpUsedObjectsInfo(ByVal stringBuilder As StringBuilder, ByVal usedObjects As StatisticsDatabaseObjectList)
            stringBuilder.AppendLine("Used Objects (" & usedObjects.Count & "):")

            For i = 0 To usedObjects.Count - 1
                stringBuilder.AppendLine(usedObjects(i).ObjectName.QualifiedName)
            Next i
        End Sub

        Public Shared Sub DumpUsedColumnsInfo(ByVal stringBuilder As StringBuilder, ByVal usedColumns As StatisticsFieldList)
            stringBuilder.AppendLine("Used Columns (" & usedColumns.Count & "):")

            For i As Integer = 0 To usedColumns.Count - 1
                stringBuilder.AppendLine(usedColumns(i).ObjectName.QualifiedName)
            Next i
        End Sub

        Public Shared Sub DumpOutputExpressionsInfo(ByVal stringBuilder As StringBuilder, ByVal outputExpressions As StatisticsOutputColumnList)
            stringBuilder.AppendLine("Output Expressions (" & outputExpressions.Count & "):")

            For i As Integer = 0 To outputExpressions.Count - 1
                stringBuilder.AppendLine(outputExpressions(i).Expression)
            Next i
        End Sub

        Public Shared Sub DumpQueryStatisticsInfo(ByVal stringBuilder As StringBuilder, ByVal queryStatistics As QueryStatistics)
            DumpUsedObjectsInfo(stringBuilder, queryStatistics.UsedDatabaseObjects)

            stringBuilder.AppendLine()
            DumpUsedColumnsInfo(stringBuilder, queryStatistics.UsedDatabaseObjectFields)

            stringBuilder.AppendLine()
            DumpOutputExpressionsInfo(stringBuilder, queryStatistics.OutputColumns)
        End Sub
    End Class
End Namespace
