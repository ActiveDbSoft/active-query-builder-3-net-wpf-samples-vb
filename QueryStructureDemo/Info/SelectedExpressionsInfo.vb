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
Imports System.Data
Imports System.Text
Imports ActiveQueryBuilder.Core

Namespace Info
    Public Class SelectedExpressionsInfo
        Private Shared Sub DumpSelectedExpressionInfo(ByVal stringBuilder As StringBuilder, ByVal selectedExpression As QueryColumnListItem)
            ' write full sql fragment of selected expression
            stringBuilder.AppendLine(selectedExpression.ExpressionString)

            ' write alias
            If Not String.IsNullOrEmpty(selectedExpression.AliasString) Then
                stringBuilder.AppendLine("  alias: " & selectedExpression.AliasString)
            End If

            ' write datasource reference (if any)
            If selectedExpression.ExpressionDatasource IsNot Nothing Then
                stringBuilder.AppendLine("  datasource: " & selectedExpression.ExpressionDatasource.GetResultSQL())
            End If

            ' write metadata information (if any)
            If selectedExpression.ExpressionField Is Nothing Then
                Return
            End If

            Dim field = selectedExpression.ExpressionField
            stringBuilder.AppendLine("  field name: " & field.Name)

            Dim s = System.Enum.GetName(GetType(DbType), field.FieldType)
            stringBuilder.AppendLine("  field type: " & s)
        End Sub

        Public Shared Sub DumpSelectedExpressionsInfoFromUnionSubQuery(ByVal stringBuilder As StringBuilder, ByVal unionSubQuery As UnionSubQuery)
            ' get list of CriteriaItems
            Dim criteriaList As QueryColumnList = unionSubQuery.QueryColumnList

            ' dump all items
            For i As Integer = 0 To criteriaList.Count - 1
                Dim criteriaItem As QueryColumnListItem = criteriaList(i)

                ' only items have .Select property set to True goes to SELECT list
                If Not criteriaItem.Selected Then
                    Continue For
                End If

                ' separator
                If stringBuilder.Length > 0 Then
                    stringBuilder.AppendLine()
                End If

                DumpSelectedExpressionInfo(stringBuilder, criteriaItem)
                DumpSelectedExpressionsStatistics(stringBuilder, criteriaItem)
            Next i
        End Sub

        Private Shared Sub DumpSelectedExpressionsStatistics(ByVal stringBuilder As StringBuilder, ByVal criteriaItem As QueryColumnListItem)
            stringBuilder.AppendLine()
            stringBuilder.AppendLine("Statistic:")
            StatisticsInfo.DumpQueryStatisticsInfo(stringBuilder, criteriaItem.QueryStatistics)
        End Sub
    End Class
End Namespace
