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
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Linq
Imports ActiveQueryBuilder.Core
Imports FullFeaturedDemo.Windows

Namespace Common
    Public NotInheritable Class Helpers
        Private Sub New()
        End Sub
        Public Shared Function CreateSqlCommand(sqlCommand As String, sqlQuery As SQLQuery) As DbCommand
            Dim command As DbCommand = DirectCast(sqlQuery.SQLContext.MetadataProvider.Connection.CreateCommand(), DbCommand)
            command.CommandText = sqlCommand

            ' handle the query parameters
            If sqlQuery.QueryParameters.Count <= 0 Then
                Return command
            End If

            For Each p As Parameter In sqlQuery.QueryParameters.Where(Function(item) Not command.Parameters.Contains(item.FullName))
                Dim parameter As New SqlParameter() With {
                    .ParameterName = p.FullName,
                    .DbType = p.DataType
                }
                command.Parameters.Add(parameter)
            Next
            Dim qpf = New QueryParametersWindow(command)
            qpf.ShowDialog()
            Return command
        End Function

        Public Shared Function ExecuteSql(sqlCommand As String, sqlQuery As SQLQuery) As DataView
            If String.IsNullOrEmpty(sqlCommand) Then
                Return Nothing
            End If

            If sqlQuery.SQLContext.MetadataProvider Is Nothing Then
                Return Nothing
            End If

            If Not sqlQuery.SQLContext.MetadataProvider.Connected Then
                sqlQuery.SQLContext.MetadataProvider.Connect()
            End If

            If String.IsNullOrEmpty(sqlCommand) Then
                Return Nothing
            End If

            If Not sqlQuery.SQLContext.MetadataProvider.Connected Then
                sqlQuery.SQLContext.MetadataProvider.Connect()
            End If

            Dim command = CreateSqlCommand(sqlCommand, sqlQuery)

            Dim table As New DataTable("result")

            Using dbReader = command.ExecuteReader()

                For i As Integer = 0 To dbReader.FieldCount - 1
                    table.Columns.Add(dbReader.GetName(i))
                Next

                While dbReader.Read()
                    Dim values As Object() = New Object(dbReader.FieldCount - 1) {}
                    dbReader.GetValues(values)
                    table.Rows.Add(values)
                End While

                Return table.DefaultView
            End Using
        End Function
    End Class
End Namespace
