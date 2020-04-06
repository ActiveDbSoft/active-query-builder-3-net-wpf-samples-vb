'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Linq
Imports ActiveQueryBuilder.Core
Imports Windows.QueryInformationWindows

Namespace GeneralAssembly
	Public Class SqlHelpers
		Private Const AtNameParamFormat As String = "@s"
		Private Const ColonNameParamFormat As String = ":s"
		Private Const QuestionParamFormat As String = "?"
		Private Const QuestionNumberParamFormat As String = "?n"
		Private Const QuestionNameParamFormat As String = "?s"

		Public Shared Function GetAcceptableParametersFormats(metadataProvider As BaseMetadataProvider, syntaxProvider As BaseSyntaxProvider) As List(Of String)
			If TypeOf metadataProvider Is MSSQLMetadataProvider Then
				Return New List(Of String) From {AtNameParamFormat}
			End If

			If TypeOf metadataProvider Is OracleNativeMetadataProvider Then
				Return New List(Of String) From {ColonNameParamFormat}
			End If

			If TypeOf metadataProvider Is PostgreSQLMetadataProvider Then
				Return New List(Of String) From {ColonNameParamFormat}
			End If

			If TypeOf metadataProvider Is MySQLMetadataProvider Then
				Return New List(Of String) From {AtNameParamFormat, QuestionParamFormat, QuestionNumberParamFormat, QuestionNameParamFormat}
			End If

			If TypeOf metadataProvider Is OLEDBMetadataProvider Then
				If TypeOf syntaxProvider Is MSAccessSyntaxProvider Then
					Return New List(Of String) From {AtNameParamFormat, ColonNameParamFormat, QuestionParamFormat}
				End If

				If TypeOf syntaxProvider Is MSSQLSyntaxProvider Then
					Return New List(Of String) From {QuestionParamFormat}
				End If

				If TypeOf syntaxProvider Is OracleSyntaxProvider Then
					Return New List(Of String) From {ColonNameParamFormat, QuestionParamFormat, QuestionNumberParamFormat}
				End If

				If TypeOf syntaxProvider Is DB2SyntaxProvider Then
					Return New List(Of String) From {QuestionParamFormat}
				End If
			End If

			If TypeOf metadataProvider Is ODBCMetadataProvider Then
				If TypeOf syntaxProvider Is MSAccessSyntaxProvider Then
					Return New List(Of String) From {QuestionParamFormat}
				End If

				If TypeOf syntaxProvider Is MSSQLSyntaxProvider Then
					Return New List(Of String) From {QuestionParamFormat}
				End If

				If TypeOf syntaxProvider Is MySQLSyntaxProvider Then
					Return New List(Of String) From {QuestionParamFormat}
				End If

				If TypeOf syntaxProvider Is PostgreSQLSyntaxProvider Then
					Return New List(Of String) From {QuestionParamFormat}
				End If

				If TypeOf syntaxProvider Is OracleSyntaxProvider Then
					Return New List(Of String) From {ColonNameParamFormat, QuestionParamFormat, QuestionNumberParamFormat}
				End If

				If TypeOf syntaxProvider Is DB2SyntaxProvider Then
					Return New List(Of String) From {QuestionParamFormat}
				End If
			End If

			Return New List(Of String)()
		End Function

		Public Shared Function CheckParameters(metadataProvider As BaseMetadataProvider, syntaxProvider As BaseSyntaxProvider, parameters As ParameterList) As Boolean
			Dim acceptableFormats = GetAcceptableParametersFormats(metadataProvider, syntaxProvider)

			If acceptableFormats.Count = 0 Then
				Return True
			End If

			For Each parameter In parameters
				If Not acceptableFormats.Any(Function(x) IsSatisfiesFormat(parameter.FullName, x)) Then
					Return False
				End If
			Next parameter

			Return True
		End Function

		Private Shared Function IsSatisfiesFormat(name As String, format As String) As Boolean
			If String.IsNullOrEmpty(name) OrElse String.IsNullOrEmpty(format) Then
				Return False
			End If

			If format.Chars(0) <> name.Chars(0) Then
				Return False
			End If

			Dim lastChar = format.Last()
			If lastChar = "?"c Then
				Return name = format
			End If

			If lastChar = "s"c Then
				Return name.Length > 1 AndAlso Char.IsLetter(name.Chars(1))
			End If

			If lastChar = "n"c Then
				If name.Length = 1 Then
					Return False
				End If

				For Each c In name
					If Not Char.IsDigit(c) Then
						Return False
					End If
				Next c

				Return True
			End If

			Return False
		End Function

		Private Shared Function CreateSqlCommand(sqlCommand As String, sqlQuery As SQLQuery) As DbCommand
			Dim command As DbCommand = DirectCast(sqlQuery.SQLContext.MetadataProvider.Connection.CreateCommand(), DbCommand)
			command.CommandText = sqlCommand

			' handle the query parameters
			If sqlQuery.QueryParameters.Count <= 0 Then
				Return command
			End If

			For Each p As Parameter In sqlQuery.QueryParameters.Where(Function(item) Not command.Parameters.Contains(item.FullName))
				Dim parameter As SqlParameter = New SqlParameter With {
					.ParameterName = p.FullName,
					.DbType = p.DataType
				}
				command.Parameters.Add(parameter)
			Next p
			Dim qpf = New QueryParametersWindow(command)
			qpf.ShowDialog()
			Return command
		End Function

		Public Shared Function GetDataView(sqlCommand As String, sqlQuery As SQLQuery) As DataView
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
				Next i

				Do While dbReader.Read()
					Dim values(dbReader.FieldCount - 1) As Object
					dbReader.GetValues(values)
					table.Rows.Add(values)
				Loop

				Return table.DefaultView
			End Using
		End Function
	End Class
End Namespace
