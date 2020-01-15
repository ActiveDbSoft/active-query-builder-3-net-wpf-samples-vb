'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core

Namespace Common
    Public NotInheritable Class Helpers
        Private Sub New()
        End Sub
        Public Shared ReadOnly ConnectionDescriptorList As New List(Of Type)() From {
            GetType(MSAccessConnectionDescriptor),
            GetType(ExcelConnectionDescriptor),
            GetType(MSSQLConnectionDescriptor),
            GetType(MSSQLAzureConnectionDescriptor),
            GetType(MySQLConnectionDescriptor),
            GetType(OracleNativeConnectionDescriptor),
            GetType(PostgreSQLConnectionDescriptor),
            GetType(ODBCConnectionDescriptor),
            GetType(OLEDBConnectionDescriptor),
            GetType(SQLiteConnectionDescriptor),
            GetType(FirebirdConnectionDescriptor)
        }

        Public Shared ReadOnly ConnectionDescriptorNames As New List(Of String)() From {
            "MS Access",
            "Excel",
            "MS SQL Server",
            "MS SQL Server Azure",
            "MySQL",
            "Oracle Native",
            "PostgreSQL",
            "Generic ODBC Connection",
            "Generic OLEDB Connection",
            "SQLite",
            "Firebird"
        }

        Private Const AtNameParamFormat As String = "@s"
        Private Const ColonNameParamFormat As String = ":s"
        Private Const QuestionParamFormat As String = "?"
        Private Const QuestionNumberParamFormat As String = "?n"
        Private Const QuestionNameParamFormat As String = "?s"

        Public Shared Function GetAcceptableParametersFormats(metadataProvider As BaseMetadataProvider, syntaxProvider As BaseSyntaxProvider) As List(Of String)
            If TypeOf metadataProvider Is MSSQLMetadataProvider Then
                Return New List(Of String)() From {
                    AtNameParamFormat
                }
            End If

            If TypeOf metadataProvider Is OracleNativeMetadataProvider Then
                Return New List(Of String)() From {
                    ColonNameParamFormat
                }
            End If

            If TypeOf metadataProvider Is PostgreSQLMetadataProvider Then
                Return New List(Of String)() From {
                    ColonNameParamFormat
                }
            End If

            If TypeOf metadataProvider Is MySQLMetadataProvider Then
                Return New List(Of String)() From {
                    AtNameParamFormat,
                    QuestionParamFormat,
                    QuestionNumberParamFormat,
                    QuestionNameParamFormat
                }
            End If

            If TypeOf metadataProvider Is OLEDBMetadataProvider Then
                If TypeOf syntaxProvider Is MSAccessSyntaxProvider Then
                    Return New List(Of String)() From {
                        AtNameParamFormat,
                        ColonNameParamFormat,
                        QuestionParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is MSSQLSyntaxProvider Then
                    Return New List(Of String)() From {
                        QuestionParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is OracleSyntaxProvider Then
                    Return New List(Of String)() From {
                        ColonNameParamFormat,
                        QuestionParamFormat,
                        QuestionNumberParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is DB2SyntaxProvider Then
                    Return New List(Of String)() From {
                        QuestionParamFormat
                    }
                End If
            End If

            If TypeOf metadataProvider Is ODBCMetadataProvider Then
                If TypeOf syntaxProvider Is MSAccessSyntaxProvider Then
                    Return New List(Of String)() From {
                        QuestionParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is MSSQLSyntaxProvider Then
                    Return New List(Of String)() From {
                        QuestionParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is MySQLSyntaxProvider Then
                    Return New List(Of String)() From {
                        QuestionParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is PostgreSQLSyntaxProvider Then
                    Return New List(Of String)() From {
                        QuestionParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is OracleSyntaxProvider Then
                    Return New List(Of String)() From {
                        ColonNameParamFormat,
                        QuestionParamFormat,
                        QuestionNumberParamFormat
                    }
                End If

                If TypeOf syntaxProvider Is DB2SyntaxProvider Then
                    Return New List(Of String)() From {
                        QuestionParamFormat
                    }
                End If
            End If

            Return New List(Of String)()
        End Function

        Public Shared Function CheckParameters(metadataProvider As BaseMetadataProvider, syntaxProvider As BaseSyntaxProvider, parameters As ParameterList) As Boolean
            Dim acceptableFormats As List(Of String) = GetAcceptableParametersFormats(metadataProvider, syntaxProvider)

            If acceptableFormats.Count = 0 Then
                Return True
            End If

            For Each parameter As Parameter In parameters
                If Not acceptableFormats.Any(Function(x) IsSatisfiesFormat(parameter.FullName, x)) Then
                    Return False
                End If
            Next

            Return True
        End Function

        Private Shared Function IsSatisfiesFormat(name As String, format As String) As Boolean
            If String.IsNullOrEmpty(name) OrElse String.IsNullOrEmpty(format) Then
                Return False
            End If

            If format(0) <> name(0) Then
                Return False
            End If

            Dim lastChar As Char = format.Last()
            If lastChar = "?"c Then
                Return name = format
            End If

            If lastChar = "s"c Then
                Return name.Length > 1 AndAlso [Char].IsLetter(name(1))
            End If

            If lastChar = "n"c Then
                If name.Length = 1 Then
                    Return False
                End If

                For Each c As Char In name
                    If Not [Char].IsDigit(c) Then
                        Return False
                    End If
                Next

                Return True
            End If

            Return False
        End Function

        Public Enum SourceType
            File
            [New]
            UserQueries
        End Enum

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

            Dim command As DbCommand = CreateSqlCommand(sqlCommand, sqlQuery)

            Dim table As New DataTable("result")

            Using dbReader As DbDataReader = command.ExecuteReader()

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

    Friend Class SaveExistQueryDialog
        Inherits WindowMessage
        Public Result As System.Nullable(Of Boolean)

        Private Const WIDTH_BUTTON As Integer = 90

        Public Sub New()
            Title = "Save dialog"
            Text = "Save changes?"
            ContentAlignment = HorizontalAlignment.Center
            Result = Nothing

            Dim saveButton As Button = New Button() With {
                .Content = "Save",
                .Width = WIDTH_BUTTON,
                .IsDefault = True
            }
            AddHandler saveButton.Click, Sub()
                                             Result = True
                                             Close()

                                         End Sub

            Dim notSaveButton As Button = New Button() With {
                .Margin = New Thickness(5, 0, 5, 0),
                .Content = "Don't save",
                .Width = WIDTH_BUTTON,
                .IsCancel = True
            }
            AddHandler notSaveButton.Click, Sub()
                                                Result = False
                                                Close()

                                            End Sub

            Dim continueButton As Button = New Button() With {
                .Content = "Continue edit",
                .Width = WIDTH_BUTTON
            }
            AddHandler continueButton.Click, Sub() Close()

            Buttons.Add(saveButton)
            Buttons.Add(notSaveButton)
            Buttons.Add(continueButton)
        End Sub
    End Class

    Friend Class SaveAsWindowDialog
        Inherits Window
        Public Enum ActionSave
            UserQuery
            File
            NotSave
            [Continue]
        End Enum

        Public Property Action As ActionSave

        Private Const WIDTH_BUTTON As Integer = 120

        Public Sub New(nameQuery As String)
            Background = New SolidColorBrush(SystemColors.ControlColor)
            WindowStartupLocation = WindowStartupLocation.Manual
            SizeToContent = SizeToContent.WidthAndHeight

            ResizeMode = ResizeMode.NoResize
            ShowInTaskbar = False
            Title = "Save Dialog"

            Action = ActionSave.NotSave

            Dim border As Border = New Border() With {
                .BorderThickness = New Thickness(0),
                .BorderBrush = Brushes.Gray
            }

            Dim root As StackPanel = New StackPanel() With {
                .Margin = New Thickness(10)
            }
            border.Child = root
            Content = border

            Dim message As TextBlock = New TextBlock() With {
                .Text = "Save changes to the [" & nameQuery & "]?",
                .HorizontalAlignment = HorizontalAlignment.Center
            }

            Dim bottomStack As StackPanel = New StackPanel() With {
                .Orientation = Orientation.Horizontal,
                .HorizontalAlignment = HorizontalAlignment.Center,
                .Margin = New Thickness(0, 10, 0, 0)
            }

            Dim buttonSaveFile As Button = New Button() With {
                .Width = WIDTH_BUTTON,
                .Content = "Save to file..."
            }
            AddHandler buttonSaveFile.Click, Sub()
                                                 Action = ActionSave.File
                                                 Close()

                                             End Sub

            Dim buttonSaveUserQuery As Button = New Button() With {
                .Width = WIDTH_BUTTON,
                .Content = "Save as User Query",
                .Margin = New Thickness(5, 0, 5, 0)
            }
            AddHandler buttonSaveUserQuery.Click, Sub()
                                                      Action = ActionSave.UserQuery
                                                      Close()

                                                  End Sub

            Dim buttonNotSave As Button = New Button() With {
                .Width = WIDTH_BUTTON,
                .Content = "Don't save",
                .Margin = New Thickness(0, 0, 5, 0)
            }
            AddHandler buttonNotSave.Click, Sub()
                                                Action = ActionSave.NotSave
                                                Close()

                                            End Sub


            Dim buttonCancel As Button = New Button() With {
                .Width = WIDTH_BUTTON,
                .Content = "Cancel"
            }
            AddHandler buttonCancel.Click, Sub()
                                               Action = ActionSave.[Continue]
                                               Close()

                                           End Sub

            root.Children.Add(message)
            root.Children.Add(bottomStack)

            bottomStack.Children.Add(buttonSaveFile)
            bottomStack.Children.Add(buttonSaveUserQuery)
            bottomStack.Children.Add(buttonNotSave)
            bottomStack.Children.Add(buttonCancel)
        End Sub
    End Class
End Namespace
