'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Markup
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core

Namespace Common
    Public NotInheritable Class Helpers
        Private Sub New()
        End Sub
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
            Dim qpf As QueryParametersWindow = New QueryParametersWindow(command)
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

        Public Shared Function GetLocalizedText(text As String, lang As XmlLanguage) As String
            If String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            Dim properties As FieldInfo() = GetType(Constants).GetFields()

            Dim [property] As FieldInfo = properties.FirstOrDefault(Function(prop) prop.Name = text)

            If [property] Is Nothing Then
                Return text
            End If

            Dim constValue As String = [property].GetValue(Nothing).ToString()

            Return ActiveQueryBuilder.Core.Helpers.Localizer.GetString([property].Name, ActiveQueryBuilder.View.WPF.Helpers.ConvertLanguageFromNative(lang), constValue)
        End Function
    End Class

    Friend Class SaveExistQueryDialog
        Inherits WindowMessage
        Public Result As Nullable(Of Boolean)

        Private Const WidthButton As Integer = 90

        Public Sub New()
            Title = "Save dialog"
            Text = "Save changes?"
            ContetnAlignment = HorizontalAlignment.Center
            Result = Nothing

            Dim saveButton As Button = New Button() With {
                .Content = "Save",
                .Width = WidthButton,
                .IsDefault = True
            }
            AddHandler saveButton.Click, Sub()
                                             Result = True
                                             Close()

                                         End Sub

            Dim notSaveButton As Button = New Button() With {
                .Margin = New Thickness(5, 0, 5, 0),
                .Content = "Don't save",
                .Width = WidthButton,
                .IsCancel = True
            }
            AddHandler notSaveButton.Click, Sub()
                                                Result = False
                                                Close()

                                            End Sub

            Dim continueButton As Button = New Button() With {
                .Content = "Continue edit",
                .Width = WidthButton
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

        Private Const WidthButton As Integer = 120

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

            Dim messsage As TextBlock = New TextBlock() With {
                .Text = "Save changes to the [" & nameQuery & "]?",
                .HorizontalAlignment = HorizontalAlignment.Center
            }

            Dim bottomStack As StackPanel = New StackPanel() With {
                .Orientation = Orientation.Horizontal,
                .HorizontalAlignment = HorizontalAlignment.Center,
                .Margin = New Thickness(0, 10, 0, 0)
            }

            Dim buttonSaveFile As Button = New Button() With {
                .Width = WidthButton,
                .Content = "Save to file..."
            }
            AddHandler buttonSaveFile.Click, Sub()
                                                 Action = ActionSave.File
                                                 Close()

                                             End Sub

            Dim buttonSaveUserQuery As Button = New Button() With {
                .Width = WidthButton,
                .Content = "Save as User Query",
                .Margin = New Thickness(5, 0, 5, 0)
            }
            AddHandler buttonSaveUserQuery.Click, Sub()
                                                      Action = ActionSave.UserQuery
                                                      Close()

                                                  End Sub

            Dim buttonNotSave As Button = New Button() With {
                .Width = WidthButton,
                .Content = "Don't save",
                .Margin = New Thickness(0, 0, 5, 0)
            }
            AddHandler buttonNotSave.Click, Sub()
                                                Action = ActionSave.NotSave
                                                Close()

                                            End Sub


            Dim buttonCancel As Button = New Button() With {
                .Width = WidthButton,
                .Content = "Cancel"
            }
            AddHandler buttonCancel.Click, Sub()
                                               Action = ActionSave.[Continue]
                                               Close()

                                           End Sub

            root.Children.Add(messsage)
            root.Children.Add(bottomStack)

            bottomStack.Children.Add(buttonSaveFile)
            bottomStack.Children.Add(buttonSaveUserQuery)
            bottomStack.Children.Add(buttonNotSave)
            bottomStack.Children.Add(buttonCancel)
        End Sub
    End Class
End Namespace
