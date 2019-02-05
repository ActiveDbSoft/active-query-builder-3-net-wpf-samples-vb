'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.QueryView

Class MainWindow
       Public Sub New()
        InitializeComponent()

        AddHandler Loaded, AddressOf MainWindow_Loaded
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

        ' set syntax provider
        QBuilder.SyntaxProvider = New MSSQLSyntaxProvider()

        ' Fill metadata container from the XML file. (For demonstration purposes.)
        Try
            QBuilder.QueryColumnListOptions.UseCustomExpressionBuilder = (AffectedColumns.ConditionColumns & AffectedColumns.ExpressionColumn)
            QBuilder.MetadataLoadingOptions.OfflineMode = True
            QBuilder.MetadataContainer.ImportFromXML("Northwind.xml")
            QBuilder.InitializeDatabaseSchemaTree()

            QBuilder.SQL = "SELECT Orders.OrderID, Orders.CustomerID, Orders.OrderDate, [Order Details].ProductID," & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "[Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount" & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "  FROM Orders INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID" & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "  WHERE Orders.OrderID > 0 AND [Order Details].Discount > 0"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub QBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
        ' Text of SQL query has been updated by the query builder.
        SqlEditor.Text = QBuilder.FormattedSQL
    End Sub

    Private Sub SqlEditor_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            QBuilder.SQL = SqlEditor.Text
            ShowErrorBanner(DirectCast(sender, FrameworkElement), "")
        Catch ex As SQLParsingException
            ' Set caret to error position
            SqlEditor.SelectionStart = ex.ErrorPos.pos
            ' Report error
            ShowErrorBanner(DirectCast(sender, FrameworkElement), ex.Message)
        End Try
    End Sub

    Public Sub ShowErrorBanner(control As FrameworkElement, text As String)
        ' Show new banner if text is not empty
       ErrorBox.Message = text
    End Sub

    Private Sub QBuilder_OnCustomExpressionBuilder(querycolumnlistitem As QueryColumnListItem, conditionIndex As Integer, expression As String)
        Dim msg As MessageContainer = New MessageContainer(Me) With {
            .Title = "Edit " & (If(conditionIndex <> -1, "condition", "expression")),
            .TextContent = expression
        }
        If msg.ShowDialog() <> True Then
            Return
        End If

        ' Update the criteria list with new expression text.
        If conditionIndex > -1 Then
            ' it's one of condition columns
            querycolumnlistitem.ConditionStrings(conditionIndex) = msg.TextContent
        Else
            ' it's the Expression column
            querycolumnlistitem.ExpressionString = msg.TextContent
        End If
    End Sub

    Public Class MessageContainer
        Inherits Window
        Public Property TextContent() As String
            Get
                Return _textBox.Text
            End Get
            Set(value As String)
                _textBox.Text = value
            End Set
        End Property
        Private ReadOnly _textBox As TextBox

        Public Sub New(owner As Window)
            owner = owner
            WindowStartupLocation = WindowStartupLocation.CenterOwner
            Width = 400
            Height = 300
            ShowInTaskbar = False

            Dim grid As Grid = New Grid() With {
                .Margin = New Thickness(5)
            }

            grid.RowDefinitions.Add(New RowDefinition() With {
                .Height = New GridLength(1, GridUnitType.Star)
            })
            grid.RowDefinitions.Add(New RowDefinition() With {
                .Height = GridLength.Auto
            })

            _textBox = New TextBox() With {
                .VerticalScrollBarVisibility = ScrollBarVisibility.Visible,
                .TextWrapping = TextWrapping.Wrap
            }

            _textBox.SetValue(Grid.RowProperty, 0)
            grid.Children.Add(_textBox)

            Dim stack As StackPanel = New StackPanel() With {
                .HorizontalAlignment = HorizontalAlignment.Right,
                .Orientation = Controls.Orientation.Horizontal,
                .Margin = New Thickness(0, 10, 0, 0)
            }

            stack.SetValue(Grid.RowProperty, 1)
            grid.Children.Add(stack)

            Dim buttonOk As Button = New Button() With {
                .Height = 23,
                .Width = 74,
                .Content = "OK"
            }

            AddHandler buttonOk.Click, Sub() DialogResult = True

            Dim buttonCancel As Button = New Button() With {
                .Height = 23,
                .Width = 74,
                .Content = "Cancel",
                .Margin = New Thickness(5, 0, 0, 0)
            }

            AddHandler buttonCancel.Click, Sub() Close()

            stack.Children.Add(buttonOk)
            stack.Children.Add(buttonCancel)
            Content = grid
        End Sub
    End Class

    Private Sub SqlEditor_OnTextChanged(sender As Object, e As EventArgs)
        ErrorBox.Message = string.Empty
    End Sub
End Class
