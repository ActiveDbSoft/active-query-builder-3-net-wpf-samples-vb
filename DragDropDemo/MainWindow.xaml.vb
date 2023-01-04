''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2023 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View

Class MainWindow
    Private _dragBoxFromMouseDown As Rect = Rect.Empty
    Private _errorPosition As Integer = -1
    Private _lastValidSql As String

    Public Sub New()

        InitializeComponent()
        AddHandler Loaded, AddressOf MainWindow_Loaded
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

        ' Fill query builder with demo data
        QueryBuilder1.SyntaxProvider = New MSSQLSyntaxProvider()
        QueryBuilder1.MetadataLoadingOptions.OfflineMode = True
        QueryBuilder1.MetadataContainer.ImportFromXML("Northwind.xml")
        QueryBuilder1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub ListBox1_OnMouseDown(sender As Object, e As MouseButtonEventArgs)
        ' Prepare drag'n'drop:
        Dim mousePosition = e.GetPosition(DirectCast(sender, FrameworkElement))

        If IsContainsItemAtPoint(mousePosition) Then

            Dim dragSize = New Windows.Size(SystemParameters.MinimumHorizontalDragDistance, SystemParameters.MinimumVerticalDragDistance)
            _dragBoxFromMouseDown = New Rect(New Windows.Point(mousePosition.X - (dragSize.Width / 2), mousePosition.Y - (dragSize.Height / 2)), dragSize)
        Else
            _dragBoxFromMouseDown = Rect.Empty
        End If
    End Sub

    Private Function IsContainsItemAtPoint(target As Windows.Point) As Boolean
        Dim answer As Boolean = False

        For i = 0 To ListBox1.Items.Count - 1
            Dim item As ListBoxItem = DirectCast(ListBox1.ItemContainerGenerator.ContainerFromIndex(i), ListBoxItem)
            If item Is Nothing Then
                Continue For
            End If

            Dim rect As Rect = New Rect(item.TranslatePoint(New Windows.Point(0, 0), ListBox1), New Windows.Size(item.ActualWidth, item.ActualHeight))

            If Not rect.Contains(target) Then
                Continue For
            End If

            answer = True
            Exit For
        Next

        Return answer
    End Function

    Private Sub ListBox1_OnMouseUp(sender As Object, e As MouseButtonEventArgs)
        _dragBoxFromMouseDown = Rect.Empty
    End Sub

    Private Sub ListBox1_OnMouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        ' Double click will add the object in automatic position:
        If ListBox1.SelectedIndex = -1 Then
            Return
        End If
        Dim objectName As String = DirectCast(ListBox1.SelectedItem, ListBoxItem).Content.ToString()
        QueryBuilder1.AddObjectToActiveUnionSubQuery(objectName)
    End Sub

    Private Sub ListBox1_OnMouseMove(sender As Object, e As MouseEventArgs)
        ' Do drag:
        Dim mousePosition = e.GetPosition(DirectCast(sender, FrameworkElement))
        If ListBox1.SelectedIndex = -1 Then
            Return
        End If

        If e.LeftButton <> MouseButtonState.Pressed Then
            Return
        End If

        If _dragBoxFromMouseDown = Rect.Empty OrElse _dragBoxFromMouseDown.Contains(mousePosition.X, mousePosition.Y) Then
            Return
        End If

        Dim objectName As String = DirectCast(ListBox1.SelectedItem, ListBoxItem).Content.ToString()
        Dim metadataObject As MetadataObject = QueryBuilder1.MetadataContainer.FindItem(Of MetadataObject)(objectName)

        If metadataObject Is Nothing Then
            Return
        End If

        Dim dragObject As MetadataDragObject = New MetadataDragObject()
        dragObject.MetadataDragged.Add(metadataObject)

        DragDrop.DoDragDrop(ListBox1, dragObject, DragDropEffects.Copy)
    End Sub

    Private Sub QueryBuilder1_OnSQLUpdated(sender As Object, e As EventArgs)
        If Not IsLoaded Then
            Return
        End If
        ' Handle the event raised by SQL Builder object that the text of SQL query is changed

        ' Hide error banner if any
        ErrorBox.Visibility = Visibility.Collapsed

        ' update the text box
        TextBox1.Text = QueryBuilder1.FormattedSQL
        _lastValidSql = QueryBuilder1.FormattedSQL
    End Sub

    Private Sub TextBox_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)

        Try
            ' Update the query builder with manually edited query text:
            QueryBuilder1.SQL = TextBox1.Text
            _lastValidSql = QueryBuilder1.FormattedSQL
            ' Hide error banner if any
            ErrorBox.Visibility = Visibility.Collapsed
        Catch ex As SQLParsingException
            ' Set caret to error position
            TextBox1.SelectionStart = ex.ErrorPos.pos

            ' Show banner with error text
            _errorPosition = ex.ErrorPos.pos
            ErrorBox.Show(ex.Message, QueryBuilder1.SyntaxProvider)
        End Try
    End Sub

    Private Sub TextBox1_OnTextChanged(sender As Object, e As TextChangedEventArgs)
        ErrorBox.Visibility = Visibility.Collapsed
    End Sub

    Private Sub ErrorBox_OnGoToErrorPosition(sender As Object, e As EventArgs)
        TextBox1.Focus()

        If _errorPosition <> -1 Then
            TextBox1.ScrollToLine(TextBox1.GetLineIndexFromCharacterIndex(_errorPosition))
            TextBox1.CaretIndex = _errorPosition
        End If
    End Sub

    Private Sub ErrorBox_OnRevertValidText(sender As Object, e As EventArgs)
        TextBox1.Text = _lastValidSql
        ErrorBox.Visibility = Visibility.Collapsed
        TextBox1.Focus()
    End Sub
End Class
