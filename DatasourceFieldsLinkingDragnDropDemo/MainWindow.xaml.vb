'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports ActiveQueryBuilder.Core

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
			QBuilder.MetadataLoadingOptions.OfflineMode = True
			QBuilder.MetadataContainer.ImportFromXML("Northwind.xml")
			QBuilder.InitializeDatabaseSchemaTree()

			QBuilder.SQL = "SELECT Orders.OrderID, Orders.CustomerID, Orders.OrderDate, [Order Details].ProductID," & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "[Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount" & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "  FROM Orders INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID" & vbCr & vbLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "  WHERE Orders.OrderID > 0 AND [Order Details].Discount > 0"
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try
	End Sub

	Private Sub QBuilder_OnSQLUpdated(sender As Object, e As EventArgs)
		' Update the text of SQL query when it's changed in the query builder.
		SqlEditor.Text = QBuilder.FormattedSQL
	End Sub

	Private Sub SqlEditor_OnLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
		Try
			' Feed the text from text editor to the query builder when user exits the editor.
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
		' Display error banner if passed text is not empty.
		ErrorBox.Message = text
	End Sub
    
	''' <summary>
	''' The handler checks if the dragged field is a part of the primary key and denies dragging if it's not the case.
	''' </summary>
	Private Sub QBuilder_OnBeforeDatasourceFieldDrag(datasource As DataSource, field As MetadataField, ByRef abort As Boolean)
		If CheckBoxBeforeDsFieldDrag.IsChecked <> True Then
			Return
		End If

		' deny dragging if a field isn't a part of the primary key
		If Not field.PrimaryKey Then
			BoxLogEvents.Text = "OnBeforeDatasourceFieldDrag for field """ & Convert.ToString(field.Name) & " "" deny" & Environment.NewLine & BoxLogEvents.Text
			abort = True
			Return
		End If

		BoxLogEvents.Text = "OnBeforeDatasourceFieldDrag for field """ & Convert.ToString(field.Name) & " "" allow" & Environment.NewLine & BoxLogEvents.Text
	End Sub

	''' <summary>
	''' Checking the feasibility of creating a link between the given fields.
	''' </summary>
	Private Sub QBuilder_OnLinkDragOver(fromDataDource As DataSource, fromField As MetadataField, toDataSource As DataSource, toField As MetadataField, correspondingmetadataforeignkey As MetadataForeignKey, ByRef abort As Boolean)
		If CheckBoxLinkDragOver.IsChecked <> True Then
			Return
		End If

		' Allow creation of links between fields of the same data type.
		If fromField.FieldType = toField.FieldType Then
			BoxLogEvents.Text = "OnLinkDragOver from field """ & Convert.ToString(fromField.Name) & """ to field """ & Convert.ToString(toField.Name) & """ allow" & Environment.NewLine & BoxLogEvents.Text
			Return
		End If

		BoxLogEvents.Text = "OnLinkDragOver from field """ & Convert.ToString(fromField.Name) & """ to field """ & Convert.ToString(toField.Name) & """ deny" & Environment.NewLine & BoxLogEvents.Text

		abort = True
	End Sub

	Private Sub BoxLogEvents_OnTextChanged(sender As Object, e As TextChangedEventArgs)
	    ErrorBox.Message = string.Empty
	End Sub
End Class
