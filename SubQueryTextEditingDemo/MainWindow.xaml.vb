'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core

Namespace SubQueryTextEditingDemo
	Friend Enum ModeEditor
		Entire
		SubQuery
		Expression
	End Enum
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Public Partial Class MainWindow
		Private _popupMessage As Popup
		Private ReadOnly _timerCleanMessage As Timer
		Private _mode As ModeEditor

		Public Sub New()
			InitializeComponent()
			_timerCleanMessage = New Timer(AddressOf ClosePopup)

			AddHandler Loaded, AddressOf MainWindow_Loaded
			_mode = ModeEditor.Entire
		End Sub

		Private Sub ClosePopup(state As Object)
			If _popupMessage Is Nothing Then
				Return
			End If

			Dispatcher.BeginInvoke(DirectCast(Sub() 
			_popupMessage.IsOpen = False
			_popupMessage = Nothing

End Sub, Action))
		End Sub

		Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
			RemoveHandler Loaded, AddressOf MainWindow_Loaded

			Builder.SyntaxProvider = New MSSQLSyntaxProvider()
			Builder.MetadataProvider = New MSSQLMetadataProvider()

			Builder.MetadataContainer.LoadingOptions.OfflineMode = True
			Builder.MetadataContainer.ImportFromXML("Northwind.xml")

			Builder.DatabaseSchemaViewOptions.DefaultExpandLevel = 2
			Builder.InitializeDatabaseSchemaTree()

			TextEditor.QueryProvider = Builder

			Builder.SQL = "Select * From Customers"

			Breadcrumb.QueryView = Builder.QueryView

			AddHandler Builder.ActiveUnionSubQueryChanging, AddressOf Builder_ActiveUnionSubQueryChanging
			AddHandler Builder.ActiveUnionSubQueryChanged, AddressOf Builder_ActiveUnionSubQueryChanged

			AddHandler Breadcrumb.SizeChanged, AddressOf Breadcrumb_SizeChanged
		End Sub

		Private Sub Breadcrumb_SizeChanged(sender As Object, e As SizeChangedEventArgs)
			BottomGrid.InvalidateVisual()
		End Sub

		Private Function CheckSql() As SQLParsingException
			Try
				Dim sql as String = TextEditor.Text.Trim()

				Select Case _mode
					Case ModeEditor.Entire
						If Not String.IsNullOrEmpty(sql) Then
							Builder.SQLContext.ParseSelect(sql)
						End If
						Exit Select
					Case ModeEditor.SubQuery
						Builder.SQLContext.ParseSelect(sql)
						Exit Select
					Case ModeEditor.Expression
						Builder.SQLContext.ParseSelect(sql)
						Exit Select
					Case Else
						Throw New ArgumentOutOfRangeException()
				End Select

				Return Nothing
			Catch ex As SQLParsingException
				Return ex
			End Try
		End Function

		Private Sub Builder_ActiveUnionSubQueryChanged(sender As Object, e As EventArgs)
			ApplyText()
		End Sub

		Private Sub ApplyText()
			Dim sqlFormattingOptions As SQLFormattingOptions = Builder.SQLFormattingOptions

			If TextEditor Is Nothing Then
				Return
			End If

			Select Case _mode
				Case ModeEditor.Entire
					TextEditor.Text = Builder.FormattedSQL
					Exit Select
				Case ModeEditor.SubQuery
					If Builder.ActiveUnionSubQuery Is Nothing Then
						Exit Select
					End If
					Dim subQuery as SubQuery = Builder.ActiveUnionSubQuery.ParentSubQuery
					TextEditor.Text = FormattedSQLBuilder.GetSQL(subQuery, sqlFormattingOptions)
					Exit Select
				Case ModeEditor.Expression
					If Builder.ActiveUnionSubQuery Is Nothing Then
						Exit Select
					End If
					Dim unionSubQuery as UnionSubQuery = Builder.ActiveUnionSubQuery
					TextEditor.Text = FormattedSQLBuilder.GetSQL(unionSubQuery, sqlFormattingOptions)
					Exit Select
				Case Else
					Throw New ArgumentOutOfRangeException()
			End Select
		End Sub

		Private Sub Builder_ActiveUnionSubQueryChanging(sender As Object, e As ActiveQueryBuilder.View.SubQueryChangingEventArgs)
			Dim exception As SQLParsingException = CheckSql()

			If exception Is Nothing Then
				Return
			End If

			e.Abort = True

			ShowErrorBanner(TextEditor, exception)
		End Sub

		Private Sub Builder_OnSQLUpdated(sender As Object, e As EventArgs)
			ApplyText()
		End Sub

		Private Sub ShowErrorBanner(element As UIElement, exception As Exception)
			If Not (TypeOf exception Is SQLParsingException) Then
				Return
			End If

			If _popupMessage IsNot Nothing Then
				_popupMessage.IsOpen = False
				_popupMessage = Nothing
			End If

			Dim label as TextBlock= New TextBlock() With {
				.Text = exception.Message, 
            .Background = Brushes.LightPink, 
            .Padding = New Thickness(10)
			}

			Dim textContent as Border = New Border() With {
				.Child = label,
				.BorderBrush = Brushes.Black,
				.BorderThickness = New Thickness(1)
			}

			textContent.Measure(New Size(Double.MaxValue, Double.MaxValue))
            Dim rect As Rect = New Rect(New Point(0, 0), textContent.DesiredSize) 
			textContent.Arrange(rect)

			_popupMessage = New Popup() With {
				.PlacementTarget = element,
				.Placement = PlacementMode.Left,
				.Child = textContent
			}

			_popupMessage.HorizontalOffset -= textContent.ActualWidth + 5
			_popupMessage.VerticalOffset += 5

			_popupMessage.IsOpen = True
			_timerCleanMessage.Change(3000, Timeout.Infinite)
		End Sub

		Private Sub TextEditor_OnGotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
			If _popupMessage Is Nothing Then
				Return
			End If

			_popupMessage.IsOpen = False
			_popupMessage = Nothing
		End Sub

		Private Sub ToggleButton_OnChecked(sender As Object, e As RoutedEventArgs)
			Try
				PopupSwitch.IsOpen = False

				If Equals(sender, RadioButtonEntire) Then
					_mode = ModeEditor.Entire
					Return
				End If

				If Equals(sender, RadioButtonSubQuery) Then
					_mode = ModeEditor.SubQuery
					Return
				End If

				_mode = ModeEditor.Expression
			Finally
				ApplyText()
			End Try
		End Sub

		Private Sub ButtonSwitch_OnClick(sender As Object, e As RoutedEventArgs)
			PopupSwitch.IsOpen = True
		End Sub

		Private Sub TextEditor_OnPreviewLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
			Dim text as String = TextEditor.Text.Trim()

			Dim isSuccess As Boolean = True

			Builder.BeginUpdate()

			Try
				If Not String.IsNullOrEmpty(text) Then
					Builder.SQLContext.ParseSelect(text)
				End If

				Select Case _mode
					Case ModeEditor.Entire
						Builder.SQL = text
						Exit Select
					Case ModeEditor.SubQuery
						Dim subQuery = Builder.ActiveUnionSubQuery.ParentSubQuery
						subQuery.SQL = text
						Exit Select
					Case ModeEditor.Expression
						Dim unionSubQuery = Builder.ActiveUnionSubQuery
						unionSubQuery.SQL = text
						Exit Select
					Case Else
						Throw New ArgumentOutOfRangeException()
				End Select
			Catch exception As Exception
				isSuccess = False
				ShowErrorBanner(TextEditor, exception)
			Finally
				Builder.EndUpdate()
			End Try

			e.Handled = Not isSuccess
		End Sub
	End Class
End Namespace
