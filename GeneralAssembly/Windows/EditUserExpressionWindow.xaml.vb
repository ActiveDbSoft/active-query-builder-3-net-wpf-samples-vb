'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports ActiveQueryBuilder.View.QueryView
Imports Common

Namespace Windows
	''' <summary>
	''' Interaction logic for EditUserExpressionWindow.xaml
	''' </summary>
	Partial Public Class EditUserExpressionWindow
		Private _queryView As IQueryView
		Private _editingUserExpression As UserExpressionVisualItem

		Public Sub New()
			InitializeComponent()

'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not handle local variables named the same as class members well:
			For Each name_Renamed In System.Enum.GetNames(GetType(DbType))
				ComboboxDbTypes.Items.Add(New SelectableItem(name_Renamed))
			Next name_Renamed
		End Sub

		Public Sub Load(queryView As IQueryView)
			ListBoxUserExpressions.Items.Clear()
			_editingUserExpression = Nothing

			If _queryView Is Nothing Then
				_queryView = queryView
			End If

			If _queryView Is Nothing Then
				Return
			End If

			For Each expression In _queryView.UserPredefinedConditions
				ListBoxUserExpressions.Items.Add(New UserExpressionVisualItem(expression))
			Next expression
		End Sub

		Private Sub ClearFormButton_OnClick(sender As Object, e As RoutedEventArgs)
			ResetForm()
		End Sub

		Private Sub AddFormButton_OnClick(sender As Object, e As RoutedEventArgs)
			SaveUserExpression()
		End Sub

		Private Sub RemoveSelectedButton_OnClick(sender As Object, e As RoutedEventArgs)
			RemoveSelectedUserExpression()
		End Sub

		Private Sub SaveUserExpression()
			If _queryView Is Nothing Then
				Return
			End If

			Dim listTypes = ComboboxDbTypes.Items.Where(Function(x) x.IsChecked).Select(Function(selectableItem) DirectCast(System.Enum.Parse(GetType(DbType), selectableItem.Content.ToString(), True), DbType)).ToList()

			Dim userExpression = New PredefinedCondition(TextBoxCaption.Text, listTypes, TextBoxExpression.Text, CType((CheckBoxIsNeedEdit.IsChecked = True), Boolean))

			If _editingUserExpression IsNot Nothing Then
				_queryView.UserPredefinedConditions.Remove(_editingUserExpression.ConditionExpression)
			End If

			_queryView.UserPredefinedConditions.Add(userExpression)

			Load(Nothing)

			ResetForm()
		End Sub

		Private Sub RemoveSelectedUserExpression()
			Dim itemForRemove = ListBoxUserExpressions.SelectedItems.OfType(Of UserExpressionVisualItem)().ToList()

			For Each item In itemForRemove
				_queryView.UserPredefinedConditions.Remove(item.ConditionExpression)
			Next item

			Load(Nothing)
		End Sub

		Private Sub ResetForm()
			_editingUserExpression = Nothing
			TextBoxCaption.Text = String.Empty
			TextBoxExpression.Text = String.Empty
			CheckBoxIsNeedEdit.IsChecked = False
			For Each selectableItem In ComboboxDbTypes.Items
				selectableItem.IsChecked = False
			Next selectableItem
		End Sub

		Private Sub EditExpressionButton_OnClick(sender As Object, e As RoutedEventArgs)
			If ListBoxUserExpressions.SelectedItems.Count <> 1 Then
				Return
			End If

			_editingUserExpression = TryCast(ListBoxUserExpressions.SelectedItem, UserExpressionVisualItem)

			If _editingUserExpression Is Nothing Then
				Return
			End If

			TextBoxCaption.Text = _editingUserExpression.Caption
			TextBoxExpression.Text = _editingUserExpression.Expression
			CheckBoxIsNeedEdit.IsChecked = _editingUserExpression.IsNeedEdit

			For Each item In _editingUserExpression.ShowOnlyForDbTypes.Select(Function(type) ComboboxDbTypes.Items.First(Function(x) String.Equals(x.Content.ToString(), type.ToString(), StringComparison.InvariantCultureIgnoreCase)))
				item.IsChecked = True
			Next item
		End Sub
	End Class

	Public Class UserExpressionVisualItem
		Public Property ShowOnlyForDbTypes() As List(Of DbType)

		Public Property Caption() As String
		Public Property Expression() As String
		Public Property IsNeedEdit() As Boolean

        Public Property ConditionExpression As PredefinedCondition

        Public Sub New(conditionExpression As PredefinedCondition)
			Me.ConditionExpression = conditionExpression
			ShowOnlyForDbTypes = New List(Of DbType)()

			Caption = conditionExpression.Caption
			Expression = conditionExpression.Expression
			IsNeedEdit = conditionExpression.IsNeedEdit

			ShowOnlyForDbTypes.AddRange(conditionExpression.ShowOnlyFor)
		End Sub

		Public Overrides Function ToString() As String
			Dim types = If(ShowOnlyForDbTypes.Count = 0, "For all types", "For types: ")
			For Each dbType In ShowOnlyForDbTypes
				If ShowOnlyForDbTypes.IndexOf(dbType) <> 0 Then
					types &= ", "
				End If

				types &= dbType
			Next dbType

			Return $"{Caption}, [{Expression}], Is need edit: {IsNeedEdit}, {types}"
		End Function
	End Class
End Namespace
