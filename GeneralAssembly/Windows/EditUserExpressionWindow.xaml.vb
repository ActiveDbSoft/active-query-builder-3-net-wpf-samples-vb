//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports ActiveQueryBuilder.View.QueryView
Imports GeneralAssembly.Common

Namespace Windows
    ''' <summary>
    ''' Interaction logic for EditUserExpressionWindow.xaml
    ''' </summary>
    Partial Public Class EditUserExpressionWindow
        Private _queryView As IQueryView
        Private _selectedPredefinedCondition As UserExpressionVisualItem

        Private ReadOnly _source As New ObservableCollection(Of UserExpressionVisualItem)()

        Public Sub New()
            InitializeComponent()

            ListBoxUserExpressions.ItemsSource = _source

            'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not handle local variables named the same as class members well:
            For Each name_Renamed In System.Enum.GetNames(GetType(DbType))
                ComboboxDbTypes.Items.Add(New SelectableItem(name_Renamed))
            Next name_Renamed
        End Sub

        Public Sub Load(queryView As IQueryView)
            _source.Clear()
            _selectedPredefinedCondition = Nothing

            If _queryView Is Nothing Then
                _queryView = queryView
            End If

            If _queryView Is Nothing Then
                Return
            End If

            For Each expression In _queryView.UserPredefinedConditions
                _source.Add(New UserExpressionVisualItem(expression))
            Next expression
        End Sub

        Private Sub ButtonSaveForm_OnClick(sender As Object, e As RoutedEventArgs)
            SaveUserExpression()
        End Sub

        Private Function SaveUserExpression() As Boolean
            If _queryView Is Nothing Then
                Return False
            End If

            Try
                Dim token As Token = Nothing
                Dim result = _queryView.Query.SQLContext.ParseLogicalExpression(TextBoxExpression.Text, False, False, token)
                If result Is Nothing AndAlso token IsNot Nothing Then
                    Throw New SQLParsingException(String.Format(ActiveQueryBuilder.Core.Helpers.Localizer.GetString(NameOf(LocalizableConstantsUI.strInvalidCondition), WPF.Helpers.ConvertLanguageFromNative(Language), LocalizableConstantsUI.strInvalidCondition), TextBoxExpression.Text), token)
                End If
            Catch exception As Exception
                MessageBox.Show(exception.Message, "Invalid SQL", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.None)

                TextBoxExpression.Focus()
                Keyboard.Focus(TextBoxExpression)
                Return False
            End Try

            Dim listTypes = ComboboxDbTypes.Items.Where(Function(x) x.IsChecked).Select(Function(selectableItem) DirectCast(System.Enum.Parse(GetType(DbType), selectableItem.Content.ToString(), True), DbType)).ToList()

            Dim userExpression = New PredefinedCondition(TextBoxCaption.Text, listTypes, TextBoxExpression.Text, CheckBoxIsNeedEdit.IsChecked.HasValue AndAlso CheckBoxIsNeedEdit.IsChecked.Value = True)

            Dim index = -1
            If _selectedPredefinedCondition IsNot Nothing Then
                index = _queryView.UserPredefinedConditions.IndexOf(_selectedPredefinedCondition.PredefinedCondition)
                _queryView.UserPredefinedConditions.Remove(_selectedPredefinedCondition.PredefinedCondition)
            End If

            If _queryView.UserPredefinedConditions.Any(Function(x) String.Compare(x.Caption, TextBoxCaption.Text, StringComparison.InvariantCultureIgnoreCase) = 0) Then
                MessageBox.Show($"Condition with caption ""{TextBoxCaption.Text}"" already exist", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK)

                Keyboard.Focus(TextBoxCaption)

                Return False
            End If

            Try
                If index <> -1 Then
                    _queryView.UserPredefinedConditions.Insert(index, userExpression)
                Else
                    _queryView.UserPredefinedConditions.Add(userExpression)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End Try

            Load(Nothing)

            ResetForm()

            If index >= 0 Then
                ListBoxUserExpressions.SelectedIndex = index
            End If
            Return True
        End Function

        Private Sub RemoveSelectedUserExpression()
            Dim itemForRemove = ListBoxUserExpressions.SelectedItems.OfType(Of UserExpressionVisualItem)().ToList()

            For Each item In itemForRemove
                _queryView.UserPredefinedConditions.Remove(item.PredefinedCondition)
            Next item

            Load(Nothing)

            ResetForm()
        End Sub

        Private Sub ResetForm()
            _selectedPredefinedCondition = Nothing
            TextBoxCaption.Text = String.Empty
            TextBoxExpression.Text = String.Empty
            CheckBoxIsNeedEdit.IsChecked = False

            For Each selectableItem In ComboboxDbTypes.Items
                selectableItem.IsChecked = False
            Next selectableItem
        End Sub

        Private Sub ListBoxUserExpressions_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            UpdateForm()
        End Sub

        Private Sub UpdateForm()
            Dim enable = TypeOf ListBoxUserExpressions.SelectedItem Is UserExpressionVisualItem
            ButtonCopyCurrent.IsEnabled = enable
            ButtonDelete.IsEnabled = enable
            ButtonMoveUp.IsEnabled = enable
            ButtonMoveDown.IsEnabled = enable

            If Not enable Then
                Return
            End If

            _selectedPredefinedCondition = CType(ListBoxUserExpressions.SelectedItem, UserExpressionVisualItem)

            TextBoxCaption.Text = _selectedPredefinedCondition.Caption
            TextBoxExpression.Text = _selectedPredefinedCondition.Condition
            CheckBoxIsNeedEdit.IsChecked = _selectedPredefinedCondition.IsNeedEdit

            For Each item In
                _selectedPredefinedCondition.ShowOnlyForDbTypes.Select(
                    Function(type) ComboboxDbTypes.Items.First(Function(x) String.Equals(x.Content.ToString(),
                                                                                         type.ToString(),
                                                                                           StringComparison.InvariantCultureIgnoreCase)))
                item.IsChecked = True
            Next item

            ButtonSave.IsEnabled = False
        End Sub

        Private Sub ButtonAddNew_OnClick(sender As Object, e As RoutedEventArgs)
            ResetForm()
            ListBoxUserExpressions.SelectedItem = Nothing
            Keyboard.Focus(TextBoxCaption)
        End Sub

        Private Sub ButtonCopyCurrent_OnClick(sender As Object, e As RoutedEventArgs)
            _selectedPredefinedCondition = TryCast(ListBoxUserExpressions.SelectedItem, UserExpressionVisualItem)

            If _selectedPredefinedCondition Is Nothing Then
                Return
            End If

            'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not handle local variables named the same as class members well:
            Dim name_Renamed = _selectedPredefinedCondition.Caption

            Dim newName = ""

            For i = 0 To 999
                newName = $"{name_Renamed} Copy"
                If i > 0 Then
                    newName += $" ({i})"
                End If
                If _source.Any(Function(x) String.Compare(x.Caption, newName, StringComparison.InvariantCultureIgnoreCase) = 0) Then
                    Continue For
                End If

                Exit For
            Next i

            Dim newCopy = _selectedPredefinedCondition.Copy(newName)
            Dim index = _source.IndexOf(_selectedPredefinedCondition)

            _queryView.UserPredefinedConditions.Insert(index + 1, newCopy.PredefinedCondition)

            Load(Nothing)
            ListBoxUserExpressions.SelectedIndex = index + 1
        End Sub


        Private Sub ButtonDelete_OnClick(sender As Object, e As RoutedEventArgs)
            RemoveSelectedUserExpression()
        End Sub

        Private Sub ButtonMoveUp_OnClick(sender As Object, e As RoutedEventArgs)
            Dim selectedItem = TryCast(ListBoxUserExpressions.SelectedItem, UserExpressionVisualItem)

            If selectedItem Is Nothing Then
                Return
            End If

            Dim index = _source.IndexOf(selectedItem)

            If index - 1 < 0 Then
                Return
            End If

            ActiveQueryBuilder.Core.Helpers.IListMove(_queryView.UserPredefinedConditions, index, index - 1)

            Load(Nothing)

            ListBoxUserExpressions.SelectedIndex = index - 1
        End Sub

        Private Sub ButtonMoveDown_OnClick(sender As Object, e As RoutedEventArgs)
            Dim selectedItem = TryCast(ListBoxUserExpressions.SelectedItem, UserExpressionVisualItem)

            If selectedItem Is Nothing Then
                Return
            End If

            Dim index = _source.IndexOf(selectedItem)

            If index + 1 >= _source.Count Then
                Return
            End If

            ActiveQueryBuilder.Core.Helpers.IListMove(_queryView.UserPredefinedConditions, index, index + 1)

            Load(Nothing)

            ListBoxUserExpressions.SelectedIndex = index + 1
        End Sub

        Private Sub CheckChangingItem()
            If _selectedPredefinedCondition Is Nothing Then
                ButtonSave.IsEnabled = Not String.IsNullOrEmpty(TextBoxCaption.Text) AndAlso Not String.IsNullOrEmpty(TextBoxExpression.Text)

                Return
            End If

            Dim dbTypes = ComboboxDbTypes.Items.Where(Function(x) x.Content IsNot Nothing AndAlso x.IsChecked).Select(Function(x) DirectCast(System.Enum.Parse(GetType(DbType), x.Content.ToString(), True), DbType)).ToList()

            Dim changed = String.Compare(TextBoxCaption.Text, _selectedPredefinedCondition.Caption, StringComparison.InvariantCulture) <> 0 _
                OrElse String.Compare(TextBoxExpression.Text, _selectedPredefinedCondition.Condition, StringComparison.InvariantCulture) <> 0 _
                OrElse CheckBoxIsNeedEdit.IsChecked <> _selectedPredefinedCondition.IsNeedEdit _
                OrElse Not ((_selectedPredefinedCondition.ShowOnlyForDbTypes.Count = dbTypes.Count) AndAlso Not _selectedPredefinedCondition.ShowOnlyForDbTypes.Except(dbTypes).Any())

            ButtonSave.IsEnabled = changed.HasValue AndAlso changed.Value
        End Sub

        Private Sub TextBoxExpression_OnTextChanged(sender As Object, e As EventArgs)
            CheckChangingItem()
        End Sub

        Private Sub ComboboxDbTypes_OnItemCheckStateChanged(sender As Object, e As EventArgs)
            CheckChangingItem()
        End Sub

        Private Sub CheckBoxIsNeedEdit_OnCheckChanged(sender As Object, e As RoutedEventArgs)
            CheckChangingItem()
        End Sub

        Private Sub TextBoxCaption_OnTextChanged(sender As Object, e As TextChangedEventArgs)
            CheckChangingItem()
        End Sub

        Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub

        Protected Overrides Sub OnClosing(e As CancelEventArgs)
            If ButtonSave.IsEnabled Then
                Dim result = MessageBox.Show("Save changes to the current condition?", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Information)

                If Equals(result, MessageBoxResult.Yes) AndAlso Not SaveUserExpression() Then
                    e.Cancel = True
                End If
            End If

            MyBase.OnClosing(e)
        End Sub

        Private Sub ButtonRevert_OnClick(sender As Object, e As RoutedEventArgs)
            ResetForm()
            UpdateForm()
        End Sub
    End Class

    Public Class UserExpressionVisualItem
        Public Property ShowOnlyForDbTypes() As List(Of DbType)

        Public ReadOnly Property Caption() As String
        Public ReadOnly Property Condition() As String
        Public ReadOnly Property IsNeedEdit() As Boolean

        Public ReadOnly Property PredefinedCondition() As PredefinedCondition


        Public Sub New(predefinedCondition As PredefinedCondition)
            Me.PredefinedCondition = predefinedCondition
            ShowOnlyForDbTypes = New List(Of DbType)()

            Caption = predefinedCondition.Caption
            Condition = predefinedCondition.Expression
            IsNeedEdit = predefinedCondition.IsNeedEdit

            ShowOnlyForDbTypes.AddRange(predefinedCondition.ShowOnlyFor)
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Caption}"
        End Function

        Public Function Copy(newName As String) As UserExpressionVisualItem
            Return New UserExpressionVisualItem(New PredefinedCondition(newName, PredefinedCondition.ShowOnlyFor, PredefinedCondition.Expression, PredefinedCondition.IsNeedEdit))
        End Function
    End Class
End Namespace
