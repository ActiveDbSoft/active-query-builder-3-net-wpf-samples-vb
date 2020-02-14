'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows
Imports System.Windows.Controls
Imports ActiveQueryBuilder.Core

Namespace PropertiesForm.Tabs
	''' <summary>
	''' Interaction logic for MainQueryTab.xaml
	''' </summary>
	Public Partial Class MainQueryTab
		Public Property Options As SQLFormattingOptions

        Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(sqlFormattingOptions As SQLFormattingOptions)
			InitializeComponent()

			For Each value As KeywordFormat In GetType(KeywordFormat).GetEnumValues()
				ComboBoxKeywordCase.Items.Add(value)
			Next

			Options = sqlFormattingOptions
			LoadOptions()
		End Sub

		' Load options to form
		Public Sub LoadOptions()
			CheckBoxIndents.IsChecked = Options.DynamicIndents
			CheckBoxRightMargin.IsChecked = Options.DynamicRightMargin

			If Options.RightMargin > 0 Then
				CheckBoxEnableWordWrap.IsChecked = False
				UpDownCharacterInLine.Value = Options.RightMargin
			Else
				' no margin
				CheckBoxEnableWordWrap.IsChecked = False
				UpDownCharacterInLine.Value = 80
			End If
			CheckBoxWithinAnd.IsChecked = Options.ParenthesizeANDGroups
			CheckBoxSingleCondition.IsChecked = Options.ParenthesizeSingleCriterion

			ComboBoxKeywordCase.SelectedItem = Options.KeywordFormat
		End Sub

		Private Sub CheckBoxEnableWordWrap_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Options.RightMargin = If(Not CheckBoxEnableWordWrap.IsChecked = True, 0, UpDownCharacterInLine.Value)
		End Sub

		Private Sub CheckBoxIndents_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Options.DynamicIndents = CType((CheckBoxIndents.IsChecked = True), Boolean)
		End Sub

		Private Sub CheckBoxRightMargin_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Options.DynamicRightMargin = CType((CheckBoxRightMargin.IsChecked = True), Boolean)
		End Sub

		Private Sub CheckBoxWithinAnd_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Options.ParenthesizeANDGroups = CType((CheckBoxWithinAnd.IsChecked = True), Boolean)
		End Sub

		Private Sub CheckBoxSingleCondition_OnChanged(sender As Object, e As RoutedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Options.ParenthesizeSingleCriterion = CType((CheckBoxSingleCondition.IsChecked = True), Boolean)
		End Sub

		Private Sub ComboBoxKeywordCase_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
			If Not IsInitialized Then
				Return
			End If

			Options.KeywordFormat = CType(ComboBoxKeywordCase.SelectedItem, KeywordFormat)
		End Sub

		Private Sub UpDownCharacterInLine_OnValueChanged(sender As Object, e As EventArgs)
			If Not IsInitialized Then
				Return
			End If

			If CheckBoxEnableWordWrap.IsChecked = True Then
				Options.RightMargin = UpDownCharacterInLine.Value
			End If
		End Sub
	End Class
End Namespace
