﻿'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel
Imports ActiveQueryBuilder.Core

Namespace PropertiesForm
	''' <summary>
	''' Interaction logic for GeneralPage.xaml
	''' </summary>
	<ToolboxItem(False)> _
	Public Partial Class GeneralPage
		Private ReadOnly _sqlFormattingOptions As SQLFormattingOptions

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(sqlFormattingOptions As SQLFormattingOptions)
			_sqlFormattingOptions = sqlFormattingOptions

			InitializeComponent()

			cbWordWrap.IsChecked = (_sqlFormattingOptions.RightMargin <> 0)
			updownRightMargin.IsEnabled = cbWordWrap.IsChecked.Value

			updownRightMargin.Value = If(_sqlFormattingOptions.RightMargin = 0, 80, _sqlFormattingOptions.RightMargin)

			comboKeywordsCasing.Items.Add("Capitalized")
			comboKeywordsCasing.Items.Add("Uppercase")
			comboKeywordsCasing.Items.Add("Lowercase")

			comboKeywordsCasing.SelectedIndex = CInt(_sqlFormattingOptions.KeywordFormat)

			AddHandler cbWordWrap.Checked, AddressOf checkWordWrap_CheckedChanged
			AddHandler cbWordWrap.Unchecked, AddressOf checkWordWrap_CheckedChanged
			AddHandler updownRightMargin.ValueChanged, AddressOf updownRightMargin_ValueChanged
			AddHandler comboKeywordsCasing.SelectionChanged, AddressOf comboKeywordsCasing_SelectedIndexChanged
		End Sub

		Private Sub comboKeywordsCasing_SelectedIndexChanged(sender As Object, e As EventArgs)
			ApplyChanges()
		End Sub

		Private Sub checkWordWrap_CheckedChanged(sender As Object, e As EventArgs)
			updownRightMargin.IsEnabled = cbWordWrap.IsChecked.HasValue AndAlso cbWordWrap.IsChecked.Value
			ApplyChanges()
		End Sub

		Private Sub updownRightMargin_ValueChanged(sender As Object, e As EventArgs)
			ApplyChanges()
		End Sub

		Public Sub ApplyChanges()
			If cbWordWrap.IsChecked.HasValue AndAlso cbWordWrap.IsChecked.Value Then
				_sqlFormattingOptions.RightMargin = updownRightMargin.Value
			Else
				_sqlFormattingOptions.RightMargin = 0
			End If

			_sqlFormattingOptions.KeywordFormat = CType(comboKeywordsCasing.SelectedIndex, KeywordFormat)
		End Sub
	End Class
End Namespace
