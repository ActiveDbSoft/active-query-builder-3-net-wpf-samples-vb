﻿'*******************************************************************'
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

Namespace PropertiesForm
	''' <summary>
	''' Interaction logic for SqlGenerationPage.xaml
	''' </summary>
	Public Partial Class SqlGenerationPage
		Inherits UserControl
		Private ReadOnly _generationOptions As SQLGenerationOptions
		Private ReadOnly _formattingOptions As SQLFormattingOptions

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(generationOptions As SQLGenerationOptions, formattingOptions As SQLFormattingOptions)

		    InitializeComponent()

			_generationOptions = generationOptions
			_formattingOptions = formattingOptions

			For Each value As ObjectPrefixSkipping In [Enum].GetValues(_generationOptions.ObjectPrefixSkipping.[GetType]())
				cbObjectPrefixSkipping.Items.Add(value)
			Next

			cbObjectPrefixSkipping.SelectedItem = _generationOptions.ObjectPrefixSkipping
			cbQuoteAllIdentifiers.IsChecked = _generationOptions.QuoteIdentifiers = IdentQuotation.All
		End Sub

		Private Sub cbObjectPrefixSkipping_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
			_generationOptions.ObjectPrefixSkipping = CType(cbObjectPrefixSkipping.SelectedItem, ObjectPrefixSkipping)
			_formattingOptions.ObjectPrefixSkipping = CType(cbObjectPrefixSkipping.SelectedItem, ObjectPrefixSkipping)
		End Sub

		Private Sub cbQuoteAllIdentifiers_Unchecked_1(sender As Object, e As RoutedEventArgs)
			_generationOptions.QuoteIdentifiers = IdentQuotation.IfNeed
			_formattingOptions.QuoteIdentifiers = IdentQuotation.IfNeed
		End Sub

		Private Sub cbQuoteAllIdentifiers_Checked_1(sender As Object, e As RoutedEventArgs)
			_generationOptions.QuoteIdentifiers = IdentQuotation.All
			_formattingOptions.QuoteIdentifiers = IdentQuotation.All
		End Sub
	End Class
End Namespace