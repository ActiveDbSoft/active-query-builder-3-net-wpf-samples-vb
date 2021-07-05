//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System
Imports System.ComponentModel
Imports ActiveQueryBuilder.View.QueryView
Imports ActiveQueryBuilder.View.WPF

Namespace QueryBuilderProperties
	''' <summary>
	''' Interaction logic for MiscellaneousPage.xaml
	''' </summary>
	<ToolboxItem(False)>
	Partial Public Class MiscellaneousPage
		Private ReadOnly _queryBuilder As QueryBuilder

		Public Property Modified() As Boolean

		Public Sub New(qb As QueryBuilder)
			Modified = False
			_queryBuilder = qb

			InitializeComponent()

			comboLinksStyle.Items.Add("Simple style")
			comboLinksStyle.Items.Add("MS Access style")
			comboLinksStyle.Items.Add("SQL Server Enterprise Manager style")

			If _queryBuilder.DesignPaneOptions.LinkStyle = LinkStyle.Simple Then
				comboLinksStyle.SelectedIndex = 0
			ElseIf _queryBuilder.DesignPaneOptions.LinkStyle = LinkStyle.MSAccess Then
				comboLinksStyle.SelectedIndex = 1
			ElseIf _queryBuilder.DesignPaneOptions.LinkStyle = LinkStyle.MSSQL Then
				comboLinksStyle.SelectedIndex = 2
			End If

			AddHandler comboLinksStyle.SelectionChanged, AddressOf Changed
		End Sub

		Public Sub New()
			Modified = False
			InitializeComponent()
		End Sub

		Public Sub ApplyChanges()
			If Modified Then
				Select Case comboLinksStyle.SelectedIndex
					Case 0
						_queryBuilder.DesignPaneOptions.LinkStyle = LinkStyle.Simple
					Case 2
						_queryBuilder.DesignPaneOptions.LinkStyle = LinkStyle.MSSQL
					Case Else
						_queryBuilder.DesignPaneOptions.LinkStyle = LinkStyle.MSAccess
				End Select
			End If
		End Sub

		Private Sub Changed(sender As Object, e As EventArgs)
			Modified = True
		End Sub
	End Class
End Namespace
