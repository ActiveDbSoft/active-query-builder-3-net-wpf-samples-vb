'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel
Imports System.Globalization
Imports System.Windows.Controls
Imports ActiveQueryBuilder.View.DatabaseSchemaView
Imports ActiveQueryBuilder.View.WPF

Namespace PropertiesForm
	''' <summary>
	''' Interaction logic for DatabaseSchemaViewPage.xaml
	''' </summary>
	<ToolboxItem(False)> _
	Public Partial Class DatabaseSchemaViewPage
		Private ReadOnly _queryBuilder As QueryBuilder
		Public Property Modified() As Boolean
			Get
				Return m_Modified
			End Get
			Set
				m_Modified = Value
			End Set
		End Property
		Private m_Modified As Boolean


		Public Sub New(queryBuilder As QueryBuilder)
			Modified = False
			_queryBuilder = queryBuilder

			InitializeComponent()

			cbGroupByServers.IsChecked = queryBuilder.MetadataStructure.Options.GroupByServers
			cbGroupByDatabases.IsChecked = queryBuilder.MetadataStructure.Options.GroupByDatabases
			cbGroupBySchemas.IsChecked = queryBuilder.MetadataStructure.Options.GroupBySchemas
			cbGroupByTypes.IsChecked = queryBuilder.MetadataStructure.Options.GroupByTypes
			cbShowFields.IsChecked = queryBuilder.MetadataStructure.Options.ShowFields

			cmbSortObjectsBy.Items.Add("Sort by Name")
			cmbSortObjectsBy.Items.Add("Sort by Type, Name")
			cmbSortObjectsBy.Items.Add("No sorting")
			cmbSortObjectsBy.SelectedIndex = CInt(queryBuilder.DatabaseSchemaViewOptions.SortingType)

			cmbDefaultExpandLevel.Text = queryBuilder.DatabaseSchemaViewOptions.DefaultExpandLevel.ToString(CultureInfo.InvariantCulture)

			AddHandler cbGroupByServers.Checked, AddressOf Changed
			AddHandler cbGroupByServers.Unchecked, AddressOf Changed
			AddHandler cbGroupByDatabases.Checked, AddressOf Changed
			AddHandler cbGroupByDatabases.Unchecked, AddressOf Changed
			AddHandler cbGroupBySchemas.Checked, AddressOf Changed
			AddHandler cbGroupBySchemas.Unchecked, AddressOf Changed
			AddHandler cbGroupByTypes.Checked, AddressOf Changed
			AddHandler cbGroupByTypes.Unchecked, AddressOf Changed
			AddHandler cbShowFields.Checked, AddressOf Changed
			AddHandler cbShowFields.Unchecked, AddressOf Changed
			AddHandler cmbSortObjectsBy.SelectionChanged, AddressOf Changed
			AddHandler cmbDefaultExpandLevel.SelectionChanged, AddressOf Changed
		End Sub

		Public Sub New()
			Modified = False
			InitializeComponent()
		End Sub

		Public Sub ApplyChanges()
			If Not Modified Then
				Return
			End If

			Dim metadataStructureOptions = _queryBuilder.MetadataStructure.Options
			metadataStructureOptions.BeginUpdate()

			Try
				metadataStructureOptions.GroupByServers = cbGroupByServers.IsChecked.HasValue AndAlso cbGroupByServers.IsChecked.Value
				metadataStructureOptions.GroupByDatabases = cbGroupByDatabases.IsChecked.HasValue AndAlso cbGroupByDatabases.IsChecked.Value
				metadataStructureOptions.GroupBySchemas = cbGroupBySchemas.IsChecked.HasValue AndAlso cbGroupBySchemas.IsChecked.Value
				metadataStructureOptions.GroupByTypes = cbGroupByTypes.IsChecked.HasValue AndAlso cbGroupByTypes.IsChecked.Value
				metadataStructureOptions.ShowFields = cbShowFields.IsChecked.HasValue AndAlso cbShowFields.IsChecked.Value
			Finally
				metadataStructureOptions.EndUpdate()
			End Try

			Dim databaseSchemaViewOptions = _queryBuilder.DatabaseSchemaViewOptions
			databaseSchemaViewOptions.BeginUpdate()

			Try
				_queryBuilder.DatabaseSchemaViewOptions.SortingType = CType(cmbSortObjectsBy.SelectedIndex, ObjectsSortingType)

				Dim defaultExpandLevel As Integer
				If Integer.TryParse(cmbDefaultExpandLevel.Text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, defaultExpandLevel) Then
					_queryBuilder.DatabaseSchemaViewOptions.DefaultExpandLevel = defaultExpandLevel
				End If
			Finally
				databaseSchemaViewOptions.EndUpdate()
			End Try
		End Sub

		Private Sub Changed(sender As Object, e As EventArgs)
			Modified = True
		End Sub

		Private Sub CmbDefaultExpandLevel_OnTextChanged(sender As Object, e As TextChangedEventArgs)
			Modified = True
		End Sub
	End Class
End Namespace
