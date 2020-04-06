'*******************************************************************'
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
Imports ActiveQueryBuilder.View.DatabaseSchemaView
Imports ActiveQueryBuilder.View.WPF
Imports Common

Namespace QueryBuilderProperties
	''' <summary>
	''' Interaction logic for DatabaseSchemaViewPage.xaml
	''' </summary>
	<ToolboxItem(False)>
	Partial Friend Class DatabaseSchemaViewPage
		Private _expandMetadataType As MetadataType
		Private ReadOnly _queryBuilder As QueryBuilder
		Public Property Modified() As Boolean


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

			_expandMetadataType = queryBuilder.DatabaseSchemaView.Options.DefaultExpandMetadataType
			FillComboBox(GetType(MetadataType))
			SetExpandType(queryBuilder.DatabaseSchemaView.Options.DefaultExpandMetadataType)
		End Sub
		Private Sub FillComboBox(enumType As Type)
			Dim flags = GetFlagsFromType(enumType)
			For Each flag In flags
				cmbDefaultExpandLevel.Items.Add(New SelectableItem(flag))
			Next flag
		End Sub

		Private Sub SetExpandType(value As Object)
			cmbDefaultExpandLevel.ClearCheckedItems()
			Dim decomposed = DecomposeEnum(value)
			For i As Integer = 0 To cmbDefaultExpandLevel.Items.Count - 1
				If decomposed.Contains(CInt(Math.Truncate(cmbDefaultExpandLevel.Items(i).Content))) Then
					cmbDefaultExpandLevel.SetItemChecked(i, True)
				End If
			Next i
		End Sub

		Private Function DecomposeEnum(value As Object) As List(Of Integer)
			' decomposite enum by degrees of 2
			Dim binary = Convert.ToString(DirectCast(value, Integer), 2).Reverse().ToList()
			Dim result = New List(Of Integer)()
			For i As Integer = 0 To binary.Count - 1
				If binary(i) = "1"c Then
					result.Add(CInt(Math.Truncate(Math.Pow(2, i))))
				End If
			Next i

			Return result
		End Function

        Private Function GetFlagsFromType(enumType As Type) As List(Of [Enum])
            Dim values As Array = [Enum].GetValues(enumType)
            Dim result = New List(Of [Enum])()
            For Each value As Object In values
                ' filter unity items
                If IsDegreeOf2(CInt(value)) Then
                    result.Add(DirectCast(value, [Enum]))
                End If
            Next

            Return result
        End Function

        Private Function IsDegreeOf2(n As Integer) As Boolean
            Return n <> 0 AndAlso (n And (n - 1)) = 0
        End Function

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

				databaseSchemaViewOptions.DefaultExpandMetadataType = GetExpandType()
			Finally
				databaseSchemaViewOptions.EndUpdate()
			End Try
		End Sub

		Private Function GetExpandType() As MetadataType
			Dim intValue = CInt(_expandMetadataType)

			For i As Integer = 0 To cmbDefaultExpandLevel.Items.Count - 1
				If cmbDefaultExpandLevel.IsItemChecked(i) Then
					intValue = intValue Or CInt(Math.Truncate(cmbDefaultExpandLevel.Items(i).Content))
				Else
					intValue = intValue And Not CInt(Math.Truncate(cmbDefaultExpandLevel.Items(i).Content))
				End If
			Next i

			Return CType(intValue, MetadataType)
		End Function

		Private Sub Changed(sender As Object, e As EventArgs)
			Modified = True
		End Sub

		Private Sub CmbDefaultExpandLevel_OnItemCheckStateChanged(sender As Object, e As EventArgs)
			_expandMetadataType = GetExpandType()
			Changed(Me, EventArgs.Empty)
		End Sub
	End Class
End Namespace
