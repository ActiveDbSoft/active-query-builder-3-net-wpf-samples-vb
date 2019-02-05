'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.ComponentModel
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF
Imports Microsoft.Win32

Namespace PropertiesForm
	''' <summary>
	''' Interaction logic for OfflineModePage.xaml
	''' </summary>
	<ToolboxItem(False)> _
	Public Partial Class OfflineModePage
		Private ReadOnly _metadataContainerCopy As MetadataContainer

		Private ReadOnly _openDialog As OpenFileDialog
		Private ReadOnly _saveDialog As SaveFileDialog
		Private ReadOnly _sqlContext As SQLContext

		Public Property Modified() As Boolean
			Get
				Return m_Modified
			End Get
			Set
				m_Modified = Value
			End Set
		End Property
		Private m_Modified As Boolean

		Public Sub New(sqlContext As SQLContext)
			_sqlContext = sqlContext
            _openDialog = New OpenFileDialog() With {
                .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*", _
                .Title = "Select XML file to load metadata from" _
            }

            _saveDialog = New SaveFileDialog() With {
                .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*", _
                .Title = "Select XML file to save metadata to" _
            }

			Modified = False

			_metadataContainerCopy = New MetadataContainer(_sqlContext)
			_metadataContainerCopy.Assign(_sqlContext.MetadataContainer)

			InitializeComponent()

			cbOfflineMode.IsChecked = _sqlContext.LoadingOptions.OfflineMode

			UpdateMode()

			AddHandler cbOfflineMode.Checked, AddressOf checkOfflineMode_CheckedChanged
			AddHandler cbOfflineMode.Unchecked, AddressOf checkOfflineMode_CheckedChanged
			AddHandler bEditMetadata.Click, AddressOf buttonEditMetadata_Click
			AddHandler bSaveToXML.Click, AddressOf buttonSaveToXML_Click
			AddHandler bLoadFromXML.Click, AddressOf buttonLoadFromXML_Click
			AddHandler bLoadMetadata.Click, AddressOf buttonLoadMetadata_Click
		End Sub

		Public Sub ApplyChanges()
			If Modified Then
				_sqlContext.LoadingOptions.OfflineMode = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value

				If _sqlContext.LoadingOptions.OfflineMode Then
					If _sqlContext.MetadataProvider IsNot Nothing Then
						_sqlContext.MetadataProvider.Disconnect()
					End If

					_sqlContext.MetadataContainer.Assign(_metadataContainerCopy)
				Else
					_sqlContext.MetadataContainer.Items.Clear()
				End If
			End If
		End Sub

		Private Sub checkOfflineMode_CheckedChanged(sender As Object, e As EventArgs)
			Modified = True
			UpdateMode()
		End Sub

		Private Sub buttonLoadMetadata_Click(sender As Object, e As EventArgs)
			_metadataContainerCopy.BeginUpdate()

			Try
				Dim f As New MetadataContainerLoadForm(_metadataContainerCopy, False)

				If f.ShowDialog() = True Then
					Modified = True
					cbOfflineMode.IsChecked = True

				End If
			Finally
				_metadataContainerCopy.EndUpdate()
			End Try
		End Sub

		Private Sub UpdateMode()
			' lMetadataObjectCount.Font = new Font(lMetadataObjectCount.Font, (cbOfflineMode.IsChecked.HasValue && cbOfflineMode.IsChecked.Value) ? FontStyle.Bold : FontStyle.Regular);
			bLoadFromXML.IsEnabled = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value
			bSaveToXML.IsEnabled = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value
			bEditMetadata.IsEnabled = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value

			UpdateMetadataStats()
		End Sub

		Private Sub UpdateMetadataStats()
			Dim metadataObjects As List(Of MetadataObject) = _metadataContainerCopy.Items.GetItemsRecursive(Of MetadataObject)(MetadataType.Objects)
			Dim t As Integer = 0, v As Integer = 0, p As Integer = 0, s As Integer = 0

			For i As Integer = 0 To metadataObjects.Count - 1
				Dim mo As MetadataObject = metadataObjects(i)

				If mo.Type = MetadataType.Table Then
					t += 1
				ElseIf mo.Type = MetadataType.View Then
					v += 1
				ElseIf mo.Type = MetadataType.Procedure Then
					p += 1
				ElseIf mo.Type = MetadataType.Synonym Then
					s += 1
				End If
			Next

			Dim tmp = "Loaded Metadata: {0} tables, {1} views, {2} procedures, {3} synonyms."
			lMetadataObjectCount.Text = String.Format(tmp, t, v, p, s)
		End Sub

		Private Sub buttonLoadFromXML_Click(sender As Object, e As EventArgs)
			If _openDialog.ShowDialog() <> True Then
				Return
			End If

			_metadataContainerCopy.ImportFromXML(_openDialog.FileName)
			Modified = True
			UpdateMetadataStats()
		End Sub

		Private Sub buttonSaveToXML_Click(sender As Object, e As EventArgs)
			If _saveDialog.ShowDialog() = True Then
				_metadataContainerCopy.ExportToXML(_saveDialog.FileName)
			End If
		End Sub

		Private Sub buttonEditMetadata_Click(sender As Object, e As EventArgs)
			If QueryBuilder.EditMetadataContainer(_metadataContainerCopy, _sqlContext.MetadataStructure, _sqlContext.LoadingOptions) Then
				Modified = True
			End If
		End Sub
	End Class
End Namespace
