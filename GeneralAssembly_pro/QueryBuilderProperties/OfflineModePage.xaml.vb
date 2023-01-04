''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2023 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.WPF
Imports Microsoft.Win32

Namespace QueryBuilderProperties
    ''' <summary>
    ''' Interaction logic for OfflineModePage.xaml
    ''' </summary>
    <ToolboxItem(False)>
    Partial Public Class OfflineModePage
        Private ReadOnly _sqlContext As SQLContext
        Private ReadOnly _sqlContextCopy As SQLContext

        Private ReadOnly _openDialog As OpenFileDialog
        Private ReadOnly _saveDialog As SaveFileDialog

        Public Property Modified() As Boolean

        Public Sub New(context As SQLContext)
            _sqlContext = context
            _sqlContextCopy = New SQLContext()
            _sqlContextCopy.Assign(context)

            _openDialog = New OpenFileDialog With {
                .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
                .Title = "Select XML file to load metadata from"
            }

            _saveDialog = New SaveFileDialog With {
                .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
                .Title = "Select XML file to save metadata to"
            }

            'Modified = false;
            '_queryBuilder = queryBuilder;
            '_syntaxProvider = syntaxProvider;

            '_metadataContainerCopy = new MetadataContainer(queryBuilder.SQLContext);
            '_metadataContainerCopy.Assign(_queryBuilder.MetadataContainer);

            InitializeComponent()

            cbOfflineMode.IsChecked= _sqlContextCopy.LoadingOptions.OfflineMode

            UpdateMode()

            AddHandler cbOfflineMode.Checked, AddressOf checkOfflineMode_CheckedChanged
            AddHandler cbOfflineMode.Unchecked, AddressOf checkOfflineMode_CheckedChanged
            AddHandler bEditMetadata.Click, AddressOf buttonEditMetadata_Click
            AddHandler bSaveToXML.Click, AddressOf buttonSaveToXML_Click
            AddHandler bLoadFromXML.Click, AddressOf buttonLoadFromXML_Click
        End Sub

        Public Sub ApplyChanges()
            If Modified Then
                _sqlContextCopy.LoadingOptions.OfflineMode = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value

                If _sqlContextCopy.LoadingOptions.OfflineMode Then
                    _sqlContextCopy.MetadataProvider?.Disconnect()

                    _sqlContext.Assign(_sqlContextCopy)
                Else
                    _sqlContext.MetadataContainer.Items.Clear()
                End If
            End If
        End Sub

        Private Sub checkOfflineMode_CheckedChanged(sender As Object, e As EventArgs)
            Modified = True
            UpdateMode()
        End Sub

        Private Sub UpdateMode()
           ' lMetadataObjectCount.Font = new Font(lMetadataObjectCount.Font, (cbOfflineMode.IsChecked.HasValue && cbOfflineMode.IsChecked.Value) ? FontStyle.Bold : FontStyle.Regular);
            bLoadFromXML.IsEnabled = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value
            bSaveToXML.IsEnabled = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value
            bEditMetadata.IsEnabled = cbOfflineMode.IsChecked.HasValue AndAlso cbOfflineMode.IsChecked.Value

            UpdateMetadataStats()
        End Sub

        Private Sub UpdateMetadataStats()
            Dim metadataObjects As List(Of MetadataObject) = _sqlContextCopy.MetadataContainer.Items.GetItemsRecursive(Of MetadataObject)(MetadataType.Objects)
            Dim t As Integer = 0, v As Integer = 0, p As Integer = 0, s As Integer = 0

            For i = 0 To metadataObjects.Count - 1
                Dim mo As MetadataObject = metadataObjects(i)

                Select Case mo.Type
                    Case MetadataType.Table
                        t += 1
                    Case MetadataType.View
                        v += 1
                    Case MetadataType.Procedure
                        p += 1
                    Case MetadataType.Synonym
                        s += 1
                End Select
            Next i

            Dim tmp = "Loaded Metadata: {0} tables, {1} views, {2} procedures, {3} synonyms."
            lMetadataObjectCount.Text = String.Format(tmp, t, v, p, s)
        End Sub

        Private Sub buttonLoadFromXML_Click(sender As Object, e As EventArgs)
            If _openDialog.ShowDialog().Equals(True) Then
                _sqlContextCopy.MetadataContainer.ImportFromXML(_openDialog.FileName)
                Modified = True
                UpdateMetadataStats()
            End If
        End Sub

        Private Sub buttonSaveToXML_Click(sender As Object, e As EventArgs)
            If _saveDialog.ShowDialog().Equals(True) Then
                _sqlContextCopy.MetadataContainer.ExportToXML(_saveDialog.FileName)
            End If
        End Sub

        Private Sub buttonEditMetadata_Click(sender As Object, e As EventArgs)
            If QueryBuilder.EditMetadataContainer(_sqlContextCopy) Then
                Modified = True
            End If
        End Sub
    End Class
End Namespace
