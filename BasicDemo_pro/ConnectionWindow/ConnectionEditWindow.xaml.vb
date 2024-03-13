''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.PropertiesEditors
Imports ActiveQueryBuilder.View
Imports ActiveQueryBuilder.View.PropertiesEditors
Imports ActiveQueryBuilder.View.WPF
Imports ActiveQueryBuilder.View.WPF.Images
Imports GeneralAssembly
Imports Helpers = ActiveQueryBuilder.Core.Helpers

Namespace ConnectionWindow
    ''' <summary>
    ''' Interaction logic for ConnectionEditWindow.xaml
    ''' </summary>
    Partial Public Class ConnectionEditWindow
        Private Class ListViewItem
            Public Property Name() As String
            Public Property Icon() As CImage
        End Class

        Private _connection As BaseConnectionDescriptor

        Public ReadOnly Property Connection() As BaseConnectionDescriptor
            Get
                Return _connection
            End Get
        End Property

        Public Sub New()
            InitializeComponent()
            _connection = New MSSQLConnectionDescriptor()

            FillConnectionTypes()
            FillSyntaxTypes()

            cbConnectionType.SelectedItem = _connection.GetDescription()
            cbLoadFromDefaultDatabase.Visibility = If(_connection.SyntaxProvider.IsSupportDatabases(), Visibility.Visible, Visibility.Collapsed)
            cbLoadFromDefaultDatabase.IsChecked = _connection.MetadataLoadingOptions.LoadDefaultDatabaseOnly

            UpdateConnectionPropertiesFrames()
        End Sub

        Private Sub FillConnectionTypes()
            'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not handle local variables named the same as class members well:
            For Each name_Renamed In Misc.ConnectionDescriptorNames
                cbConnectionType.Items.Add(name_Renamed)
            Next name_Renamed
        End Sub

        Private Sub FillSyntaxTypes()
            For Each syntax As Type In Helpers.SyntaxProviderList
                Dim instance = TryCast(Activator.CreateInstance(syntax), BaseSyntaxProvider)
                cbSyntax.Items.Add(instance.Description)
            Next syntax
        End Sub

        Private Sub liFilter_Selected(sender As Object, e As RoutedEventArgs)
            If tpFilter IsNot Nothing Then
                InitializeFilterPage()
            End If
        End Sub

        Private Sub liConnection_Selected(sender As Object, e As RoutedEventArgs)
            If tpConnection IsNot Nothing Then
                tpConnection.IsSelected = True
            End If
        End Sub

        Private _isFilterPageInitialized As Boolean

        Private Sub InitializeFilterPage()
            If Not _isFilterPageInitialized Then
                Cursor = Cursors.Wait
                Try
                    databaseSchemaView1.SQLContext = _connection.GetSqlContext()
                    ClearFilters(databaseSchemaView1.SQLContext.LoadingOptions)
                    databaseSchemaView1.InitializeDatabaseSchemaTree()
                    LoadFilters()
                    _isFilterPageInitialized = True
                Catch
                    _isFilterPageInitialized = False
                Finally
                    Cursor = Nothing
                End Try
            End If

            tpFilter.IsSelected = True
        End Sub

        Private Sub ClearFilters(options As MetadataLoadingOptions)
            options.ExcludeFilter.Objects.Clear()
            options.IncludeFilter.Objects.Clear()
            options.ExcludeFilter.Schemas.Clear()
            options.IncludeFilter.Schemas.Clear()
        End Sub

        Private Sub LoadFilters()
            LoadIncludeFilters()
            LoadExcludeFilters()
        End Sub

        Private Sub LoadIncludeFilters()
            Dim filter = _connection.MetadataLoadingOptions.IncludeFilter
            LoadFilterTo(filter, lvInclude)
        End Sub

        Private Sub LoadExcludeFilters()
            Dim filter = _connection.MetadataLoadingOptions.ExcludeFilter
            LoadFilterTo(filter, lvExclude)
        End Sub

        Private Sub LoadFilterTo(filter As MetadataSimpleFilter, listBox As ListView)
            For Each filterObject In filter.Objects
                Dim item = FindItemByName(filterObject)
                listBox.Items.Add(New ListViewItem With {
                    .Name = filterObject,
                    .Icon = GetImage(item)
                })
            Next filterObject

            For Each filterSchema In filter.Schemas
                Dim item = FindItemByName(filterSchema)
                listBox.Items.Add(New ListViewItem With {
                    .Name = filterSchema,
                    .Icon = GetImage(item)
                })
            Next filterSchema
        End Sub

        Private Function FindItemByName(name As String) As MetadataItem
            Return databaseSchemaView1.MetadataStructure.MetadataItem.FindItem(Of MetadataItem)(name)
        End Function

        Private Sub CbConnectionType_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim descriptorType = GetSelectedDescriptorType()
            If _connection IsNot Nothing AndAlso _connection.GetType() Is descriptorType Then
                Return
            End If

            _connection = CreateConnectionDescriptor(descriptorType)

            If _connection Is Nothing Then
                LockUI()
                Return
            Else
                UnlockUI()
            End If

            UpdateConnectionPropertiesFrames()
        End Sub

        Private Sub LockUI()
            lbMenu.IsEnabled = False
            ButtonOk.IsEnabled = False
            cbLoadFromDefaultDatabase.Visibility = Visibility.Collapsed

            RemoveConnectionPropertiesFrame()
            RemoveSyntaxFrame()
        End Sub

        Private Sub UnlockUI()
            lbMenu.IsEnabled = True
            ButtonOk.IsEnabled = True
            cbLoadFromDefaultDatabase.Visibility = Visibility.Visible
        End Sub

        Private Sub UpdateConnectionPropertiesFrames()
            SetupSyntaxCombobox()
            RecreateConnectionFrame()
            RecreateSyntaxFrame()
        End Sub

        Private Sub SetupSyntaxCombobox()
            If IsGenericConnection() Then
                rowSyntax.Height = New GridLength(25)
                cbSyntax.SelectedItem = _connection.SyntaxProvider.Description
            Else
                rowSyntax.Height = New GridLength(0)
            End If
        End Sub

        Private Function IsGenericConnection() As Boolean
            Return TypeOf _connection Is ODBCConnectionDescriptor OrElse TypeOf _connection Is OLEDBConnectionDescriptor
        End Function

        Private Function GetSelectedDescriptorType() As Type
            Return Misc.ConnectionDescriptorList(cbConnectionType.SelectedIndex)
        End Function

        Private Function CreateConnectionDescriptor(type As Type) As BaseConnectionDescriptor
            Try
                Return TryCast(Activator.CreateInstance(type), BaseConnectionDescriptor)
            Catch e As Exception
                Dim message = If(e.InnerException IsNot Nothing, e.InnerException.Message, e.Message)
                MessageBox.Show(message & ControlChars.CrLf & " " & ControlChars.CrLf & "To fix this error you may need to install the appropriate database client software or " & ControlChars.CrLf & " re-compile the project from sources and add the needed assemblies to the References section.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)

                Return Nothing
            End Try
        End Function

        Private Sub RecreateConnectionFrame()
            RemoveConnectionPropertiesFrame()

            pbConnection.ClearProperties()
            Dim container = PropertiesFactory.GetPropertiesContainer(_connection.MetadataProperties)
            TryCast(pbConnection, IPropertiesControl).SetProperties(container)
        End Sub

        Private Sub RemoveConnectionPropertiesFrame()
            ClearProperties(_connection.MetadataProperties)
        End Sub

        Private Sub ClearProperties(properties As ObjectProperties)
            properties.GroupProperties.Clear()
            properties.PropertiesEditors.Clear()
        End Sub

        Private Sub RecreateSyntaxFrame()
            RemoveSyntaxFrame()
            Dim syntxProps = _connection.SyntaxProperties
            If syntxProps Is Nothing Then
                Return
            End If

            Dim container = PropertiesFactory.GetPropertiesContainer(syntxProps)
            TryCast(pbSyntax, IPropertiesControl).SetProperties(container)

            cbLoadFromDefaultDatabase.Visibility = If(_connection.SyntaxProvider.IsSupportDatabases(), Visibility.Visible, Visibility.Hidden)
        End Sub

        Private Sub RemoveSyntaxFrame()
            pbSyntax.ClearProperties()
            Dim syntxProps = _connection.SyntaxProperties
            If syntxProps Is Nothing Then
                Return
            End If

            ClearProperties(syntxProps)
        End Sub

        Private Sub CbSyntax_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If Not IsGenericConnection() Then
                Return
            End If

            Dim syntaxType = GetSelectedSyntaxType()
            If _connection.SyntaxProvider.GetType() Is syntaxType Then
                Return
            End If

            _connection.SyntaxProvider = CreateSyntaxProvider(syntaxType)

            RecreateSyntaxFrame()
        End Sub

        Private Function GetSelectedSyntaxType() As Type
            Return Helpers.SyntaxProviderList(cbSyntax.SelectedIndex)
        End Function

        Private Function CreateSyntaxProvider(type As Type) As BaseSyntaxProvider
            Return TryCast(Activator.CreateInstance(type), BaseSyntaxProvider)
        End Function

        Private Sub CbLoadFromDefaultDatabase_OnChecked(sender As Object, e As RoutedEventArgs)
            _connection.MetadataLoadingOptions.LoadDefaultDatabaseOnly = If(cbLoadFromDefaultDatabase.IsChecked, False)
        End Sub

        Private Sub BtnAdd_OnClick(sender As Object, e As RoutedEventArgs)
            If tcFilter.SelectedItem Is tpInclude Then
                AddIncludeFilter(databaseSchemaView1.SelectedItems)
            ElseIf tcFilter.SelectedItem Is tpExclude Then
                AddExcludeFilter(databaseSchemaView1.SelectedItems)
            End If
        End Sub

        Private Sub AddIncludeFilter(items() As MetadataStructureItem)
            Dim filter = _connection.MetadataLoadingOptions.IncludeFilter
            For Each structureItem In items
                Dim metadataItem = structureItem.MetadataItem
                If metadataItem Is Nothing Then
                    Continue For
                End If

                If metadataItem.Type.IsNamespace() Then
                    filter.Schemas.Add(metadataItem.NameFull)
                    lvInclude.Items.Add(New ListViewItem With {
                        .Name = metadataItem.NameFull,
                        .Icon = GetImage(metadataItem)
                    })
                ElseIf metadataItem.Type.IsObject() Then
                    filter.Objects.Add(metadataItem.NameFull)
                    lvInclude.Items.Add(New ListViewItem With {
                        .Name = metadataItem.NameFull,
                        .Icon = GetImage(metadataItem)
                    })
                End If
            Next structureItem
        End Sub

        Private Sub AddExcludeFilter(items() As MetadataStructureItem)
            Dim filter = _connection.MetadataLoadingOptions.ExcludeFilter
            For Each structureItem In items
                Dim metadataItem = structureItem.MetadataItem
                If metadataItem Is Nothing Then
                    Continue For
                End If

                If metadataItem.Type.IsNamespace() Then
                    filter.Schemas.Add(metadataItem.NameFull)
                    lvExclude.Items.Add(New ListViewItem With {
                        .Name = metadataItem.NameFull,
                        .Icon = GetImage(metadataItem)
                    })
                ElseIf metadataItem.Type.IsObject() Then
                    filter.Objects.Add(metadataItem.NameFull)
                    lvExclude.Items.Add(New ListViewItem With {
                        .Name = metadataItem.NameFull,
                        .Icon = GetImage(metadataItem)
                    })
                End If
            Next structureItem
        End Sub

        Private Sub DeleteFilter(itemName As String)
            Dim filter As MetadataSimpleFilter = Nothing
            If tcFilter.SelectedItem Is tpInclude Then
                filter = _connection.MetadataLoadingOptions.IncludeFilter
            ElseIf tcFilter.SelectedItem Is tpExclude Then
                filter = _connection.MetadataLoadingOptions.ExcludeFilter
            End If

            If filter IsNot Nothing Then
                filter.Objects.Remove(itemName)
                filter.Schemas.Remove(itemName)
            End If

            If tcFilter.SelectedItem Is tpInclude Then
                Dim item = lvInclude.Items.Cast(Of ListViewItem)().FirstOrDefault(Function(x) x.Name = itemName)
                If item IsNot Nothing Then
                    lvInclude.Items.Remove(item)
                End If
            ElseIf tcFilter.SelectedItem Is tpExclude Then
                Dim item = lvExclude.Items.Cast(Of ListViewItem)().FirstOrDefault(Function(x) x.Name = itemName)
                If item IsNot Nothing Then
                    lvExclude.Items.Remove(item)
                End If
            End If
        End Sub

        Private Function GetImage(item As MetadataItem) As CImage
            If item Is Nothing Then
                Return Nothing
            End If

            Select Case item.Type
                Case MetadataType.Server
                    Return Metadata.Server.Value
                Case MetadataType.Database
                    Return Metadata.Database.Value
                Case MetadataType.Schema
                    Return Metadata.Schema.Value
                Case MetadataType.Package
                    Return Metadata.Package.Value
                Case MetadataType.Table
                    Return Metadata.UserTable.Value
                Case MetadataType.View
                    Return Metadata.UserView.Value
                Case MetadataType.Procedure
                    Return Metadata.UserProcedure.Value
                Case MetadataType.Synonym
                    Return Metadata.UserSynonym.Value
                Case Else
                    Return Nothing
            End Select
        End Function

        Private Sub BtnRemove_OnClick(sender As Object, e As RoutedEventArgs)
            If tcFilter.SelectedItem Is tpInclude Then
                Dim itemsToDelete = New List(Of ListViewItem)()
                For Each selectedItem As ListViewItem In lvInclude.SelectedItems
                    itemsToDelete.Add(selectedItem)
                Next selectedItem

                For Each item In itemsToDelete
                    DeleteFilter(item.Name)
                Next item
            ElseIf tcFilter.SelectedItem Is tpExclude Then
                Dim itemsToDelete = New List(Of ListViewItem)()
                For Each selectedItem As ListViewItem In lvExclude.SelectedItems
                    itemsToDelete.Add(selectedItem)
                Next selectedItem

                For Each item In itemsToDelete
                    DeleteFilter(item.Name)
                Next item
            End If
        End Sub

        Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = False
        End Sub

        Private Sub LvExclude_OnDragOver(sender As Object, e As DragEventArgs)
            e.Effects = DragDropEffects.All
        End Sub

        Private Sub LvInclude_OnDragOver(sender As Object, e As DragEventArgs)
            e.Effects = DragDropEffects.All
        End Sub

        Private Sub DropItems(e As DragEventArgs, toInclude As Boolean)
            Dim dragObject = TryCast(e.Data.GetData(e.Data.GetFormats()(0)), MetadataDragObject)
            If dragObject IsNot Nothing Then
                If toInclude Then
                    AddIncludeFilter(dragObject.MetadataStructureItems.ToArray())
                Else
                    AddExcludeFilter(dragObject.MetadataStructureItems.ToArray())
                End If
            End If
        End Sub

        Private Sub LvInclude_OnDrop(sender As Object, e As DragEventArgs)
            DropItems(e, True)
        End Sub

        Private Sub LvExclude_OnDrop(sender As Object, e As DragEventArgs)
            DropItems(e, False)
        End Sub

        Private Sub DatabaseSchemaView1_OnItemDoubleClick(sender As Object, item As MetadataStructureItem)
            BtnAdd_OnClick(Me, Nothing)
        End Sub
    End Class
End Namespace
