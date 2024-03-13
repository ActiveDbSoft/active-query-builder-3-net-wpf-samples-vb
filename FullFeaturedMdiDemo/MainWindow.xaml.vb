''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Globalization
Imports System.IO
Imports System.Windows.Markup
Imports Microsoft.Win32
Imports Helpers = ActiveQueryBuilder.Core.Helpers
Imports ActiveQueryBuilder.View.EventHandlers.MetadataStructureItems
Imports BuildInfo = ActiveQueryBuilder.Core.BuildInfo
Imports GeneralAssembly
Imports FullFeaturedMdiDemo.AutoGeneratedProperties
Imports GeneralAssembly.Windows
Imports FullFeaturedMdiDemo.Common
Imports GeneralAssembly.Windows.SaveWindows
Imports FullFeaturedMdiDemo.MdiControl
Imports FullFeaturedMdiDemo.Connection
Imports GeneralAssembly.Common

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private _selectedConnection As ConnectionInfo
    Private _sqlContext As SQLContext
    Private ReadOnly _sqlFormattingOptions As SQLFormattingOptions
    Private ReadOnly _sqlGenerationOptions As SQLGenerationOptions
    Private _showHintConnection As Boolean = True
    Private _options As Options

    Public Sub New()
        ' Options to present the formatted SQL query text to end-user
        ' Use names of virtual objects, do not replace them with appropriate derived tables
        _sqlFormattingOptions = New SQLFormattingOptions With {.ExpandVirtualObjects = False}

        ' Options to generate the SQL query text for execution against a database server
        ' Replace virtual objects with derived tables
        _sqlGenerationOptions = New SQLGenerationOptions With {.ExpandVirtualObjects = True}

        InitializeComponent()

        AddHandler Closing, AddressOf MainWindow_Closing
        AddHandler MdiContainer1.ActiveWindowChanged, AddressOf MdiContainer1_ActiveWindowChanged
        AddHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive

        Dim currentLang = System.Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName

        LoadLanguage()

        Dim defLang = "en"

        If Helpers.Localizer.Languages.Contains(currentLang.ToLower()) Then
            Language = XmlLanguage.GetLanguage(currentLang)
            defLang = currentLang.ToLower()
        End If

        Dim menuItem = MenuItemLanguage.Items.Cast(Of MenuItem)().First(Function(item) CStr(item.Tag) = defLang)
        menuItem.IsChecked = True

        TryToLoadOptions()

        ' DEMO WARNING
        If BuildInfo.GetEdition() = BuildInfo.Edition.Trial Then
            Dim trialNoticePanel = New Border With {
                .BorderBrush = System.Windows.Media.Brushes.Black,
                .BorderThickness = New Thickness(1),
                .Background = Media.Brushes.LightGreen,
                .Padding = New Thickness(5),
                .Margin = New Thickness(0, 0, 0, 2)
            }
            trialNoticePanel.SetValue(Grid.RowProperty, 1)

            Dim label = New TextBlock With {
                .Text = "Generation of random aliases for the query output columns is the limitation of the trial version. The full version is free from this behavior.",
                .HorizontalAlignment = HorizontalAlignment.Left,
                .VerticalAlignment = VerticalAlignment.Top
            }

            Dim button = New Button With {
                .Background = Media.Brushes.Transparent,
                .Padding = New Thickness(0),
                .BorderThickness = New Thickness(0),
                .Cursor = Cursors.Hand,
                .Margin = New Thickness(0, 0, 5, 0),
                .HorizontalAlignment = HorizontalAlignment.Right,
                .VerticalAlignment = VerticalAlignment.Center,
                .Content = New Controls.Image With {
                    .Source = My.Resources.cancel.GetImageSource(),
                    .Stretch = Stretch.None
                }
            }

            AddHandler button.Click, Sub()
                                         GridRoot.Visibility = Visibility.Collapsed
                                     End Sub

            trialNoticePanel.Child = label
            GridRoot.Children.Add(trialNoticePanel)
            GridRoot.Children.Add(button)
        End If

    End Sub

    Private Sub TryToLoadOptions()
        If String.IsNullOrEmpty(My.Settings.Default.Options) Then
            Return
        End If

        _options = New Options()
        _options.CreateDefaultOptions()
        Try
            _options.DeserializeFromString(My.Settings.Default.Options)
        Catch
            _options = Nothing
            My.Settings.Default.Options = String.Empty
        End Try
    End Sub

    Private _shown As Boolean

    Protected Overrides Sub OnContentRendered(e As EventArgs)
        MyBase.OnContentRendered(e)

        If _shown Then
            Return
        End If

        _shown = True

        CommandNew_OnExecuted(Me, Nothing)
    End Sub

    Private Sub MdiContainer1_ActiveWindowChanged(sender As Object, args As EventArgs)
        Dim window = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If window IsNot Nothing Then
            QueriesView.QueryView = window.QueryView
            QueriesView.SQLQuery = window.SqlQuery
        End If
    End Sub

    Private Sub LoadLanguage()
        'INSTANT VB NOTE: The variable language was renamed since Visual Basic does not handle local variables named the same as class members well:
        For Each language_Renamed In Helpers.Localizer.Languages
            If language_Renamed.ToLower() = "auto" OrElse language_Renamed.ToLower() = "default" Then
                Continue For
            End If

            Dim culture = New CultureInfo(language_Renamed)

            Dim stroke = String.Format("{0}", culture.DisplayName)

            Dim menuItem = New MenuItem With {
                .Header = stroke,
                .Tag = language_Renamed,
                .IsCheckable = True
            }

            MenuItemLanguage.Items.Add(menuItem)
            menuItem.SetValue(GroupedMenuBehavior.OptionGroupNameProperty, "group")
        Next language_Renamed
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs)
        RemoveHandler Closing, AddressOf MainWindow_Closing
        RemoveHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
    End Sub

    Private Sub Hooks_DispatcherInactive(sender As Object, e As EventArgs)
        MenuItemSave.IsEnabled = MdiContainer1.Children.Count > 0
        MenuItemQueryStatistics.IsEnabled = MdiContainer1.Children.Count > 0
        MenuItemSaveIco.IsEnabled = MdiContainer1.Children.Count > 0

        MenuItemUndo.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanUndo())
        MenuItemRedo.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanRedo())
        MenuItemCopy.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanCopy())
        MenuItemCopyIco.IsEnabled = MenuItemCopy.IsEnabled
        MenuItemPaste.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanPaste())
        MenuItemPasteIco.IsEnabled = MenuItemPaste.IsEnabled
        MenuItemCut.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanCut())
        MenuItemCutIco.IsEnabled = MenuItemCut.IsEnabled
        MenuItemSelectAll.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanSelectAll())
        MenuItemNewQuery.IsEnabled = _sqlContext IsNot Nothing
        MenuItemNewQueryToolBar.IsEnabled = MenuItemNewQuery.IsEnabled

        MenuItemQueryAddDerived.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanAddDerivedTable()

        MenuItemEditMetadata.IsEnabled = _sqlContext IsNot Nothing

        MenuItemCopyUnionSq.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanCopyUnionSubQuery()
        MenuItemAddUnionSq.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanAddUnionSubQuery()
        MenuItemProp.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanShowProperties()

        MenuItemAddObject.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso CType(MdiContainer1.ActiveChild, ChildWindow).CanAddObject()
        MenuItemUserExpression.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing
        MenuItemProperties.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing
        MenuItemProperties.Header = If(MenuItemProperties.IsEnabled, "Properties", "Properties (open a query to edit)")

        For Each item In MetadataItemMenu.Items.Cast(Of FrameworkElement)().Where(Function(x) TypeOf x Is MenuItem).ToList()
            item.IsEnabled = _sqlContext IsNot Nothing
        Next item
    End Sub

    Private Sub MenuItemNewQuery_OnClick(sender As Object, e As RoutedEventArgs)
        Dim window = CreateChildWindow()

        MdiContainer1.Children.Add(window)

        Dim options = window.GetOptions()
        _options = options

        For Each child In MdiContainer1.Children.OfType(Of ChildWindow)()
            child.SetOptions(options)
        Next child
    End Sub

    Private Function CreateChildWindow(Optional caption As String = "") As ChildWindow
        If _sqlContext Is Nothing Then
            Return Nothing
        End If

        'INSTANT VB NOTE: The variable title was renamed since Visual Basic does not handle local variables named the same as class members well:
        Dim title_Renamed = If(String.IsNullOrEmpty(caption), "New Query", caption)
        If MdiContainer1.Children.Any(Function(x) x.Title = title_Renamed) Then
            For i = 1 To 999
                Dim counter = i
                If MdiContainer1.Children.Any(Function(x) x.Title = title_Renamed & " (" & counter & ")") Then
                    Continue For
                End If
                title_Renamed &= " (" & i & ")"
                Exit For
            Next i
        End If

        Dim window = New ChildWindow(_sqlContext, DatabaseSchemaView1) With {
            .State = StateWindow.Maximized,
            .Title = title_Renamed,
            .SqlFormattingOptions = _sqlFormattingOptions,
            .SqlGenerationOptions = _sqlGenerationOptions
        }

        AddHandler window.Closing, AddressOf Window_Closing
        AddHandler window.SaveQueryEvent, AddressOf Window_SaveQueryEvent
        AddHandler window.SaveAsInFileEvent, AddressOf Window_SaveAsInFileEvent
        AddHandler window.SaveAsNewUserQueryEvent, AddressOf Window_SaveAsNewUserQueryEvent

        If _options IsNot Nothing Then
            window.SetOptions(_options)
        End If

        Return window
    End Function

    Private Sub Window_Closing(sender As Object, e As EventArgs)
        Dim window = TryCast(sender, ChildWindow)

        If window Is Nothing Then
            Return
        End If

        RemoveHandler window.Closing, AddressOf Window_Closing
        RemoveHandler window.SaveQueryEvent, AddressOf Window_SaveQueryEvent
        RemoveHandler window.SaveAsInFileEvent, AddressOf Window_SaveAsInFileEvent
        RemoveHandler window.SaveAsNewUserQueryEvent, AddressOf Window_SaveAsNewUserQueryEvent
    End Sub

    Private Function InitializeSqlContext() As Boolean
        Try
            Mouse.OverrideCursor = Cursors.Wait

            ' setup the query builder with metadata and syntax providers
            If _selectedConnection.IsXmlFile Then
                _sqlContext = New SQLContext With {
                    .SyntaxProvider = _selectedConnection.ConnectionDescriptor.SyntaxProvider
                    }

                _sqlContext.LoadingOptions.OfflineMode = True
                _sqlContext.MetadataStructureOptions.AllowFavourites = True

                Try
                    _sqlContext.MetadataContainer.ImportFromXML(_selectedConnection.XMLPath)
                Catch e As Exception
                    MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else
                Try
                    _sqlContext = _selectedConnection.ConnectionDescriptor.GetSqlContext()
                    _sqlContext.MetadataStructureOptions.AllowFavourites = True
                Catch e As Exception
                    MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            End If

            TextBlockConnectionName.Text = _selectedConnection.Name

            DatabaseSchemaView1.SQLContext = _sqlContext
            DatabaseSchemaView1.InitializeDatabaseSchemaTree()

            QueriesView.SQLContext = _sqlContext
            QueriesView.SQLQuery = New SQLQuery(_sqlContext)
            QueriesView.Initialize()

            If MdiContainer1.Children.Count > 0 Then
                'INSTANT VB NOTE: The variable mdiChildWindow was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
                For Each mdiChildWindow_Renamed In MdiContainer1.Children.ToList()
                    mdiChildWindow_Renamed.Close()
                Next mdiChildWindow_Renamed
            End If
        Finally
            If _sqlContext IsNot Nothing AndAlso _sqlContext.MetadataContainer.LoadingOptions.OfflineMode Then
                TsmiOfflineMode.IsChecked = True
            End If

            Mouse.OverrideCursor = Nothing
        End Try

        Return True
    End Function

#Region "Executed commands"
    Private Sub CommandOpen_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim openFileDialog1 = New OpenFileDialog With {
            .DefaultExt = "sql",
            .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
        }

        If Not openFileDialog1.ShowDialog().Equals(True) Then
            Return
        End If

        Dim sb = New StringBuilder()

        Using sr = New StreamReader(openFileDialog1.FileName)
            Dim s As String

            'INSTANT VB WARNING: An assignment within expression was extracted from the following statement:
            'ORIGINAL LINE: while ((s = sr.ReadLine()) != null)
            s = sr.ReadLine()
            Do While s IsNot Nothing
                sb.AppendLine(s)
                s = sr.ReadLine()
            Loop
        End Using

        If _sqlContext Is Nothing Then
            CommandNew_OnExecuted(Nothing, Nothing)
        End If

        Try
            Mouse.OverrideCursor = Cursors.Wait

            Dim window = CreateChildWindow(Path.GetFileName(openFileDialog1.FileName))

            window.FileSourceUrl = openFileDialog1.FileName
            window.QueryText = sb.ToString()
            window.SqlFormattingOptions = _sqlFormattingOptions
            window.SqlSourceType = SourceType.File

            MdiContainer1.Children.Add(window)
        Finally
            Mouse.OverrideCursor = Nothing
        End Try
    End Sub

    Private Sub CommandSave_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If

        SaveInFile(CType(MdiContainer1.ActiveChild, ChildWindow))
    End Sub

    Private Sub CommandExit_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Close()
    End Sub

    Private Sub CommandNew_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim cf = New DatabaseConnectionWindow(_showHintConnection) With {.Owner = Me}

        _showHintConnection = False

        If cf.ShowDialog() <> True Then
            Return
        End If
        _selectedConnection = cf.SelectedConnection

        If Not InitializeSqlContext() Then
            Return
        End If

        If String.IsNullOrEmpty(_selectedConnection.UserQueries) Then
            Return
        End If

        Dim bytes = Encoding.UTF8.GetBytes(_selectedConnection.UserQueries)

        Using reader = New MemoryStream(bytes)
            QueriesView.ImportFromXML(reader)
        End Using
    End Sub

    Private Sub CommandUndo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Undo()
    End Sub

    Private Sub CommandRedo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Redo()
    End Sub

    Private Sub CommandCopy_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Copy()
    End Sub

    Private Sub CommandPaste_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Paste()
    End Sub

    Private Sub CommandCut_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Cut()
    End Sub

    Private Sub CommandSelectAll_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.SelectAll()
    End Sub
#End Region

#Region "MenuItem ckick"
    Private Sub MenuItemQueryStatistics_OnClick(sender As Object, e As RoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child IsNot Nothing Then
            child.ShowQueryStatistics()
        End If
    End Sub

    Private Sub MenuItemCascade_OnClick(sender As Object, e As RoutedEventArgs)
        MdiContainer1.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub MenuItemVertical_OnClick(sender As Object, e As RoutedEventArgs)
        MdiContainer1.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub MenuItemHorizontal_OnClick(sender As Object, e As RoutedEventArgs)
        MdiContainer1.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub MenuItemQueryAddDerived_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.AddDerivedTable()
    End Sub

    Private Sub MenuItemCopyUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.CopyUnionSubQuery()
    End Sub

    Private Sub MenuItemAddUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.AddUnionSubQuery()
    End Sub

    Private Sub MenuItemProp_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.PropertiesQuery()
    End Sub

    Private Sub MenuItemAddObject_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.AddObject()
    End Sub

    Private Sub MenuItemProperties_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)
        Dim propWindow = New QueryPropertiesWindow(window, DatabaseSchemaView1)
        propWindow.ShowDialog()
    End Sub

    Private Sub MenuItem_OfflineMode_OnChecked(sender As Object, e As RoutedEventArgs)
        Dim menuItem = DirectCast(sender, MenuItem)

        If menuItem.IsChecked Then
            Try
                Cursor = Cursors.Wait

                _sqlContext.MetadataContainer.LoadAll(True)
            Finally
                Cursor = Cursors.Arrow
            End Try
        End If

        _sqlContext.MetadataContainer.LoadingOptions.OfflineMode = menuItem.IsChecked

    End Sub

    Private Sub MenuItemEditMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        QueryBuilder.EditMetadataContainer(_sqlContext)
    End Sub

    Private Sub MenuItem_RefreshMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' to refresh metadata, just clear already loaded items

        If _sqlContext Is Nothing OrElse _sqlContext.MetadataProvider Is Nothing OrElse Not _sqlContext.MetadataProvider.Connected Then
            Return
        End If

        _sqlContext.MetadataContainer.Clear()
        _sqlContext.MetadataContainer.LoadAll(True)

        DatabaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_ClearMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        ' to refresh metadata, just clear already loaded items

        If _sqlContext Is Nothing Then
            Return
        End If

        _sqlContext.MetadataContainer.Clear()
    End Sub

    Private Sub MenuItem_LoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog = New OpenFileDialog With {.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"}

        If Not fileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        _sqlContext.MetadataContainer.LoadingOptions.OfflineMode = True
        _sqlContext.MetadataContainer.ImportFromXML(fileDialog.FileName)

        DatabaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_SaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog = New SaveFileDialog With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
            .FileName = "Metadata.xml"
        }

        If Not fileDialog.ShowDialog().Equals(True) Then
            Return
        End If

        _sqlContext.MetadataContainer.LoadAll(True)
        _sqlContext.MetadataContainer.ExportToXML(fileDialog.FileName)
    End Sub

    Private Sub MenuItem_About_OnClick(sender As Object, e As RoutedEventArgs)
        Dim f = New AboutWindow With {.Owner = Me}

        f.ShowDialog()
    End Sub
#End Region

    Private Sub DatabaseSchemaView_OnItemDoubleClick(sender As Object, clickeditem As MetadataStructureItem)
        If clickeditem.MetadataItem Is Nothing Then
            Return
        End If

        ' Adding a table to the currently active query.
        Dim objectMetadata = (MetadataType.ObjectMetadata And clickeditem.MetadataItem.Type) <> 0
        Dim obj = (MetadataType.Objects And clickeditem.MetadataItem.Type) <> 0

        If Not obj AndAlso Not objectMetadata Then
            Return
        End If

        If MdiContainer1.ActiveChild Is Nothing AndAlso MdiContainer1.Children.Count = 0 Then
            Dim childWindow = CreateChildWindow()
            If childWindow Is Nothing Then
                Return
            End If

            MdiContainer1.Children.Add(childWindow)
            MdiContainer1.ActiveChild = childWindow
        End If

        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)

        If window Is Nothing Then
            Return
        End If

        Dim metadataItem = clickeditem.MetadataItem

        If metadataItem Is Nothing Then
            Return
        End If

        If (metadataItem.Type And MetadataType.Objects) <= 0 AndAlso metadataItem.Type <> MetadataType.Field Then
            Return
        End If

        Using qualifiedName = metadataItem.GetSQLQualifiedName(Nothing, True)
            window.QueryView.AddObjectToActiveUnionSubQuery(qualifiedName.GetSQL())
        End Using
    End Sub

    Private Sub LanguageMenuItemChecked(sender As Object, e As RoutedEventArgs)
        e.Handled = True

        Dim menuItem = DirectCast(sender, MenuItem)
        Dim lng = menuItem.Tag.ToString()
        Language = XmlLanguage.GetLanguage(lng)

    End Sub

    Private Sub QueriesView_OnEditUserQuery(sender As Object, e As MetadataStructureItemCancelEventArgs)
        ' Opening the user query in a new query window.
        If e.MetadataStructureItem Is Nothing Then
            Return
        End If

        Dim window = CreateChildWindow(e.MetadataStructureItem.MetadataItem.Name)

        window.UserMetadataStructureItem = e.MetadataStructureItem
        window.SqlSourceType = SourceType.UserQueries
        MdiContainer1.Children.Add(window)
        MdiContainer1.ActiveChild = window

        window.QueryText = CType(e.MetadataStructureItem.MetadataItem, MetadataObject).Expression
    End Sub

    Private Sub Window_SaveQueryEvent(sender As Object, e As EventArgs)
        ' Saving the current query
        Dim windowChild = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If windowChild Is Nothing Then
            Return
        End If

        Select Case windowChild.SqlSourceType
                ' as a new user query
            Case SourceType.[New]
                SaveNewUserQuery(windowChild)
                ' in a text file
            Case SourceType.File
                SaveInFile(windowChild)
                ' replacing an exising user query 
            Case SourceType.UserQueries
                SaveUserQuery(windowChild)
            Case Else
                Throw New ArgumentOutOfRangeException()
        End Select

        If windowChild.IsNeedClose Then
            windowChild.Close()
        End If
    End Sub

    Private Sub Window_SaveAsInFileEvent(sender As Object, e As EventArgs)
        Dim windowChild = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If windowChild Is Nothing Then
            Return
        End If

        SaveInFile(windowChild)

        If windowChild.IsNeedClose Then
            windowChild.Close()
        End If
    End Sub

    Private Sub Window_SaveAsNewUserQueryEvent(sender As Object, e As EventArgs)
        Dim windowChild = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If windowChild Is Nothing Then
            Return
        End If

        SaveNewUserQuery(windowChild)

        If windowChild.IsNeedClose Then
            windowChild.Close()
        End If
    End Sub

    Private Shared Sub SaveInFile(windowChild As ChildWindow)
        If String.IsNullOrEmpty(windowChild.FileSourceUrl) OrElse Not File.Exists(windowChild.FileSourceUrl) Then
            Dim saveFileDialog1 = New SaveFileDialog() With {
                .DefaultExt = "sql",
                .FileName = "query",
                .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
            }

            If saveFileDialog1.ShowDialog() <> True Then
                Return
            End If

            windowChild.SqlSourceType = SourceType.File
            windowChild.FileSourceUrl = saveFileDialog1.FileName
        End If

        Using sw = New StreamWriter(windowChild.FileSourceUrl)
            sw.Write(windowChild.FormattedQueryText)
        End Using
        windowChild.IsModified = False
    End Sub

    Private Sub SaveUserQuery(childWindow As ChildWindow)
        If Not childWindow.IsModified Then
            Return
        End If
        If childWindow.UserMetadataStructureItem Is Nothing Then
            Return
        End If

        If Not UserQueries.IsUserQueryExist(childWindow.SqlQuery.SQLContext.MetadataContainer, childWindow.UserMetadataStructureItem.MetadataName) Then
            Return
        End If

        UserQueries.SaveUserQuery(childWindow.SqlQuery.SQLContext.MetadataContainer, childWindow.UserMetadataStructureItem, childWindow.FormattedQueryText, ActiveQueryBuilder.View.Helpers.GetLayout(childWindow.SqlQuery.QueryRoot))

        childWindow.IsModified = False
        SaveSettings()
    End Sub

    Private Sub SaveNewUserQuery(childWindow As ChildWindow)
        'INSTANT VB NOTE: The variable title was renamed since Visual Basic does not handle local variables named the same as class members well:
        Dim title_Renamed As String
        Dim newItem As MetadataStructureItem = Nothing
        If String.IsNullOrEmpty(childWindow.QueryText) Then
            Return
        End If

        Do
            Dim window = New WindowNameQuery With {.Owner = Me}
            Dim answer = window.ShowDialog()
            If Not answer.Equals(True) Then
                Return
            End If

            title_Renamed = window.NameQuery

            If UserQueries.IsUserQueryExist(childWindow.SqlQuery.SQLContext.MetadataContainer, title_Renamed) Then
                Dim path = QueriesView.GetPathAtUserQuery(title_Renamed)
                Dim message = If(String.IsNullOrEmpty(path), "The same-named query already exists in the root folder.", String.Format("The same-named query already exists in the ""{0}"" folder.", path))

                MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Continue Do
            End If

            Dim atItem = If(QueriesView.FocusedItem, QueriesView.MetadataStructure)

            If Not UserQueries.IsFolder(atItem) Then
                atItem = atItem.Parent
            End If

            newItem = UserQueries.AddUserQuery(childWindow.SqlQuery.SQLContext.MetadataContainer, atItem, title_Renamed, childWindow.FormattedQueryText, CInt(DefaultImageListImageIndices.VirtualObject), ActiveQueryBuilder.View.Helpers.GetLayout(childWindow.SqlQuery.QueryRoot))

            Exit Do

        Loop While True

        childWindow.Title = title_Renamed
        childWindow.UserMetadataStructureItem = newItem
        childWindow.SqlSourceType = SourceType.UserQueries
        childWindow.IsModified = False

        SaveSettings()
    End Sub

    ' Saving user queries to the connection settings
    Private Sub SaveSettings()
        Try
            Using stream As New MemoryStream()
                QueriesView.ExportToXML(stream)
                stream.Position = 0

                Using reader As New StreamReader(stream)
                    _selectedConnection.UserQueries = reader.ReadToEnd()
                End Using
            End Using
        Catch exception As Exception
            Throw New QueryBuilderException(exception.Message, exception)
        End Try

        My.Settings.Default.XmlFiles = App.XmlFiles
        My.Settings.Default.Connections = App.Connections
        My.Settings.Default.Save()

        QueriesView.Refresh()
    End Sub

    ' Closing the current query window on deleting the corresponding user query.
    Private Sub QueriesView_OnUserQueryItemRemoved(sender As Object, item As MetadataStructureItem)
        Dim childWindow = MdiContainer1.Children.Cast(Of ChildWindow)().FirstOrDefault(Function(x) x.UserMetadataStructureItem Is item)

        If childWindow IsNot Nothing Then
            childWindow.ForceClose()
        End If

        SaveSettings()
    End Sub

    Private Sub QueriesView_OnErrorMessage(sender As Object, eventArgs As MetadataStructureItemErrorEventArgs)
        Dim wMessage = New WindowMessage With {
            .Owner = Me,
            .Text = eventArgs.Message,
            .WindowStartupLocation = WindowStartupLocation.CenterOwner,
            .Title = "UserQueries error"
        }

        Dim buttonOk = New Button With {
            .Content = "OK",
            .HorizontalAlignment = HorizontalAlignment.Center,
            .Width = 75
        }
        AddHandler buttonOk.Click, Sub()
                                       wMessage.Close()
                                   End Sub

        wMessage.Buttons.Add(buttonOk)
        wMessage.ShowDialog()
    End Sub

    Private Sub QueriesView_OnUserQueryItemRenamed(sender As Object, e As MetadataStructureItemTextChangedEventArgs)
        SaveSettings()
    End Sub

    Private Sub QueriesView_OnValidateItemContextMenu(sender As Object, e As MetadataStructureItemMenuEventArgs)
        If e.MetadataStructureItem Is Nothing Then
            Return
        End If

        Dim obj = TryCast(e.MetadataStructureItem.MetadataItem, MetadataObject)
        If obj Is Nothing Then
            Return
        End If

        e.Menu.AddItem("Copy SQL", AddressOf Execute_SqlExpression, False, True, Nothing, obj.Expression)
    End Sub

    Private Shared Sub Execute_SqlExpression(sender As Object, eventArgs As EventArgs)
        Dim item = DirectCast(sender, ICustomMenuItem)

        Clipboard.SetText(item.Tag.ToString(), TextDataFormat.UnicodeText)
    End Sub

    Private Sub MenuItemExecuteUserQuery_OnClick(sender As Object, e As RoutedEventArgs)
        If QueriesView.FocusedItem Is Nothing Then
            Return
        End If

        Dim window = CreateChildWindow(QueriesView.FocusedItem.MetadataItem.Name)

        window.UserMetadataStructureItem = QueriesView.FocusedItem
        window.SqlSourceType = SourceType.UserQueries
        MdiContainer1.Children.Add(window)
        MdiContainer1.ActiveChild = window

        window.QueryText = CType(QueriesView.FocusedItem.MetadataItem, MetadataObject).Expression
        window.OpenExecuteTab()
    End Sub

    Private Sub QueriesView_OnSelectedItemChanged(sender As Object, e As EventArgs)
        MenuItemExecuteUserQuery.IsEnabled = QueriesView.FocusedItem IsNot Nothing AndAlso Not QueriesView.FocusedItem.IsFolder()
    End Sub

    Private Sub MenuItemUserExpression_OnClick(sender As Object, e As RoutedEventArgs)
        Dim childWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If childWindow Is Nothing Then Return

        Dim window = New EditUserExpressionWindow With {
                .Owner = Me,
                .WindowStartupLocation = WindowStartupLocation.CenterOwner}
        window.Load(childWindow.QueryView)
        window.ShowDialog()
    End Sub
End Class
