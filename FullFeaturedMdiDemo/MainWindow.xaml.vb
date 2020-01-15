'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Media
Imports System.Windows.Threading
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View
Imports FullFeaturedMdiDemo.Common
Imports FullFeaturedMdiDemo.MdiControl
Imports FullFeaturedMdiDemo.PropertiesForm
Imports Microsoft.Win32
Imports Helpers = ActiveQueryBuilder.Core.Helpers
Imports ActiveQueryBuilder.View.EventHandlers.MetadataStructureItems
Imports ActiveQueryBuilder.View.WPF
Imports BuildInfo = ActiveQueryBuilder.Core.BuildInfo

Partial Public Class MainWindow
    Private _selectedConnection As ConnectionInfo
    Private _sqlContext As SQLContext
    Private ReadOnly _sqlFormattingOptions As SQLFormattingOptions
    Private ReadOnly _sqlGenerationOptions As SQLGenerationOptions
    Private _showHintConnection As Boolean = True
    Private _options As Options

    Public Sub New()
        _sqlFormattingOptions = New SQLFormattingOptions With {
            .ExpandVirtualObjects = False
        }
        _sqlGenerationOptions = New SQLGenerationOptions With {
            .ExpandVirtualObjects = True
        }
        InitializeComponent()
        AddHandler Closing, AddressOf MainWindow_Closing
        AddHandler MdiContainer1.ActiveWindowChanged, AddressOf MdiContainer1_ActiveWindowChanged
        AddHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
        Dim currentLang As String = System.Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName
        LoadLanguage()

        Dim defLang As String = "en"

        If Helpers.Localizer.Languages.Contains(currentLang.ToLower()) Then
            Language = XmlLanguage.GetLanguage(currentLang)
            defLang = currentLang.ToLower()
        End If

        Dim menuItem As MenuItem = MenuItemLanguage.Items.Cast(Of MenuItem)().First(Function(item As MenuItem) DirectCast(item.Tag, String) = defLang)
        menuItem.IsChecked = True

        TryToLoadOptions()

        ' DEMO WARNING
        If ActiveQueryBuilder.Core.BuildInfo.GetEdition() = BuildInfo.Edition.Trial Then
            Dim trialNoticePanel As Border = New Border() With {
                    .BorderBrush = Brushes.Black,
                    .BorderThickness = New Thickness(1),
                    .Background = Brushes.LightGreen,
                    .Padding = New Thickness(5),
                    .Margin = New Thickness(0, 0, 0, 2)
                    }
            trialNoticePanel.SetValue(Grid.RowProperty, 1)

            Dim label As TextBlock = New TextBlock() With {
                    .Text = "Generation of random aliases for the query output columns is the limitation of the trial version. The full version is free from this behavior.",
                    .HorizontalAlignment = HorizontalAlignment.Left,
                    .VerticalAlignment = VerticalAlignment.Top
                    }

            Dim button As Button = New Button() With {
                        .Background = Brushes.Transparent,
                        .Padding = New Thickness(0),
                        .BorderThickness = New Thickness(0),
                        .Cursor = Cursors.Hand,
                        .Margin = New Thickness(0, 0, 5, 0),
                        .HorizontalAlignment = HorizontalAlignment.Right,
                        .VerticalAlignment = VerticalAlignment.Center,
                        .Content = New Image() With {.Stretch = Stretch.None}
                    }

            AddHandler button.Click, Sub() GridRoot.Visibility = Visibility.Collapsed

            trialNoticePanel.Child = label
            GridRoot.Children.Add(trialNoticePanel)
            GridRoot.Children.Add(button)
        End If
    End Sub

    Private Sub TryToLoadOptions()
        If String.IsNullOrEmpty(FullFeaturedMdiDemo.Properties.Settings.Default.Options) Then Return
        _options = New Options()
        _options.CreateDefaultOptions()

        Try
            _options.DeserializeFromString(Properties.Settings.Default.Options)
        Catch
            _options = Nothing
            Properties.Settings.Default.Options = String.Empty
        End Try
    End Sub

    Private _shown As Boolean

    Protected Overrides Sub OnContentRendered(e As EventArgs)
        MyBase.OnContentRendered(e)
        If _shown Then Return
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
        For Each lng As String In Helpers.Localizer.Languages
            If lng.ToLower() = "auto" OrElse lng.ToLower() = "default" Then
                Continue For
            End If

            Dim culture As CultureInfo = New CultureInfo(lng)

            Dim stroke As String = String.Format("{0}", culture.DisplayName)

            Dim menuItem As MenuItem = New MenuItem() With {
                    .Header = stroke,
                    .Tag = lng,
                    .IsCheckable = True
                    }

            MenuItemLanguage.Items.Add(menuItem)
            menuItem.SetValue(MenuBehavior.OptionGroupNameProperty, "group")
        Next
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs)
        RemoveHandler Closing, AddressOf MainWindow_Closing
        RemoveHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
    End Sub

    Private Sub Hooks_DispatcherInactive(sender As Object, e As EventArgs)
        MenuItemSave.IsEnabled = MdiContainer1.Children.Count > 0
        MenuItemQueryStatistics.IsEnabled = MdiContainer1.Children.Count > 0
        MenuItemSaveIco.IsEnabled = MdiContainer1.Children.Count > 0
        MenuItemUndo.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanUndo())
        MenuItemRedo.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanRedo())
        MenuItemCopyIco.IsEnabled = Impl.Assign(MenuItemCopy.IsEnabled, (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanCopy()))
        MenuItemPasteIco.IsEnabled = Impl.Assign(MenuItemPaste.IsEnabled, (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanPaste()))
        MenuItemCutIco.IsEnabled = Impl.Assign(MenuItemCut.IsEnabled, (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanCut()))
        MenuItemSelectAll.IsEnabled = (CType(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanSelectAll())
        MenuItemNewQueryToolBar.IsEnabled = Impl.Assign(MenuItemNewQuery.IsEnabled, _sqlContext IsNot Nothing)
        MenuItemQueryAddDerived.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanAddDerivedTable()
        MenuItemEditMetadata.IsEnabled = _sqlContext IsNot Nothing
        MenuItemCopyUnionSq.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanCopyUnionSubQuery()
        MenuItemAddUnionSq.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanAddUnionSubQuery()
        MenuItemProp.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanShowProperties()
        MenuItemAddObject.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso (CType(MdiContainer1.ActiveChild, ChildWindow)).CanAddObject()
        MenuItemProperties.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing
        MenuItemProperties.Header = If(MenuItemProperties.IsEnabled, "Properties", "Properties (open a query to edit)")

        For Each item As MenuItem In MetadataItemMenu.Items.Cast(Of FrameworkElement)().Where(Function(x) TypeOf x Is MenuItem).ToList()
            item.IsEnabled = _sqlContext IsNot Nothing
        Next
    End Sub

    Private Sub MenuItemNewQuery_OnClick(sender As Object, e As RoutedEventArgs)
        Dim window As ChildWindow = CreateChildWindow()
        MdiContainer1.Children.Add(window)
        Dim options As Options = window.GetOptions()
        _options = options

        For Each child As ChildWindow In MdiContainer1.Children.OfType(Of ChildWindow)()
            child.SetOptions(options)
        Next
    End Sub

    Private Function CreateChildWindow(Optional caption As String = "") As ChildWindow
        If _sqlContext Is Nothing Then Return Nothing
        Dim title = If(String.IsNullOrEmpty(caption), "New Query", caption)

        If MdiContainer1.Children.Any(Function(x) x.Title = title) Then

            For i As Integer = 1 To 1000 - 1
                If MdiContainer1.Children.Any(Function(x) x.Title = title & " (" & i & ")") Then Continue For
                title += " (" & i & ")"
                Exit For
            Next
        End If

        Dim window As ChildWindow = New ChildWindow(_sqlContext, DatabaseSchemaView1) With {
            .State = StateWindow.Maximized,
            .Title = title,
            .SqlFormattingOptions = _sqlFormattingOptions,
            .SqlGenerationOptions = _sqlGenerationOptions
        }
        AddHandler window.Closing, AddressOf Window_Closing
        AddHandler window.SaveQueryEvent, AddressOf Window_SaveQueryEvent
        AddHandler window.SaveAsInFileEvent, AddressOf Window_SaveAsInFileEvent
        AddHandler window.SaveAsNewUserQueryEvent, AddressOf Window_SaveAsNewUserQueryEvent

        If _options IsNot Nothing Then window.SetOptions(_options)
        Return window
    End Function

    Private Sub Window_Closing(sender As Object, e As EventArgs)
        Dim window As ChildWindow = TryCast(sender, ChildWindow)
        If window Is Nothing Then Return
        RemoveHandler window.Closing, AddressOf Window_Closing
        RemoveHandler window.SaveQueryEvent, AddressOf Window_SaveQueryEvent
        RemoveHandler window.SaveAsInFileEvent, AddressOf Window_SaveAsInFileEvent
        RemoveHandler window.SaveAsNewUserQueryEvent, AddressOf Window_SaveAsNewUserQueryEvent
    End Sub

    Private Function InitializeSqlContext() As Boolean
        Try
            Mouse.OverrideCursor = Cursors.Wait

            If _selectedConnection.IsXmlFile Then
                _sqlContext = New SQLContext With {
                    .SyntaxProvider =_selectedConnection.ConnectionDescriptor.SyntaxProvider
                    }
                _sqlContext.LoadingOptions.OfflineMode = true
                _sqlContext.MetadataStructureOptions.AllowFavourites = True

                Try
                    _sqlContext.MetadataContainer.ImportFromXML(_selectedConnection.XMLPath)
                Catch e As Exception
                    MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            Else

                Try
                    _sqlContext = _selectedConnection.ConnectionDescriptor.GetSqlContext
                    _sqlContext.MetadataStructureOptions.AllowFavourites = True
                Catch e As Exception
                    MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End Try
            End If

            TextBlockConnectionName.Text = _selectedConnection.Name
            DatabaseSchemaView1.SQLContext = _sqlContext
            DatabaseSchemaView1.InitializeDatabaseSchemaTree
            QueriesView.SQLContext = _sqlContext
            QueriesView.SQLQuery = New SQLQuery(_sqlContext)
            QueriesView.Initialize

            If MdiContainer1.Children.Count > 0 Then

                For Each mdiChildWindow as MdiChildWindow In MdiContainer1.Children.ToList
                    mdiChildWindow.Close
                Next
            End If

        Finally

            If _sqlContext IsNot Nothing AndAlso _sqlContext.MetadataContainer.LoadingOptions.OfflineMode Then
                TsmiOfflineMode.IsChecked = True
            End If

            Mouse.OverrideCursor = Nothing
        End Try

        Return True
    End Function

    Private Sub CommandOpen_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog With {
            .DefaultExt = "sql",
            .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
        }
        If openFileDialog1.ShowDialog() <> True Then Return
        Dim sb = New StringBuilder()

        Using sr As StreamReader = New StreamReader(openFileDialog1.FileName)
            Dim s As String

            While (Impl.Assign(s, sr.ReadLine())) IsNot Nothing
                sb.AppendLine(s)
            End While
        End Using

        If _sqlContext Is Nothing Then CommandNew_OnExecuted(Nothing, Nothing)

        Try
            Mouse.OverrideCursor = Cursors.Wait
            Dim window As ChildWindow = CreateChildWindow(Path.GetFileName(openFileDialog1.FileName))
            window.FileSourceUrl = openFileDialog1.FileName
            window.QueryText = sb.ToString()
            window.SqlFormattingOptions = _sqlFormattingOptions
            window.SqlSourceType = Common.Helpers.SourceType.File
            MdiContainer1.Children.Add(window)
        Finally
            Mouse.OverrideCursor = Nothing
        End Try
    End Sub

    Private Sub CommandSave_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then Return
        SaveInFile(CType(MdiContainer1.ActiveChild, ChildWindow))
    End Sub

    Private Sub CommandExit_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Close()
    End Sub

    Private Sub CommandNew_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim cf As DatabaseConnectionWindow = New DatabaseConnectionWindow(_showHintConnection) With {
                .Owner = Me
                }

        _showHintConnection = False

        If cf.ShowDialog() <> True Then
            Return
        End If
        _selectedConnection = cf.SelectedConnection

        InitializeSqlContext()

        If String.IsNullOrEmpty(_selectedConnection.UserQueries) Then
            Return
        End If

        Dim bytes As Byte() = Encoding.UTF8.GetBytes(_selectedConnection.UserQueries)

        Using reader As MemoryStream = New MemoryStream(bytes)
            QueriesView.ImportFromXML(reader)
        End Using
    End Sub

    Private Sub CommandUndo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child Is Nothing Then Return
        child.Undo()
    End Sub

    Private Sub CommandRedo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child Is Nothing Then Return
        child.Redo()
    End Sub

    Private Sub CommandCopy_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child Is Nothing Then Return
        child.Copy()
    End Sub

    Private Sub CommandPaste_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child Is Nothing Then Return
        child.Paste()
    End Sub

    Private Sub CommandCut_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child Is Nothing Then Return
        child.Cut()
    End Sub

    Private Sub CommandSelectAll_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child Is Nothing Then Return
        child.SelectAll()
    End Sub

    Private Sub MenuItemQueryStatistics_OnClick(sender As Object, e As RoutedEventArgs)
        Dim child = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If child IsNot Nothing Then child.ShowQueryStatistics()
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
        If MdiContainer1.ActiveChild Is Nothing Then Return
        Dim window = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.AddDerivedTable()
    End Sub

    Private Sub MenuItemCopyUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then Return
        Dim window As ChildWindow = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.CopyUnionSubQuery()
    End Sub

    Private Sub MenuItemAddUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then Return
        Dim window As ChildWindow = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.AddUnionSubQuery()
    End Sub

    Private Sub MenuItemProp_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then Return
        Dim window As ChildWindow = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.PropertiesQuery()
    End Sub

    Private Sub MenuItemAddObject_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then Return
        Dim window As ChildWindow = CType(MdiContainer1.ActiveChild, ChildWindow)
        window.AddObject()
    End Sub

    Private Sub MenuItemProperties_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then Return
        Dim window As ChildWindow = CType(MdiContainer1.ActiveChild, ChildWindow)
        Dim propWindow As QueryPropertiesWindow = New QueryPropertiesWindow(window, DatabaseSchemaView1)
        propWindow.Owner = Me
        propWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner
        propWindow.ShowDialog()
        
    End Sub

    Private Sub MenuItem_OfflineMode_OnChecked(sender As Object, e As RoutedEventArgs)
        Dim menuItem As MenuItem = CType(sender, MenuItem)

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
        If _sqlContext Is Nothing OrElse _sqlContext.MetadataProvider Is Nothing OrElse Not _sqlContext.MetadataProvider.Connected Then Return
        _sqlContext.MetadataContainer.Clear()
        _sqlContext.MetadataContainer.LoadAll(True)
        DatabaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_ClearMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        If _sqlContext Is Nothing Then Return
        _sqlContext.MetadataContainer.Clear()
    End Sub

    Private Sub MenuItem_LoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog = New OpenFileDialog With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        }
        If fileDialog.ShowDialog() <> True Then Return
        _sqlContext.MetadataContainer.LoadingOptions.OfflineMode = True
        _sqlContext.MetadataContainer.ImportFromXML(fileDialog.FileName)
        DatabaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_SaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog As SaveFileDialog = New SaveFileDialog With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
            .FileName = "Metadata.xml"
        }
        If fileDialog.ShowDialog() <> True Then Return
        _sqlContext.MetadataContainer.LoadAll(True)
        _sqlContext.MetadataContainer.ExportToXML(fileDialog.FileName)
    End Sub

    Private Sub MenuItem_About_OnClick(sender As Object, e As RoutedEventArgs)
        Dim f As AboutForm = New AboutForm With {
            .Owner = Me
        }
        f.ShowDialog()
    End Sub

    Private Sub DatabaseSchemaView_OnItemDoubleClick(sender As Object, clickeditem As MetadataStructureItem)
        If clickeditem.MetadataItem Is Nothing Then Return
        Dim objectMetadata As Boolean = (MetadataType.ObjectMetadata And clickeditem.MetadataItem.Type) <> 0
        Dim obj As Boolean = (MetadataType.Objects And clickeditem.MetadataItem.Type) <> 0
        If Not obj AndAlso Not objectMetadata Then Return

        If MdiContainer1.ActiveChild Is Nothing AndAlso MdiContainer1.Children.Count = 0 Then
            Dim childWindow As ChildWindow = CreateChildWindow()
            If childWindow Is Nothing Then Return
            MdiContainer1.Children.Add(childWindow)
            MdiContainer1.ActiveChild = childWindow
        End If

        Dim window As ChildWindow = CType(MdiContainer1.ActiveChild, ChildWindow)
        If window Is Nothing Then Return
        Dim metadataItem As MetadataItem = clickeditem.MetadataItem
        If metadataItem Is Nothing Then Return
        If (metadataItem.Type And MetadataType.Objects) <= 0 AndAlso metadataItem.Type <> MetadataType.Field Then Return

        Using qualifiedName As SQLQualifiedName = metadataItem.GetSQLQualifiedName(Nothing, True)
            window.QueryView.AddObjectToActiveUnionSubQuery(qualifiedName.GetSQL())
        End Using
    End Sub

    Private Sub LanguageMenuItemChecked(sender As Object, e As RoutedEventArgs)
        e.Handled = True
        Dim menuItem = CType(sender, MenuItem)
        Dim lng = menuItem.Tag.ToString()
        Language = XmlLanguage.GetLanguage(lng)
    End Sub

    Private Sub QueriesView_OnEditUserQuery(sender As Object, e As MetadataStructureItemCancelEventArgs)
        If e.MetadataStructureItem Is Nothing Then Return
        Dim window = CreateChildWindow(e.MetadataStructureItem.MetadataItem.Name)
        window.UserMetadataStructureItem = e.MetadataStructureItem
        window.SqlSourceType = Common.Helpers.SourceType.UserQueries
        MdiContainer1.Children.Add(window)
        MdiContainer1.ActiveChild = window
        window.QueryText = (CType(e.MetadataStructureItem.MetadataItem, MetadataObject)).Expression
    End Sub

    Private Sub Window_SaveQueryEvent(sender As Object, e As EventArgs)
        Dim windowChild = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If windowChild Is Nothing Then Return

        Select Case windowChild.SqlSourceType
            Case Common.Helpers.SourceType.[New]
                SaveNewUserQuery(windowChild)
            Case Common.Helpers.SourceType.File
                SaveInFile(windowChild)
            Case Common.Helpers.SourceType.UserQueries
                SaveUserQuery(windowChild)
            Case Else
                Throw New ArgumentOutOfRangeException()
        End Select

        If windowChild.IsNeedClose Then windowChild.Close()
    End Sub

    Private Sub Window_SaveAsInFileEvent(sender As Object, e As EventArgs)
        Dim windowChild = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If windowChild Is Nothing Then Return
        SaveInFile(windowChild)
        If windowChild.IsNeedClose Then windowChild.Close()
    End Sub

    Private Sub Window_SaveAsNewUserQueryEvent(sender As Object, e As EventArgs)
        Dim windowChild = TryCast(MdiContainer1.ActiveChild, ChildWindow)
        If windowChild Is Nothing Then Return
        SaveNewUserQuery(windowChild)
        If windowChild.IsNeedClose Then windowChild.Close()
    End Sub

    Private Shared Sub SaveInFile(windowChild As ChildWindow)
        If String.IsNullOrEmpty(windowChild.FileSourceUrl) OrElse Not File.Exists(windowChild.FileSourceUrl) Then
            Dim saveFileDialog1 = New SaveFileDialog() With {
                .DefaultExt = "sql",
                .FileName = "query",
                .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
            }
            If saveFileDialog1.ShowDialog() <> True Then Return
            windowChild.SqlSourceType = Common.Helpers.SourceType.File
            windowChild.FileSourceUrl = saveFileDialog1.FileName
        End If

        Using sw = New StreamWriter(windowChild.FileSourceUrl)
            sw.Write(windowChild.FormattedQueryText)
        End Using

        windowChild.IsModified = False
    End Sub

    Private Sub SaveUserQuery(childWindow As ChildWindow)
        If Not childWindow.IsModified Then Return
        If childWindow.UserMetadataStructureItem Is Nothing Then Return
        If Not UserQueries.IsUserQueryExist(childWindow.SqlQuery.SQLContext.MetadataContainer, childWindow.UserMetadataStructureItem.MetadataName) Then Return
        UserQueries.SaveUserQuery(childWindow.SqlQuery.SQLContext.MetadataContainer, childWindow.UserMetadataStructureItem, childWindow.FormattedQueryText, ActiveQueryBuilder.View.Helpers.GetLayout(childWindow.SqlQuery.QueryRoot))
        childWindow.IsModified = False
        SaveSettings()
    End Sub

    Private Sub SaveNewUserQuery(childWindow As ChildWindow)
        Dim title As String
        Dim newItem As MetadataStructureItem
        If String.IsNullOrEmpty(childWindow.QueryText) Then Return

        Do
            Dim window = New WindowNameQuery With {
                .Owner = Me
            }
            Dim answer = window.ShowDialog()
            If answer <> True Then Return
            title = window.NameQuery

            If UserQueries.IsUserQueryExist(childWindow.SqlQuery.SQLContext.MetadataContainer, title) Then
                Dim path = QueriesView.GetPathAtUserQuery(title)
                Dim message = If(String.IsNullOrEmpty(path), "The same-named query already exists in the root folder.", String.Format("The same-named query already exists in the ""{0}"" folder.", path))
                MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.[Error])
                Continue Do
            End If

            Dim atItem = If(QueriesView.FocusedItem, QueriesView.MetadataStructure)
            If Not UserQueries.IsFolder(atItem) Then atItem = atItem.Parent
            newItem = UserQueries.AddUserQuery(childWindow.SqlQuery.SQLContext.MetadataContainer, atItem, title, childWindow.FormattedQueryText, CInt(DefaultImageListImageIndices.VirtualObject), ActiveQueryBuilder.View.Helpers.GetLayout(childWindow.SqlQuery.QueryRoot))
            Exit Do
        Loop While True

        childWindow.Title = title
        childWindow.UserMetadataStructureItem = newItem
        childWindow.SqlSourceType = Common.Helpers.SourceType.UserQueries
        childWindow.IsModified = False
        SaveSettings()
    End Sub

    Private Sub SaveSettings()
        Try

            Using stream As MemoryStream = New MemoryStream()
                QueriesView.ExportToXML(stream)
                stream.Position = 0

                Using reader As StreamReader = New StreamReader(stream)
                    _selectedConnection.UserQueries = reader.ReadToEnd()
                End Using
            End Using

        Catch exception As Exception
            Throw New QueryBuilderException(exception.Message, exception)
        End Try

        Properties.Settings.Default.XmlFiles = App.XmlFiles
        Properties.Settings.Default.Connections = App.Connections
        Properties.Settings.Default.Save()
        QueriesView.Refresh()
    End Sub

    Private Sub QueriesView_OnUserQueryItemRemoved(sender As Object, item As MetadataStructureItem)
        Dim childWindow As ChildWindow = MdiContainer1.Children.Cast(Of ChildWindow)().FirstOrDefault(Function(x As ChildWindow) x.UserMetadataStructureItem Is item)
        If childWindow IsNot Nothing Then childWindow.ForceClose()
        SaveSettings()
    End Sub

    Private Sub QueriesView_OnErrorMessage(sender As Object, eventArgs As MetadataStructureItemErrorEventArgs)
        Dim wMessage As WindowMessage = New WindowMessage With {
            .Owner = Me,
            .Text = eventArgs.Message,
            .WindowStartupLocation = WindowStartupLocation.CenterOwner,
            .Title = "UserQueries error"
        }
        Dim buttonOk As Button = New Button With {
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
        If e.MetadataStructureItem Is Nothing Then Return
        Dim obj = TryCast(e.MetadataStructureItem.MetadataItem, MetadataObject)
        If obj Is Nothing Then Return
        e.Menu.AddItem("Copy SQL", AddressOf Execute_SqlExpression, False, True, Nothing, obj.Expression)
    End Sub

    Private Shared Sub Execute_SqlExpression(sender As Object, eventArgs As EventArgs)
        Dim item = CType(sender, ICustomMenuItem)
        Clipboard.SetText(item.Tag.ToString(), TextDataFormat.UnicodeText)
    End Sub

    Private Sub MenuItemExecuteUserQuery_OnClick(sender As Object, e As RoutedEventArgs)
        If QueriesView.FocusedItem Is Nothing Then Return
        Dim window = CreateChildWindow(QueriesView.FocusedItem.MetadataItem.Name)
        window.UserMetadataStructureItem = QueriesView.FocusedItem
        window.SqlSourceType = Common.Helpers.SourceType.UserQueries
        MdiContainer1.Children.Add(window)
        MdiContainer1.ActiveChild = window
        window.QueryText = (CType(QueriesView.FocusedItem.MetadataItem, MetadataObject)).Expression
        window.OpenExecuteTab()
    End Sub

    Private Sub QueriesView_OnSelectedItemChanged(sender As Object, e As EventArgs)
        MenuItemExecuteUserQuery.IsEnabled = QueriesView.FocusedItem IsNot Nothing AndAlso Not QueriesView.FocusedItem.IsFolder()
    End Sub

    Private Class Impl
        <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
        Shared Function Assign(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class
End Class
