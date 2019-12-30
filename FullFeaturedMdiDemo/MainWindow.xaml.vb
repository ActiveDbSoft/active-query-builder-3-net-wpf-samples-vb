'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data.Odbc
Imports System.Data.OleDb
Imports Oracle.ManagedDataAccess.Client
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Markup
Imports System.Windows.Threading
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View
Imports FullFeaturedMdiDemo.Common
Imports FullFeaturedMdiDemo.Connection
Imports FullFeaturedMdiDemo.MdiControl
Imports FullFeaturedMdiDemo.PropertiesForm
Imports Microsoft.Win32
Imports MySql.Data.MySqlClient
Imports Npgsql
Imports Helpers = ActiveQueryBuilder.Core.Helpers
Imports ActiveQueryBuilder.View.EventHandlers.MetadataStructureItems
Imports System.Text.RegularExpressions
Imports System.Windows.Media
Imports ActiveQueryBuilder.View.WPF

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private _selectedConnection As ConnectionInfo
    Private _sqlContext As SQLContext
    Private ReadOnly _sqlFormattingOptions As SQLFormattingOptions
    Private ReadOnly _sqlGenerationOptions As SQLGenerationOptions
    Private _showHintConnection As Boolean = True

    Public Sub New()
        ' Options to present the formatted SQL query text to end-user
        ' Use names of virtual objects, do not replace them with appropriate derived tables
        _sqlFormattingOptions = New SQLFormattingOptions() With {
            .ExpandVirtualObjects = False
        }

        ' Options to generate the SQL query text for execution against a database server
        ' Replace virtual objects with derived tables
        _sqlGenerationOptions = New SQLGenerationOptions() With {
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

        Dim menuItem as MenuItem = MenuItemLanguage.Items.Cast(Of MenuItem)().First(Function(item) DirectCast(item.Tag, String) = defLang)
        menuItem.IsChecked = True

        ' DEMO WARNING
        If ActiveQueryBuilder.Core.BuildInfo.GetEdition() = ActiveQueryBuilder.Core.BuildInfo.Edition.Trial Then
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
                .Content = New Image() With {
                    .Source = WPF.Helpers.GetImageSource(cancel),
                    .Stretch = Stretch.None
                }
            }

            AddHandler button.Click, Sub() GridRoot.Visibility = Visibility.Collapsed

            trialNoticePanel.Child = label
            GridRoot.Children.Add(trialNoticePanel)
            GridRoot.Children.Add(button)
        End If
    End Sub

    Private Sub MdiContainer1_ActiveWindowChanged(sender As Object, args As EventArgs)
        Dim window As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

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

    Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        RemoveHandler Closing, AddressOf MainWindow_Closing
        RemoveHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
    End Sub

    Private Sub Hooks_DispatcherInactive(sender As Object, e As EventArgs)
        MenuItemSave.IsEnabled = MdiContainer1.Children.Count > 0
        MenuItemQueryStatistics.IsEnabled = MdiContainer1.Children.Count > 0
        MenuItemSaveIco.IsEnabled = MdiContainer1.Children.Count > 0

        MenuItemUndo.IsEnabled = (DirectCast(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanUndo())
        MenuItemRedo.IsEnabled = (DirectCast(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanRedo())
        MenuItemCopyIco.IsEnabled = InlineAssignHelper(MenuItemCopy.IsEnabled, (DirectCast(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanCopy()))
        MenuItemPasteIco.IsEnabled = InlineAssignHelper(MenuItemPaste.IsEnabled, (DirectCast(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanPaste()))
        MenuItemCutIco.IsEnabled = InlineAssignHelper(MenuItemCut.IsEnabled, (DirectCast(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanCut()))
        MenuItemSelectAll.IsEnabled = (DirectCast(MdiContainer1.ActiveChild, ChildWindow) IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanSelectAll())
        MenuItemNewQueryToolBar.IsEnabled = InlineAssignHelper(MenuItemNewQuery.IsEnabled, _sqlContext IsNot Nothing)

        MenuItemQueryAddDerived.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanAddDerivedTable()

        MenuItemCopyUnionSq.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanCopyUnionSubQuery()
        MenuItemAddUnionSq.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanAddUnionSubQuery()
        MenuItemProp.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanShowProperties()
        MenuItemAddObject.IsEnabled = MdiContainer1.ActiveChild IsNot Nothing AndAlso DirectCast(MdiContainer1.ActiveChild, ChildWindow).CanAddObject()
        MenuItemProperties.IsEnabled = (_sqlFormattingOptions IsNot Nothing AndAlso _sqlContext IsNot Nothing)
        MenuItemProperties.Header = If(MenuItemProperties.IsEnabled, "Properties", "Properties (open a query to edit)")
        For Each item As MenuItem In MetadataItemMenu.Items.Cast(Of FrameworkElement)().Where(Function(x) TypeOf x Is MenuItem).ToList()
            item.IsEnabled = _sqlContext IsNot Nothing
        Next
    End Sub

    Private Sub MenuItemNewQuery_OnClick(sender As Object, e As RoutedEventArgs)
        MdiContainer1.Children.Add(CreateChildWindow())
    End Sub

    Private Function CreateChildWindow(Optional caption As String = "") As ChildWindow
        If _sqlContext Is Nothing Then
            Return Nothing
        End If

        Dim title As String = If(String.IsNullOrEmpty(caption), "New Query", caption)
        If MdiContainer1.Children.Any(Function(x) x.Title = title) Then
            For i As Integer = 1 To 999 Step 1
                Dim z As Integer = i
                If MdiContainer1.Children.Any(Function(x) x.Title = title & $" ({z})") Then
                    Continue For
                End If
                title += $" ({i})"
                Exit For
            Next
        End If

        Dim window As ChildWindow = New ChildWindow(_sqlContext) With {
            .State = StateWindow.Maximized,
            .Title = title,
            .SqlFormattingOptions = _sqlFormattingOptions,
            .SqlGenerationOptions = _sqlGenerationOptions
        }

        AddHandler window.Closing, AddressOf Window_Closing
        AddHandler window.SaveQueryEvent, AddressOf Window_SaveQueryEvent
        AddHandler window.SaveAsInFileEvent, AddressOf Window_SaveAsInFileEvent
        AddHandler window.SaveAsNewUserQueryEvent, AddressOf Window_SaveAsNewUserQueryEvent

        Return window
    End Function

    Private Sub Window_Closing(sender As Object, e As EventArgs)
        Dim window As ChildWindow = TryCast(sender, ChildWindow)

        If window Is Nothing Then
            Return
        End If

        RemoveHandler window.Closing, AddressOf Window_Closing
        RemoveHandler window.SaveQueryEvent, AddressOf Window_SaveQueryEvent
        RemoveHandler window.SaveAsInFileEvent, AddressOf Window_SaveAsInFileEvent
        RemoveHandler window.SaveAsNewUserQueryEvent, AddressOf Window_SaveAsNewUserQueryEvent
    End Sub

    Private Sub InitializeSqlContext()
        Try
            Mouse.OverrideCursor = Cursors.Wait

            Dim metadataProvaider As BaseMetadataProvider = Nothing

            ' create new SqlConnection object using the connections string from the connection form
            If Not _selectedConnection.IsXmlFile Then
                Select Case _selectedConnection.ConnectionType
                    Case ConnectionTypes.MSSQL
                        metadataProvaider = New MSSQLMetadataProvider() With {
                            .Connection = New SqlConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.MSAccess
                        metadataProvaider = New OLEDBMetadataProvider() With {
                            .Connection = New OleDbConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.Oracle
                        ' previous version of this demo uses deprecated System.Data.OracleClient
                        ' current version uses Oracle.ManagedDataAccess.Client which doesn't support "Integrated Security" setting
                        Dim updatedConnectionString As String = Regex.Replace(_selectedConnection.ConnectionString, "Integrated Security=.*?;", "")

                        metadataProvaider = New OracleNativeMetadataProvider() With {
                            .Connection = New OracleConnection(updatedConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.MySQL
                        metadataProvaider = New MySQLMetadataProvider() With {
                            .Connection = New MySqlConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.PostgreSQL
                        metadataProvaider = New PostgreSQLMetadataProvider() With {
                            .Connection = New NpgsqlConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.OLEDB
                        metadataProvaider = New OLEDBMetadataProvider() With {
                            .Connection = New OleDbConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case ConnectionTypes.ODBC
                        metadataProvaider = New ODBCMetadataProvider() With {
                            .Connection = New OdbcConnection(_selectedConnection.ConnectionString)
                        }
                        Exit Select
                    Case Else
                        Throw New ArgumentOutOfRangeException()
                End Select
            End If

            ' setup the query builder with metadata and syntax providers
            _sqlContext = New SQLContext() With {
                .MetadataProvider = metadataProvaider,
                .SyntaxProvider = _selectedConnection.SyntaxProvider
            }

            _sqlContext.LoadingOptions.OfflineMode = (metadataProvaider Is Nothing)

            If metadataProvaider Is Nothing Then
                _sqlContext.MetadataContainer.ImportFromXML(_selectedConnection.ConnectionString)
            End If

            TextBlockConnectionName.Text = _selectedConnection.ConnectionName

            DatabaseSchemaView1.SQLContext = _sqlContext
            DatabaseSchemaView1.InitializeDatabaseSchemaTree()

            QueriesView.SQLContext = _sqlContext
            QueriesView.SQLQuery = New SQLQuery(_sqlContext)
            QueriesView.Initialize()

            If MdiContainer1.Children.Count > 0 Then
                For Each mdiChildWindow As ChildWindow In MdiContainer1.Children.ToList()
                    mdiChildWindow.Close()
                Next
            End If
        Finally
            If _sqlContext.MetadataContainer.LoadingOptions.OfflineMode Then
                TsmiOfflineMode.IsChecked = True
            End If
            Mouse.OverrideCursor = Nothing
        End Try
    End Sub

#Region "Executed commands"
    Private Sub CommandOpen_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog() With {
            .DefaultExt = "sql",
            .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
        }

        If openFileDialog1.ShowDialog() <> True Then
            Return
        End If

        Dim sb As StringBuilder = New StringBuilder()

        Using sr As StreamReader = New StreamReader(openFileDialog1.FileName)
            Dim s As String = Nothing

            While (InlineAssignHelper(s, sr.ReadLine())) IsNot Nothing
                sb.AppendLine(s)
            End While
        End Using

        If _sqlContext Is Nothing Then
            CommandNew_OnExecuted(Nothing, Nothing)
        End If

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
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If

        SaveInFile(DirectCast(MdiContainer1.ActiveChild, ChildWindow))
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

        Using reader as MemoryStream = New MemoryStream(bytes)
            QueriesView.ImportFromXML(reader)
        End Using
    End Sub

    Private Sub CommandUndo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Undo()
    End Sub

    Private Sub CommandRedo_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Redo()
    End Sub

    Private Sub CommandCopy_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Copy()
    End Sub

    Private Sub CommandPaste_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Paste()
    End Sub

    Private Sub CommandCut_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.Cut()
    End Sub

    Private Sub CommandSelectAll_OnExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim child As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If child Is Nothing Then
            Return
        End If

        child.SelectAll()
    End Sub
#End Region

#Region "MenuItem ckick"
    Private Sub MenuItemQueryStatistics_OnClick(sender As Object, e As RoutedEventArgs)
        Dim child As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)
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
        Dim window As ChildWindow = DirectCast(MdiContainer1.ActiveChild, ChildWindow)
        window.AddDerivedTable()
    End Sub

    Private Sub MenuItemCopyUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window As ChildWindow = DirectCast(MdiContainer1.ActiveChild, ChildWindow)
        window.CopyUnionSubQuery()
    End Sub

    Private Sub MenuItemAddUnionSq_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window As ChildWindow = DirectCast(MdiContainer1.ActiveChild, ChildWindow)
        window.AddUnionSubQuery()
    End Sub

    Private Sub MenuItemProp_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window As ChildWindow = DirectCast(MdiContainer1.ActiveChild, ChildWindow)
        window.PropertiesQuery()
    End Sub

    Private Sub MenuItemAddObject_OnClick(sender As Object, e As RoutedEventArgs)
        If MdiContainer1.ActiveChild Is Nothing Then
            Return
        End If
        Dim window As ChildWindow = DirectCast(MdiContainer1.ActiveChild, ChildWindow)
        window.AddObject()
    End Sub

    Private Sub MenuItemProperties_OnClick(sender As Object, e As RoutedEventArgs)
        Dim propWindow As QueryPropertiesWindow = New QueryPropertiesWindow(_sqlContext, _sqlFormattingOptions)
        propWindow.ShowDialog()
    End Sub

    Private Sub MenuItem_OfflineMode_OnChecked(sender As Object, e As RoutedEventArgs)
        Dim menuItem As MenuItem = DirectCast(sender, MenuItem)

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

    Private Sub MenuItem_LoadMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog As OpenFileDialog = New OpenFileDialog() With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        }

        If fileDialog.ShowDialog() <> True Then
            Return
        End If

        _sqlContext.MetadataContainer.LoadingOptions.OfflineMode = True
        _sqlContext.MetadataContainer.ImportFromXML(fileDialog.FileName)

        DatabaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_SaveMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        Dim fileDialog As SaveFileDialog = New SaveFileDialog() With {
            .Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*",
            .FileName = "Metadata.xml"
        }

        If fileDialog.ShowDialog() <> True Then
            Return
        End If

        _sqlContext.MetadataContainer.LoadAll(True)
        _sqlContext.MetadataContainer.ExportToXML(fileDialog.FileName)
    End Sub

    Private Sub MenuItem_About_OnClick(sender As Object, e As RoutedEventArgs)
        Dim f As AboutForm = New AboutForm() With {
            .Owner = Me
        }

        f.ShowDialog()
    End Sub
#End Region

    Private Sub DatabaseSchemaView_OnItemDoubleClick(sender As Object, clickeditem As MetadataStructureItem)
        ' Adding a table to the currently active query.
        Dim objectMetadata As Boolean = (MetadataType.ObjectMetadata And clickeditem.MetadataItem.Type) <> 0
        Dim obj As Boolean = (MetadataType.Objects And clickeditem.MetadataItem.Type) <> 0

        If Not obj AndAlso Not objectMetadata Then
            Return
        End If

        If MdiContainer1.ActiveChild Is Nothing AndAlso MdiContainer1.Children.Count = 0 Then
            Dim childWindow As ChildWindow = CreateChildWindow()
            If childWindow Is Nothing Then
                Return
            End If

            MdiContainer1.Children.Add(childWindow)
            MdiContainer1.ActiveChild = childWindow
        End If

        Dim window As ChildWindow = DirectCast(MdiContainer1.ActiveChild, ChildWindow)

        If window Is Nothing Then
            Return
        End If

        Dim metadataItem As MetadataItem = clickeditem.MetadataItem

        If metadataItem Is Nothing Then
            Return
        End If

        If (metadataItem.Type And MetadataType.Objects) <= 0 AndAlso metadataItem.Type <> MetadataType.Field Then
            Return
        End If

        Using qualifiedName As SQLQualifiedName = metadataItem.GetSQLQualifiedName(Nothing, True)
            window.QueryView.AddObjectToActiveUnionSubQuery(qualifiedName.GetSQL())
        End Using
    End Sub

    Private Sub LanguageMenuItemChecked(sender As Object, e As RoutedEventArgs)
        e.Handled = True

        Dim menuItem As MenuItem = DirectCast(sender, MenuItem)
        Dim lng As String = menuItem.Tag.ToString()
        Language = XmlLanguage.GetLanguage(lng)

    End Sub

    Private Sub QueriesView_OnEditUserQuery(sender As Object, e As MetadataStructureItemCancelEventArgs)
        ' Opening the user query in a new query window.
        If e.MetadataStructureItem Is Nothing Then
            Return
        End If

        Dim window As ChildWindow = CreateChildWindow(e.MetadataStructureItem.MetadataItem.Name)

        window.UserMetadataStructureItem = e.MetadataStructureItem
        window.SqlSourceType = Common.Helpers.SourceType.UserQueries
        MdiContainer1.Children.Add(window)
        MdiContainer1.ActiveChild = window

        window.QueryText = DirectCast(e.MetadataStructureItem.MetadataItem, MetadataObject).Expression
    End Sub

    Private Sub Window_SaveQueryEvent(sender As Object, e As EventArgs)
        ' Saving the current query
        Dim windowChild As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If windowChild Is Nothing Then
            Return
        End If

        Select Case windowChild.SqlSourceType
            ' as a new user query
            Case Common.Helpers.SourceType.[New]
                SaveNewUserQuery(windowChild)
                Exit Select
            ' in a text file
            Case Common.Helpers.SourceType.File
                SaveInFile(windowChild)
                Exit Select
            ' replacing an exising user query 
            Case Common.Helpers.SourceType.UserQueries
                SaveUserQuery(windowChild)
                Exit Select
            Case Else
                Throw New ArgumentOutOfRangeException()
        End Select

        If windowChild.IsNeedClose Then
            windowChild.Close()
        End If
    End Sub

    Private Sub Window_SaveAsInFileEvent(sender As Object, e As EventArgs)
        Dim windowChild As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

        If windowChild Is Nothing Then
            Return
        End If

        SaveInFile(windowChild)

        If windowChild.IsNeedClose Then
            windowChild.Close()
        End If
    End Sub

    Private Sub Window_SaveAsNewUserQueryEvent(sender As Object, e As EventArgs)
        Dim windowChild As ChildWindow = TryCast(MdiContainer1.ActiveChild, ChildWindow)

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
            Dim saveFileDialog1 As SaveFileDialog = New SaveFileDialog() With {
                .DefaultExt = "sql",
                .FileName = "query",
                .Filter = "SQL query files (*.sql)|*.sql|All files|*.*"
            }

            If saveFileDialog1.ShowDialog() <> True Then
                Return
            End If

            windowChild.SqlSourceType = Common.Helpers.SourceType.File
            windowChild.FileSourceUrl = saveFileDialog1.FileName
        End If

        Using sw As StreamWriter = New StreamWriter(windowChild.FileSourceUrl)
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
        Dim title As String
        Dim newItem As MetadataStructureItem = Nothing
        If String.IsNullOrEmpty(childWindow.QueryText) Then
            Return
        End If

        Do
            Dim window As WindowNameQuery = New WindowNameQuery() With {
                .Owner = Me
            }
            Dim answer As Boolean? = window.ShowDialog()
            If answer <> True Then
                Return
            End If

            title = window.NameQuery

            If UserQueries.IsUserQueryExist(childWindow.SqlQuery.SQLContext.MetadataContainer, title) Then
                Dim path As String = QueriesView.GetPathAtUserQuery(title)
                Dim message As String = If(String.IsNullOrEmpty(path), "The same-named query already exists in the root folder.", String.Format("The same-named query already exists in the ""{0}"" folder.", path))

                MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.[Error])
                Continue Do
            End If

            Dim atItem As MetadataStructureItem = If(TryCast(QueriesView.FocusedItem, MetadataStructure), QueriesView.MetadataStructure)

            If Not UserQueries.IsFolder(atItem) Then
                atItem = atItem.Parent
            End If

            newItem = UserQueries.AddUserQuery(childWindow.SqlQuery.SQLContext.MetadataContainer, atItem, title, childWindow.FormattedQueryText, CInt(DefaultImageListImageIndices.VirtualObject), ActiveQueryBuilder.View.Helpers.GetLayout(childWindow.SqlQuery.QueryRoot))

            Exit Do
        Loop While True

        childWindow.Title = title
        childWindow.UserMetadataStructureItem = newItem
        childWindow.SqlSourceType = Common.Helpers.SourceType.UserQueries
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

        Settings.[Default].XmlFiles = App.XmlFiles
        Settings.[Default].Connections = App.Connections
        Settings.[Default].Save()

        QueriesView.Refresh()
    End Sub

    ' Closing the current query window on deleting the corresponding user query.
    Private Sub QueriesView_OnUserQueryItemRemoved(sender As Object, item As MetadataStructureItem)
        Dim childWindow As ChildWindow = MdiContainer1.Children.Cast(Of ChildWindow)().FirstOrDefault(Function(x) x.UserMetadataStructureItem Is item)

        If childWindow IsNot Nothing Then
            childWindow.ForceClose()
        End If

        SaveSettings()
    End Sub

    Private Sub QueriesView_OnErrorMessage(sender As Object, eventArgs As MetadataStructureItemErrorEventArgs)
        Dim wMessage As WindowMessage = New WindowMessage() With {
            .Owner = Me,
            .Text = eventArgs.Message,
            .WindowStartupLocation = WindowStartupLocation.CenterOwner,
            .Title = "UserQueries error"
        }

        Dim buttonOk As Button = New Button() With {
            .Content = "OK",
            .HorizontalAlignment = HorizontalAlignment.Center,
            .Width = 75
        }
        AddHandler buttonOk.Click, Sub() wMessage.Close()

        wMessage.Buttons.Add(buttonOk)
        wMessage.ShowDialog()
    End Sub

    Private Sub QueriesView_OnUserQueryItemRenamed(sender As Object, e As MetadataStructureItemTextChangedEventArgs)
        SaveSettings()
    End Sub

    Private Sub QueriesView_OnValidateItemContextMenu(sender As Object, e As MetadataStructureItemMenuEventArgs)
        e.Menu.AddItem("Copy SQL", AddressOf Execute_SqlExpression, False, True, Nothing, DirectCast(e.MetadataStructureItem.MetadataItem, MetadataObject).Expression)
    End Sub

    Private Shared Sub Execute_SqlExpression(sender As Object, eventArgs As EventArgs)
        Dim item As ICustomMenuItem = DirectCast(sender, ICustomMenuItem)

        Clipboard.SetText(item.Tag.ToString(), TextDataFormat.UnicodeText)

        Debug.WriteLine("SQL: {0}", item.Tag)
    End Sub

    Private Sub MenuItemExecuteUserQuery_OnClick(sender As Object, e As RoutedEventArgs)
        If QueriesView.FocusedItem Is Nothing Then
            Return
        End If

        Dim window As ChildWindow = CreateChildWindow(QueriesView.FocusedItem.MetadataItem.Name)

        window.UserMetadataStructureItem = QueriesView.FocusedItem
        window.SqlSourceType = Common.Helpers.SourceType.UserQueries
        MdiContainer1.Children.Add(window)
        MdiContainer1.ActiveChild = window

        window.QueryText = DirectCast(QueriesView.FocusedItem.MetadataItem, MetadataObject).Expression
        window.OpenExecuteTab()
    End Sub

    Private Sub QueriesView_OnSelectedItemChanged(sender As Object, e As EventArgs)
        MenuItemExecuteUserQuery.IsEnabled = QueriesView.FocusedItem IsNot Nothing AndAlso Not QueriesView.FocusedItem.IsFolder()
    End Sub
    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function

    Private Sub MenuItemEditMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        If _sqlContext Is Nothing Then Return

        QueryBuilder.EditMetadataContainer(_sqlContext)
    End Sub

    Private Sub MenuItem_RefreshMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        If _sqlContext.MetadataProvider Is Nothing OrElse Not _sqlContext.MetadataProvider.Connected OrElse _sqlContext Is Nothing Then
            Return
        End If

        ' to refresh metadata, just clear already loaded items
        _sqlContext.MetadataContainer.Clear()
        _sqlContext.MetadataContainer.LoadAll(True)
        DatabaseSchemaView1.InitializeDatabaseSchemaTree()
    End Sub

    Private Sub MenuItem_ClearMetadata_OnClick(sender As Object, e As RoutedEventArgs)
        If _sqlContext Is Nothing Then Return

        _sqlContext.MetadataContainer.Clear()
    End Sub
End Class
