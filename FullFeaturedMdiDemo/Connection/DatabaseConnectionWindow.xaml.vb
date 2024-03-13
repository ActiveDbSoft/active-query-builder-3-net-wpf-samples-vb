''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports GeneralAssembly

Namespace Connection
    ''' <summary>
    ''' Interaction logic for DatabaseConnectionWindow.xaml
    ''' </summary>
    Partial Public Class DatabaseConnectionWindow
        Public ReadOnly Property SelectedConnection() As ConnectionInfo
            Get
                If TabControl1.SelectedIndex = 0 Then
                    If LvConnections.SelectedItems.Count > 0 Then
                        Dim item = CType(LvConnections.SelectedItem, ConnectionListItem)
                        Return DirectCast(item.Tag, ConnectionInfo)
                    End If

                    Return Nothing
                End If

                If LvXmlFiles.SelectedItems.Count > 0 Then
                    Dim item = CType(LvXmlFiles.SelectedItem, ConnectionListItem)
                    Return DirectCast(item.Tag, ConnectionInfo)
                End If

                Return Nothing
            End Get
        End Property

        Public Sub New(Optional showHint As Boolean = False)
            InitializeComponent()
            AddPresets()

            GridHint.Visibility = If(showHint, Visibility.Visible, Visibility.Collapsed)

            Dim sourcelvConnection = New ObservableCollection(Of ConnectionListItem)()

            ' fill connection list
            For i = 0 To App.Connections.Count - 1
                sourcelvConnection.Add(New ConnectionListItem With {
                    .Name = App.Connections(i).Name,
                    .Type = App.Connections(i).Type.ToString(),
                    .Tag = App.Connections(i)
                })
            Next i

            LvConnections.ItemsSource = sourcelvConnection

            If LvConnections.Items.Count > 0 Then
                LvConnections.SelectedItem = LvConnections.Items(0)
            End If

            Dim sourceXmlfiles = New ObservableCollection(Of ConnectionListItem)()

            ' fill XML files list
            For i = 0 To App.XmlFiles.Count - 1
                sourceXmlfiles.Add(New ConnectionListItem With {
                    .Name = App.XmlFiles(i).Name,
                    .Type = App.XmlFiles(i).ConnectionDescriptor.SyntaxProvider.Description,
                    .Tag = App.XmlFiles(i),
                    .UserQueries = App.XmlFiles(i).UserQueries
                })
            Next i

            LvXmlFiles.ItemsSource = sourceXmlfiles

            If LvXmlFiles.Items.Count > 0 Then
                LvXmlFiles.SelectedItem = LvXmlFiles.Items(0)
            End If
            AddHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
        End Sub

        Private Sub AddPresets()
            Dim presets = New List(Of ConnectionInfo) From {
                New ConnectionInfo("Northwind.xml", "Northwind.xml", ConnectionTypes.ODBC) With {.IsXmlFile = True},
                New ConnectionInfo(New SQLiteConnectionDescriptor(), "SQLite", ConnectionTypes.SQLite, "data source=northwind.sqlite"),
                New ConnectionInfo(New MSAccessConnectionDescriptor(), "MS Access", ConnectionTypes.MSAccess, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Nwind.mdb")
            }

            For Each preset In presets
                If Not FindConnectionInfo(preset) Then
                    If preset.IsXmlFile Then
                        App.XmlFiles.Add(preset)
                    Else
                        App.Connections.Add(preset)
                    End If
                End If
            Next preset
        End Sub

        Private Function FindConnectionInfo(connectionInfo As ConnectionInfo) As Boolean
            Dim connectionList As ConnectionList
            If connectionInfo.IsXmlFile Then
                connectionList = App.XmlFiles
            Else
                connectionList = App.Connections
            End If

            For i As Integer = 0 To connectionList.Count - 1
                If connectionList(i).Equals(connectionInfo) Then
                    Return True
                End If
            Next i

            Return False
        End Function

        Protected Overrides Sub OnClosing(e As CancelEventArgs)
            MyBase.OnClosing(e)
            RemoveHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
        End Sub

        Private Sub Hooks_DispatcherInactive(sender As Object, e As EventArgs)
            ButtonRemoveConnection.IsEnabled = (LvConnections.SelectedItems.Count > 0)
            ButtonConfigureConnection.IsEnabled = (LvConnections.SelectedItems.Count > 0)
            ButtonConfigureXml.IsEnabled = (LvXmlFiles.SelectedItems.Count > 0)
            ButtonRemoveXml.IsEnabled = (LvXmlFiles.SelectedItems.Count > 0)

            If TabControl1.SelectedIndex = 0 Then
                BtnOk.IsEnabled = (LvConnections.SelectedItems.Count > 0)
            Else
                BtnOk.IsEnabled = (LvXmlFiles.SelectedItems.Count > 0)
            End If
        End Sub

        Private Shared Function GetNewConnectionEntryName() As String
            Dim x = 0
            Dim found As Boolean
            'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not handle local variables named the same as class members well:
            Dim name_Renamed As String

            Do
                x += 1
                found = False
                name_Renamed = String.Format("Connection {0}", x)

                For i = 0 To App.Connections.Count - 1
                    If App.Connections(i).Name <> name_Renamed Then
                        Continue For
                    End If

                    found = True
                    Exit For
                Next i
            Loop While found

            Return name_Renamed
        End Function

        Private Shared Function GetNewXmlFileEntryName() As String
            Dim x = 0
            Dim found As Boolean
            'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not handle local variables named the same as class members well:
            Dim name_Renamed As String

            Do
                x += 1
                found = False
                name_Renamed = String.Format("XML File {0}", x)

                For i = 0 To App.XmlFiles.Count - 1
                    If App.XmlFiles(i).Name <> name_Renamed Then
                        Continue For
                    End If
                    found = True
                    Exit For
                Next i
            Loop While found

            Return name_Renamed
        End Function

        Private Sub ButtonAddConnection_OnClick(sender As Object, e As RoutedEventArgs)
            Dim ci = New ConnectionInfo(New MSSQLConnectionDescriptor(), GetNewConnectionEntryName(), ConnectionTypes.MSSQL, "")

            Dim cef = New ConnectionEditWindow(ci) With {.Owner = Me}

            If cef.ShowDialog() = True Then
                Dim item = New ConnectionListItem() With {
                    .Name = ci.Name,
                    .Type = ci.Type.ToString(),
                    .Tag = ci
                }

                Dim source = TryCast(LvConnections.ItemsSource, ObservableCollection(Of ConnectionListItem))
                If source IsNot Nothing Then
                    source.Add(item)
                End If

                App.Connections.Add(ci)
                LvConnections.SelectedItem = item
            End If

            LvConnections.Focus()
            My.Settings.Default.XmlFiles = App.XmlFiles

            My.Settings.Default.Connections = App.Connections
            My.Settings.Default.Save()
        End Sub

        Private Sub ButtonRemoveConnection_OnClick(sender As Object, e As RoutedEventArgs)
            Dim item = CType(LvConnections.SelectedItem, ConnectionListItem)

            If item Is Nothing Then
                Return
            End If

            Dim source = TryCast(LvConnections.ItemsSource, ObservableCollection(Of ConnectionListItem))
            If source IsNot Nothing Then
                source.Remove(item)
            End If
            App.Connections.Remove(DirectCast(item.Tag, ConnectionInfo))

            LvConnections.Focus()
        End Sub

        Private Sub ButtonConfigureConnection_OnClick(sender As Object, e As RoutedEventArgs)
            If LvConnections.SelectedItem Is Nothing Then
                Return
            End If
            Dim item = CType(LvConnections.SelectedItem, ConnectionListItem)

            Dim ci = DirectCast(item.Tag, ConnectionInfo)

            Dim cef = New ConnectionEditWindow(ci) With {.Owner = Me}

            If cef.ShowDialog() = True Then
                item.Name = ci.Name
                item.Type = ci.Type.ToString()
            End If

            LvConnections.Focus()
        End Sub

        Private Sub ButtonAddXml_OnClick(sender As Object, e As RoutedEventArgs)
            Dim ci As New ConnectionInfo(String.Empty, GetNewXmlFileEntryName(), ConnectionTypes.ODBC) With {.IsXmlFile = True}

            Dim cef = New XmlConnectionEditWindow(ci) With {.Owner = Me}

            If cef.ShowDialog() = True Then
                Dim item = New ConnectionListItem() With {
                    .Name = ci.Type.ToString(),
                    .Type = ci.ConnectionDescriptor.SyntaxProvider.Description,
                    .Tag = ci
                }

                Dim source = TryCast(LvXmlFiles.ItemsSource, ObservableCollection(Of ConnectionListItem))
                If source IsNot Nothing Then
                    source.Add(item)
                End If

                App.XmlFiles.Add(ci)
                LvXmlFiles.SelectedItem = item
            End If

            LvXmlFiles.Focus()

            My.Settings.Default.XmlFiles = App.XmlFiles

            My.Settings.Default.Connections = App.Connections
            My.Settings.Default.Save()
        End Sub

        Private Sub ButtonRemoveXml_OnClick(sender As Object, e As RoutedEventArgs)
            Dim item = CType(LvXmlFiles.SelectedItem, ConnectionListItem)
            If item Is Nothing Then
                Return
            End If

            Dim source = TryCast(LvXmlFiles.ItemsSource, ObservableCollection(Of ConnectionListItem))
            If source Is Nothing Then
                Return
            End If

            source.Remove(item)

            App.XmlFiles.Remove(DirectCast(item.Tag, ConnectionInfo))

            LvXmlFiles.Focus()

            My.Settings.Default.XmlFiles = App.XmlFiles

            My.Settings.Default.Connections = App.Connections
            My.Settings.Default.Save()
        End Sub

        Private Sub ButtonConfigureXml_OnClick(sender As Object, e As RoutedEventArgs)
            Dim item = CType(LvXmlFiles.SelectedItem, ConnectionListItem)
            If item Is Nothing Then
                Return
            End If

            Dim ci = DirectCast(item.Tag, ConnectionInfo)

            Dim cef = New XmlConnectionEditWindow(ci) With {.Owner = Me}

            If cef.ShowDialog() = True Then
                item.Name = ci.Name
                item.Type = ci.ConnectionDescriptor.SyntaxProvider.Description
            End If

            LvXmlFiles.Focus()
        End Sub

        Private Sub BtnOk_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
            Close()
        End Sub

        Private Sub ButtonBaseClose_OnClick(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub

        Private Sub LvConnections_MouseDoubleClick(sender As Object, e As System.Windows.Input.MouseButtonEventArgs)
            DialogResult = True
            Close()
        End Sub
    End Class
End Namespace
