'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Threading
Imports ActiveQueryBuilder.Core

Namespace Connection
	''' <summary>
	''' Interaction logic for DatabaseConnectionWindow.xaml
	''' </summary>
	Public Partial Class DatabaseConnectionWindow
		Public ReadOnly Property SelectedConnection() As ConnectionInfo
			Get
				If TabControl1.SelectedIndex = 0 Then
					If LvConnections.SelectedItems.Count > 0 Then
						Dim item = DirectCast(LvConnections.SelectedItem, ConnectionListItem)
						Return DirectCast(item.Tag, ConnectionInfo)
					End If

					Return Nothing
				End If

				If LvXmlFiles.SelectedItems.Count > 0 Then
					Dim item = DirectCast(LvXmlFiles.SelectedItem, ConnectionListItem)
					Return DirectCast(item.Tag, ConnectionInfo)
				End If

				Return Nothing
			End Get
		End Property

		Public Sub New(Optional showHint As Boolean = False)
			InitializeComponent()

			GridHint.Visibility = If(showHint, Visibility.Visible, Visibility.Collapsed)

			Dim sourcelvConnection = New ObservableCollection(Of ConnectionListItem)()

			' fill connection list
            For i As Int32 = 0 To App.Connections.Count - 1
                sourcelvConnection.Add(New ConnectionListItem() With { _
                    .Name = App.Connections(i).ConnectionName, _
                    .Type = App.Connections(i).ConnectionType.ToString(), _
                    .Tag = App.Connections(i) _
                })
            Next

            LvConnections.ItemsSource = sourcelvConnection

            If LvConnections.Items.Count > 0 Then
                LvConnections.SelectedItem = LvConnections.Items(0)
            End If

            ' add preset
            Dim found = False
            Dim northwind = New ConnectionInfo(ConnectionTypes.MSSQL, "Northwind.xml", "Northwind.xml", True, Nothing, "") With { _
                .SyntaxProvider = New MSSQLSyntaxProvider() _
            }

            For i As Int32 = 0 To App.XmlFiles.Count - 1
                If App.XmlFiles(i).Equals(northwind) Then
                    found = True
                End If
            Next

            If Not found Then
                App.XmlFiles.Insert(0, northwind)
            End If

            Dim sourceXmlfiles = New ObservableCollection(Of ConnectionListItem)()

            ' fill XML files list
            For i As Int32 = 0 To App.XmlFiles.Count - 1
                sourceXmlfiles.Add(New ConnectionListItem() With { _
                    .Name = App.XmlFiles(i).ConnectionName, _
                    .Type = App.XmlFiles(i).ConnectionType.ToString(), _
                    .Tag = App.XmlFiles(i), _
                    .UserQueries = App.XmlFiles(i).UserQueries _
                })
            Next

            LvXmlFiles.ItemsSource = sourceXmlfiles

            If LvXmlFiles.Items.Count > 0 Then
                LvXmlFiles.SelectedItem = LvXmlFiles.Items(0)
            End If
            AddHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
        End Sub

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
            Dim name As String

            Do
                x += 1
                found = False
                name = String.Format("Connection {0}", x)

                For i As Int32 = 0 To App.Connections.Count - 1
                    If App.Connections(i).ConnectionName <> name Then
                        Continue For
                    End If

                    found = True
                    Exit For
                Next
            Loop While found

            Return name
        End Function

        Private Shared Function GetNewXmlFileEntryName() As String
            Dim x = 0
            Dim found As Boolean
            Dim name As String

            Do
                x += 1
                found = False
                name = String.Format("XML File {0}", x)

                For i As Int32 = 0 To App.XmlFiles.Count - 1
                    If App.XmlFiles(i).ConnectionName <> name Then
                        Continue For
                    End If
                    found = True
                    Exit For
                Next
            Loop While found

            Return name
        End Function

        Private Sub ButtonAddConnection_OnClick(sender As Object, e As RoutedEventArgs)
            Dim ci = New ConnectionInfo(ConnectionTypes.MSSQL, GetNewConnectionEntryName(), Nothing, False, Nothing, "")

            Dim cef = New AddConnectionWindow(ci) With { _
                .Owner = Me _
            }

            If cef.ShowDialog() = True Then
                Dim item = New ConnectionListItem() With { _
                    .Name = ci.ConnectionName, _
                    .Type = ci.ConnectionType.ToString(), _
                    .Tag = ci _
                }

                Dim source = TryCast(LvConnections.ItemsSource, ObservableCollection(Of ConnectionListItem))
                If source IsNot Nothing Then
                    source.Add(item)
                End If

                App.Connections.Add(ci)
                LvConnections.SelectedItem = item
            End If

            LvConnections.Focus()
            Properties.Settings.[Default].XmlFiles = App.XmlFiles

            Properties.Settings.[Default].Connections = App.Connections
            Properties.Settings.[Default].Save()
        End Sub

        Private Sub ButtonRemoveConnection_OnClick(sender As Object, e As RoutedEventArgs)
            Dim item = DirectCast(LvConnections.SelectedItem, ConnectionListItem)

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
            Dim item = DirectCast(LvConnections.SelectedItem, ConnectionListItem)

            Dim ci = DirectCast(item.Tag, ConnectionInfo)

            Dim cef = New AddConnectionWindow(ci)

            If cef.ShowDialog() = True Then
                item.Name = ci.ConnectionName
                item.Type = ci.ConnectionType.ToString()
            End If

            LvConnections.Focus()
        End Sub

        Private Sub ButtonAddXml_OnClick(sender As Object, e As RoutedEventArgs)
            Dim ci = New ConnectionInfo(ConnectionTypes.MSSQL, GetNewXmlFileEntryName(), Nothing, True, Nothing, "")

            Dim cef = New AddConnectionWindow(ci) With { _
                .Owner = Me _
            }

            If cef.ShowDialog() = True Then
                Dim item = New ConnectionListItem() With { _
                    .Name = ci.ConnectionType.ToString(), _
                    .Type = ci.ConnectionType.ToString(), _
                    .Tag = ci _
                }

                Dim source = TryCast(LvXmlFiles.ItemsSource, ObservableCollection(Of ConnectionListItem))
                If source IsNot Nothing Then
                    source.Add(item)
                End If

                App.XmlFiles.Add(ci)
                LvXmlFiles.SelectedItem = item
            End If

            LvXmlFiles.Focus()

            Properties.Settings.[Default].XmlFiles = App.XmlFiles

            Properties.Settings.[Default].Connections = App.Connections
            Properties.Settings.[Default].Save()
        End Sub

		Private Sub ButtonRemoveXml_OnClick(sender As Object, e As RoutedEventArgs)
			Dim item = DirectCast(LvXmlFiles.SelectedItem, ConnectionListItem)
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

			Properties.Settings.[Default].XmlFiles = App.XmlFiles

			Properties.Settings.[Default].Connections = App.Connections
			Properties.Settings.[Default].Save()
		End Sub

		Private Sub ButtonConfigureXml_OnClick(sender As Object, e As RoutedEventArgs)
			Dim item = DirectCast(LvXmlFiles.SelectedItem, ConnectionListItem)
			If item Is Nothing Then
				Return
			End If

			Dim ci = DirectCast(item.Tag, ConnectionInfo)

			Dim cef = New AddConnectionWindow(ci)

			If cef.ShowDialog() = True Then
				item.Name = ci.ConnectionName
				item.Type = ci.ConnectionType.ToString()
			End If

			LvXmlFiles.Focus()
		End Sub

		Private Sub BtnOk_OnClick(sender As Object, e As RoutedEventArgs)
			If SelectedConnection.SyntaxProvider Is Nothing Then
				Select Case SelectedConnection.SyntaxProviderName
					Case "ANSI SQL-2003"
						SelectedConnection.SyntaxProvider = New SQL2003SyntaxProvider()
						Exit Select
					Case "ANSI SQL-92"
						SelectedConnection.SyntaxProvider = New SQL92SyntaxProvider()
						Exit Select
					Case "ANSI SQL-89"
						SelectedConnection.SyntaxProvider = New SQL89SyntaxProvider()
						Exit Select
					Case "Firebird"
						SelectedConnection.SyntaxProvider = New FirebirdSyntaxProvider()
						Exit Select
					Case "IBM DB2"
						SelectedConnection.SyntaxProvider = New DB2SyntaxProvider()
						Exit Select
					Case "IBM Informix"
						SelectedConnection.SyntaxProvider = New InformixSyntaxProvider()
						Exit Select
					Case "Microsoft Access"
						SelectedConnection.SyntaxProvider = New MSAccessSyntaxProvider()
						Exit Select
					Case "Microsoft SQL Server"
						SelectedConnection.SyntaxProvider = New MSSQLSyntaxProvider()
						Exit Select
					Case "MySQL"
						SelectedConnection.SyntaxProvider = New MySQLSyntaxProvider()
						Exit Select
					Case "Oracle"
						SelectedConnection.SyntaxProvider = New OracleSyntaxProvider()
						Exit Select
					Case "PostgreSQL"
						SelectedConnection.SyntaxProvider = New PostgreSQLSyntaxProvider()
						Exit Select
					Case "SQLite"
						SelectedConnection.SyntaxProvider = New SQLiteSyntaxProvider()
						Exit Select
					Case "Sybase"
						SelectedConnection.SyntaxProvider = New SybaseSyntaxProvider()
						Exit Select
					Case "VistaDB"
						SelectedConnection.SyntaxProvider = New VistaDBSyntaxProvider()
						Exit Select
					Case "Universal"
						SelectedConnection.SyntaxProvider = New GenericSyntaxProvider()
						Exit Select
				End Select
			End If
			DialogResult = True
			Close()
		End Sub

		Private Sub ButtonBaseClose_OnClick(sender As Object, e As RoutedEventArgs)
			Close()
		End Sub
	End Class
End Namespace
