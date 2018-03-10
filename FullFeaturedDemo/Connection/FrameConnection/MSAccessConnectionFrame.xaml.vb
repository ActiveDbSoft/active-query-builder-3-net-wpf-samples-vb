'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.Data.OleDb
Imports System.IO
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Input
Imports Microsoft.Win32

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for MSAccessConnectionFrame.xaml
	''' </summary>
	Public Partial Class MSAccessConnectionFrame
		Implements IConnectionFrame
		Private _connectionString As String
		Private _serverType As String

		Private ReadOnly _knownAceProviders As New List(Of String)() From { _
			"Microsoft.ACE.OLEDB.16.0", _
			"Microsoft.ACE.OLEDB.15.0", _
			"Microsoft.ACE.OLEDB.14.0", _
			"Microsoft.ACE.OLEDB.12.0" _
		}

		Private _openFileDialog1 As OpenFileDialog

		Public Event OnSyntaxProviderDetected As SyntaxProviderDetected Implements IConnectionFrame.OnSyntaxProviderDetected

		Public Property ConnectionString() As String Implements IConnectionFrame.ConnectionString
			Get
				Return GetConnectionString()
			End Get
			Set
				SetConnectionString(value)
			End Set
		End Property

		Public Sub New(connectionString__1 As String)
			InitializeComponent()

			If [String].IsNullOrEmpty(connectionString__1) Then
				tbUserID.Text = "Admin"
			Else
				ConnectionString = connectionString__1
			End If
		End Sub

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub SetServerType(serverType As String) Implements IConnectionFrame.SetServerType
			_serverType = serverType
		End Sub

		Private Shared Function GetProvidersList() As List(Of String)
			Dim reader = OleDbEnumerator.GetRootEnumerator()
			Dim result = New List(Of String)()
			While reader.Read()
				For i As Integer = 0 To reader.FieldCount - 1
					If reader.GetName(i) = "SOURCES_NAME" Then
						result.Add(reader.GetValue(i).ToString())
					End If
				Next
			End While
			reader.Close()

			Return result
		End Function

		Private Function DetectProvider() As String
			Dim providersList = GetProvidersList()
			Dim provider = String.Empty

			Dim ext = Path.GetExtension(tbDataSource.Text)
			If ext = ".accdb" Then
				For i As Integer = 0 To _knownAceProviders.Count - 1
					If providersList.Contains(_knownAceProviders(i)) Then
						provider = _knownAceProviders(i)
						Exit For
					End If
				Next
			ElseIf _serverType = "Access 97" Then
				provider = "Microsoft.Jet.OLEDB.3.0"
			ElseIf _serverType = "Access 2000 and newer" Then
				For i As Integer = 0 To _knownAceProviders.Count - 1
					If providersList.Contains(_knownAceProviders(i)) Then
						provider = _knownAceProviders(i)
						Exit For
					End If
				Next

				If provider = String.Empty Then
					provider = "Microsoft.Jet.OLEDB.4.0"
				End If
			End If

			Return provider
		End Function

		Public Function GetConnectionString() As String
			Try
				Dim builder As New OleDbConnectionStringBuilder() With { _
					 .ConnectionString = _connectionString, _
					 .DataSource = tbDataSource.Text, _
					 .Provider = DetectProvider() _
				}

				builder("User ID") = tbUserID.Text
				builder("Password") = tbPassword.Password

				_connectionString = builder.ConnectionString
			Catch
			End Try

			Return _connectionString
		End Function

		Public Sub SetConnectionString(value As String)
			_connectionString = value

			If Not String.IsNullOrEmpty(_connectionString) Then
				Try
					Dim builder = New OleDbConnectionStringBuilder() With { _
						.ConnectionString = _connectionString _
					}

					tbDataSource.Text = builder.DataSource
					tbUserID.Text = builder("User ID").ToString()
					tbPassword.Password = builder("Password").ToString()

					_connectionString = builder.ConnectionString
				Catch
				End Try
			End If
		End Sub

		Private Sub BtnBrowse_OnClick(sender As Object, e As RoutedEventArgs)
			_openFileDialog1 = New OpenFileDialog()

			If _openFileDialog1.ShowDialog() = True Then
				tbDataSource.Text = _openFileDialog1.FileName
			End If
		End Sub

		Private Sub BtnEditConnectionString_OnClick(sender As Object, e As EventArgs)
			Dim csef = New ConnectionStringEditWindow() With { _
				.ConnectionString = ConnectionString _
			}


			If csef.ShowDialog() <> True Then
				Return
			End If

			If csef.Modified Then
				ConnectionString = csef.ConnectionString
			End If
		End Sub

		Public Function TestConnection() As Boolean Implements IConnectionFrame.TestConnection
			Mouse.OverrideCursor = Cursors.Wait

			Try
				Dim connection = New OleDbConnection(ConnectionString)
				connection.Open()
				connection.Close()
			Catch e As Exception
				MessageBox.Show(e.Message, Assembly.GetEntryAssembly().GetName().Name)
				Return False
			Finally
				Mouse.OverrideCursor = Nothing
			End Try

			Return True
		End Function
	End Class
End Namespace
