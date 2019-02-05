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
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Input
Imports ActiveQueryBuilder.Core

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for OLEDBConnectionFrame.xaml
	''' </summary>
	Public Partial Class OLEDBConnectionFrame
		Implements IConnectionFrame
		Public Sub New()
			InitializeComponent()
		End Sub

		Private _connectionString As String

		Public Property ConnectionString() As String Implements IConnectionFrame.ConnectionString
			Get
				Return GetConnectionString()
			End Get
			Set
				SetConnectionString(value)
			End Set
		End Property

		Public Event OnSyntaxProviderDetected As SyntaxProviderDetected Implements IConnectionFrame.OnSyntaxProviderDetected

		Public Sub SetServerType(serverType As String) Implements IConnectionFrame.SetServerType

		End Sub

		Public Sub New(connectionString__1 As String)
			InitializeComponent()

			ConnectionString = connectionString__1
		End Sub

		Public Function GetConnectionString() As String
			Try
				Dim builder As New OleDbConnectionStringBuilder()
				builder.ConnectionString = tbConnectionString.Text
				_connectionString = builder.ConnectionString
			Catch
			End Try

			Return _connectionString
		End Function

		Public Sub SetConnectionString(value As String)
			_connectionString = value

			If String.IsNullOrEmpty(_connectionString) Then
				Return
			End If

			Try
				Dim builder = New OleDbConnectionStringBuilder() With { _
					.ConnectionString = _connectionString _
				}
				_connectionString = builder.ConnectionString
				tbConnectionString.Text = _connectionString
			Catch
			End Try
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

		Private Sub BtnBuild_OnClick(sender As Object, e As RoutedEventArgs)
			' Using COM interop with the OLE DB Service Component to display the Data Link Properties dialog box.
			'
			' Add reference to the Primary Interop Assembly (PIA) for ADO provided in the file ADODB.DLL:
			' select adodb from the .NET tab in Visual Studio .NET's Add Reference Dialog. 
			' You'll also need a reference to the Microsoft OLE DB Service Component 1.0 Type Library 
			' from the COM tab in Visual Studio .NET's Add Reference Dialog.

			Try
				Dim dlg = New MSDASC.DataLinks()
				Dim adodbConnection = New ADODB.Connection() With { _
					.ConnectionString = _connectionString _
				}
				Dim connection As Object = adodbConnection

				If dlg.PromptEdit(connection) Then
					_connectionString = adodbConnection.ConnectionString
					tbConnectionString.Text = _connectionString
				End If
			Catch exception As Exception
				MessageBox.Show("Failed to show OLEDB Data Link Properties dialog box." & vbLf & "Perhaps you have no required components installed or they are outdated." & vbLf & "Try to rebuild this demo from the source code." & vbLf & vbLf & exception.Message)
			End Try
		End Sub

		Private Sub tbConnectionString_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs)
			btnTest.IsEnabled = tbConnectionString.Text <> String.Empty
		End Sub

		Public Sub DoSyntaxDetected(syntaxType As Type)
			RaiseEvent OnSyntaxProviderDetected(syntaxType)
		End Sub

		Private Sub btnTest_Click(sender As Object, e As RoutedEventArgs)
			Dim metadataProvider = New OLEDBMetadataProvider() With { _
				.Connection = New OleDbConnection(ConnectionString) _
			}
			Dim syntaxProviderType As Type = Nothing

			Try
				syntaxProviderType = Helpers.AutodetectSyntaxProvider(metadataProvider)
			Catch exception As Exception
				MessageBox.Show(exception.Message, "Error")
			End Try

			DoSyntaxDetected(syntaxProviderType)
		End Sub
	End Class
End Namespace
