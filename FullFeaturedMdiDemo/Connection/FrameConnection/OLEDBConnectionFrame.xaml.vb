'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data.OleDb
Imports System.Windows
Imports System.Windows.Input
Imports MSDASC

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

        Public Sub New(connectionString1 As String)
            InitializeComponent()

            ConnectionString = connectionString1
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
                Dim builder As OleDbConnectionStringBuilder = New OleDbConnectionStringBuilder() With {
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
				Dim connection As OleDbConnection = New OleDbConnection(ConnectionString)
				connection.Open()
				connection.Close()
			Catch e As Exception
				MessageBox.Show(e.Message, App.Name)
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
				Dim dlg As DataLinks = New MSDASC.DataLinks()
                Dim adodbConnection As ADODB.Connection = New ADODB.Connection() With {
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
	End Class
End Namespace
