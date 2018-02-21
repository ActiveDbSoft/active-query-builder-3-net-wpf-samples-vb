'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data.Odbc
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Input

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for ODBCConnectionFrame.xaml
	''' </summary>
	Public Partial Class ODBCConnectionFrame
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

		Public Sub New(connectionString_1 As String)
			InitializeComponent()

			ConnectionString = connectionString_1
		End Sub

		Public Function GetConnectionString() As String
			Try
                Dim builder = New OdbcConnectionStringBuilder() With { _
                    .ConnectionString = tbConnectionString.Text _
                }
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
				Dim builder = New OdbcConnectionStringBuilder()
				builder.ConnectionString = _connectionString
				_connectionString = builder.ConnectionString
				tbConnectionString.Text = _connectionString
			Catch
			End Try
		End Sub

		Public Function TestConnection() As Boolean Implements IConnectionFrame.TestConnection
			Mouse.OverrideCursor = Cursors.Wait

			Try
				Dim connection As New OdbcConnection(ConnectionString)
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
