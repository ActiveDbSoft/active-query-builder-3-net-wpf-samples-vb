'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows
Imports System.Windows.Input
Imports MySql.Data.MySqlClient

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for MySQLConnectionFrame.xaml
	''' </summary>
	Public Partial Class MySQLConnectionFrame
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

        Public Sub New(connectionString1 As String)
            InitializeComponent()

            If Not String.IsNullOrEmpty(connectionString1) Then
                ConnectionString = connectionString1
            End If
        End Sub

		Public Function GetConnectionString() As String
			Try
				Dim builder As New MySqlConnectionStringBuilder()
				builder.ConnectionString = _connectionString

				builder.Server = tbServer.Text
				builder.Database = tbDatabase.Text
				builder.UserID = tbUserID.Text
				builder.Password = tbPassword.Password

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
                Dim builder = New MySqlConnectionStringBuilder() With {
                    .ConnectionString = _connectionString _
                }

				tbServer.Text = builder.Server
				tbDatabase.Text = builder.Database
				tbUserID.Text = builder.UserID
				tbPassword.Password = builder.Password

				_connectionString = builder.ConnectionString
			Catch
			End Try
		End Sub

		Private Sub btnEditConnectionString_Click(sender As Object, e As EventArgs)
			Dim csef = New ConnectionStringEditWindow()

			If True Then
				csef.ConnectionString = Me.ConnectionString

				If csef.ShowDialog() <> True Then
					Return
				End If
				If csef.Modified Then
					ConnectionString = csef.ConnectionString
				End If
			End If
		End Sub

		Public Function TestConnection() As Boolean Implements IConnectionFrame.TestConnection
			Mouse.OverrideCursor = Cursors.Wait

			Try
				Dim connection As New MySqlConnection(ConnectionString)
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
	End Class
End Namespace
