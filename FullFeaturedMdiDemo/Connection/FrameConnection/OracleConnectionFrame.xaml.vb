'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports Oracle.ManagedDataAccess.Client
Imports System.Windows
Imports System.Windows.Input

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for OracleConnectionFrame.xaml
	''' </summary>
	Public Partial Class OracleConnectionFrame
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

            If String.IsNullOrEmpty(connectionString1) Then
                tbUserID.IsEnabled = True
                tbPassword.IsEnabled = True
            Else
                ConnectionString = connectionString1
            End If
        End Sub

		Public Function GetConnectionString() As String
			Try
                Dim builder = New OracleConnectionStringBuilder() With {
                    .ConnectionString = _connectionString, _
                    .DataSource = tbDataSource.Text, _
                    .UserID = tbUserID.Text, _
                    .Password = tbPassword.Password _
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
                Dim builder = New OracleConnectionStringBuilder() With {
                    .ConnectionString = _connectionString _
                }

                tbDataSource.Text = builder.DataSource
                tbUserID.Text = builder.UserID
                tbPassword.Password = builder.Password

                _connectionString = builder.ConnectionString
            Catch
            End Try
        End Sub

        Private Sub btnEditConnectionString_Click(sender As Object, e As EventArgs)
            Dim csef = New ConnectionStringEditWindow() With {
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
				Dim connection = New OracleConnection(ConnectionString)
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
