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
Imports Microsoft.Win32

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for MSAccessConnectionFrame.xaml
	''' </summary>
	Public Partial Class MSAccessConnectionFrame
		Implements IConnectionFrame
		Private _connectionString As String
		Private _openFileDialog1 As OpenFileDialog

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

            If [String].IsNullOrEmpty(connectionString1) Then
                tbUserID.Text = "Admin"
            Else
                ConnectionString = connectionString1
            End If
        End Sub

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Function GetConnectionString() As String
			Try
                Dim builder = New OleDbConnectionStringBuilder() With {
                    .ConnectionString = _connectionString, _
                    .Provider = "Microsoft.ACE.OLEDB.12.0", _
                    .DataSource = tbDataSource.Text _
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
                    Dim builder = New OleDbConnectionStringBuilder() With {
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
				Dim connection = New OleDbConnection(ConnectionString)
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
