''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Input
Imports Npgsql

Namespace Connection.FrameConnection
    ''' <summary>
    ''' Interaction logic for PostgreSQLConnectionFrame.xaml
    ''' </summary>
    Partial Public Class PostgreSQLConnectionFrame
        Implements IConnectionFrame

        Public Sub New()
            InitializeComponent()
        End Sub
        Private _connectionString As String

        Public Property ConnectionString() As String Implements IConnectionFrame.ConnectionString
            Get
                Return GetConnectionString()
            End Get
            Set(value As String)
                SetConnectionString(value)
            End Set
        End Property

        Public Event OnSyntaxProviderDetected As SyntaxProviderDetected Implements IConnectionFrame.OnSyntaxProviderDetected

        Public Sub SetServerType(serverType As String) Implements IConnectionFrame.SetServerType

        End Sub

        Public Sub New(connectionString As String)
            InitializeComponent()

            If String.IsNullOrEmpty(connectionString) Then
                tbHost.Text = "localhost"
                tbPort.Text = "5432"
                tbUserName.Text = "postgres"
            Else
                Me.ConnectionString = connectionString
            End If
        End Sub

        Public Function GetConnectionString() As String
            Try
                Dim builder = New NpgsqlConnectionStringBuilder With {
                    .ConnectionString = _connectionString,
                    .Host = tbHost.Text,
                    .Port = Convert.ToInt32(tbPort.Text),
                    .Database = tbDatabase.Text,
                    .Username = tbUserName.Text,
                    .Password = tbPassword.Password
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
                Dim builder = New NpgsqlConnectionStringBuilder With {.ConnectionString = _connectionString}

                tbHost.Text = builder.Host
                tbPort.Text = builder.Port.ToString()
                tbDatabase.Text = builder.Database
                tbUserName.Text = builder.Username

                _connectionString = builder.ConnectionString
            Catch
            End Try
        End Sub

        Private Sub btnEditConnectionString_Click(sender As Object, e As EventArgs)
            Dim csef = New ConnectionStringEditWindow With {.ConnectionString = ConnectionString}

            If Not csef.ShowDialog().Equals(True) Then
                Return
            End If
            If csef.Modified Then
                ConnectionString = csef.ConnectionString
            End If
        End Sub

        Public Function TestConnection() As Boolean Implements IConnectionFrame.TestConnection
            Mouse.OverrideCursor= Cursors.Wait

            Try
                Dim connection = New NpgsqlConnection(ConnectionString)
                connection.Open()
                connection.Close()
            Catch e As Exception
                MessageBox.Show(e.Message, System.Reflection.Assembly.GetEntryAssembly().GetName().Name)
                Return False
            Finally
                Mouse.OverrideCursor = Nothing
            End Try

            Return True
        End Function
    End Class
End Namespace
