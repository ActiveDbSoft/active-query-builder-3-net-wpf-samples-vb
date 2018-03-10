'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Input

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for MSSQLConnectionFrame.xaml
	''' </summary>
	Public Partial Class MSSQLConnectionFrame
		Implements IConnectionFrame
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

		Public Sub New(connectionString_1 As String)
			InitializeComponent()
			If String.IsNullOrEmpty(connectionString_1) Then
				tbDataSource.Text = "(local)"
				cmbIntegratedSecurity.SelectedIndex = 0
				tbUserID.IsEnabled = False
				tbPassword.IsEnabled = False
				cmbInitialCatalog.SelectedIndex = 0
			Else
				ConnectionString = connectionString_1
			End If

			AddHandler cmbIntegratedSecurity.SelectionChanged, AddressOf cmbIntegratedSecurity_SelectedIndexChanged
		End Sub

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Function GetConnectionString() As String
            Dim builder = New SqlConnectionStringBuilder() With { _
                .ConnectionString = _connectionString, _
                .DataSource = tbDataSource.Text, _
                .IntegratedSecurity = (cmbIntegratedSecurity.SelectedIndex = 0), _
                .UserID = tbUserID.Text, _
                .Password = tbPassword.Password, _
                .InitialCatalog = If(cmbInitialCatalog.Text = "<default>", "", cmbInitialCatalog.Text) _
            }

            Try
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
            Dim builder = New SqlConnectionStringBuilder() With { _
                .ConnectionString = _connectionString _
            }

            Try
                tbDataSource.Text = builder.DataSource
                cmbIntegratedSecurity.SelectedIndex = If((builder.IntegratedSecurity), 0, 1)
                tbUserID.Text = builder.UserID
                tbUserID.IsEnabled = Not builder.IntegratedSecurity
                tbPassword.Password = builder.Password
                tbPassword.IsEnabled = Not builder.IntegratedSecurity
                cmbInitialCatalog.Text = builder.InitialCatalog

                _connectionString = builder.ConnectionString
            Catch
            End Try
        End Sub

        Public Function TestConnection() As Boolean Implements IConnectionFrame.TestConnection
            Mouse.OverrideCursor = Cursors.Wait

            Try
                Dim connection As New SqlConnection(ConnectionString)
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

        Private Sub cmbIntegratedSecurity_SelectedIndexChanged(sender As Object, e As EventArgs)
            tbUserID.IsEnabled = (cmbIntegratedSecurity.SelectedIndex = 1)
            tbPassword.IsEnabled = (cmbIntegratedSecurity.SelectedIndex = 1)
        End Sub

        Private Sub CmbInitialCatalog_OnDropDownOpened(sender As Object, e As EventArgs)

            Using connection = New SqlConnection(ConnectionString)
                Mouse.OverrideCursor = Cursors.Wait

                Dim currentDatabase = cmbInitialCatalog.Text

                cmbInitialCatalog.Items.Clear()
                cmbInitialCatalog.Items.Add("<default>")
                cmbInitialCatalog.SelectedIndex = 0

                Try
                    connection.Open()

                    Dim schemaTable = connection.GetSchema("Databases")

                    For Each r As DataRow In schemaTable.Rows
                        cmbInitialCatalog.Items.Add(r(0))
                    Next

                    cmbInitialCatalog.SelectedItem = Nothing
                    cmbInitialCatalog.SelectedItem = currentDatabase

                    If cmbInitialCatalog.SelectedItem Is Nothing Then
                        cmbInitialCatalog.Text = currentDatabase
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Connection Failure.")
                Finally
                    Mouse.OverrideCursor = Nothing
                End Try
            End Using
        End Sub

        Private Sub BtnEditConnectionString_OnClick(sender As Object, e As RoutedEventArgs)
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
	End Class
End Namespace
