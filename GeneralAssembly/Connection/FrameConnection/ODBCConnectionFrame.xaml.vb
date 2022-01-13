''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Data.Odbc
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Input
Imports ActiveQueryBuilder.Core

Namespace Connection.FrameConnection
    ''' <summary>
    ''' Interaction logic for ODBCConnectionFrame.xaml
    ''' </summary>
    Partial Public Class ODBCConnectionFrame
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

            Me.ConnectionString = connectionString
        End Sub

        Public Function GetConnectionString() As String
            Try
                Dim builder = New OdbcConnectionStringBuilder With {.ConnectionString = tbConnectionString.Text}
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
                MessageBox.Show(e.Message, System.Reflection.Assembly.GetEntryAssembly().GetName().Name)
                Return False
            Finally
                Mouse.OverrideCursor = Nothing
            End Try

            Return True
        End Function

        Public Sub DoSyntaxDetected(syntaxType As Type)
            RaiseEvent OnSyntaxProviderDetected(syntaxType)
        End Sub

        Private Sub btnTest_Click(sender As Object, e As RoutedEventArgs)
            Dim metadataProvider = New ODBCMetadataProvider With {.Connection = New OdbcConnection(ConnectionString)}
            Dim syntaxProviderType As Type = Nothing

            Try
                syntaxProviderType = Helpers.AutodetectSyntaxProvider(metadataProvider)
            Catch exception As Exception
                MessageBox.Show(exception.Message, "Error")
            End Try

            DoSyntaxDetected(syntaxProviderType)
        End Sub

        Private Sub tbConnectionString_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs)
            btnTest.IsEnabled = tbConnectionString.Text <> String.Empty
        End Sub
    End Class
End Namespace
