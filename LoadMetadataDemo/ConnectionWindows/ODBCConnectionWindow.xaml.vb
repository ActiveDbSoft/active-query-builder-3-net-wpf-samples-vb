''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Data.Odbc
Imports System.Windows
Imports System.Windows.Input

Namespace ConnectionWindows
    ''' <summary>
    ''' Interaction logic for ODBCConnectionWindow.xaml
    ''' </summary>
    Public Partial Class ODBCConnectionWindow
        Public ConnectionString As String = ""

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub ButtonConnect_OnClick(sender As Object, e As RoutedEventArgs)
            Dim builder As OdbcConnectionStringBuilder = New OdbcConnectionStringBuilder()

            Try
                builder.ConnectionString = textBoxConnectionString.Text

                Mouse.OverrideCursor = Cursors.Wait

                Using connection As New OdbcConnection(builder.ConnectionString)
                    Try
                        connection.Open()
                        ConnectionString = builder.ConnectionString
                        DialogResult = True
                    Catch ex As System.Exception
                        MessageBox.Show(ex.Message, "Failed to connect.")
                    Finally
                        Mouse.OverrideCursor = Nothing
                    End Try
                End Using
            Catch ae As ArgumentException
                MessageBox.Show(ae.Message, "Invalid OLE DB connection string.")
            End Try
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = False
        End Sub
    End Class
End Namespace
