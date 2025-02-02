''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Data.OleDb
Imports System.Windows
Imports System.Windows.Input
Imports Microsoft.Win32

Namespace ConnectionWindows
    ''' <summary>
    ''' Interaction logic for AccessConnectionWindow.xaml
    ''' </summary>
    Public Partial Class AccessConnectionWindow
        Public ConnectionString As String = ""
        Private _openFileDialog1 As OpenFileDialog

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub ButtonConnect_OnClick(sender As Object, e As RoutedEventArgs)
            ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0"
            ConnectionString += ";Data Source=" & textboxDatabase.Text
            ConnectionString += ";User Id=" & textboxUserName.Text
            ConnectionString += ";Password="

            If textboxPassword.Password.Length > 0 Then
                ConnectionString += textboxPassword.Password & ";"
            End If

            ' check the connection

            Using connection As New OleDbConnection(ConnectionString)
                Mouse.OverrideCursor = Cursors.Wait

                Try
                    connection.Open()
                Catch ex As System.Exception
                    MessageBox.Show(ex.Message, "Connection Failure.")
                    DialogResult = True
                Finally
                    Mouse.OverrideCursor = Nothing
                End Try
            End Using
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = False
        End Sub

        Private Sub ButtonOpenDb_OnClick(sender As Object, e As RoutedEventArgs)
            _openFileDialog1 = New OpenFileDialog()
            If _openFileDialog1.ShowDialog() = True Then
                textboxDatabase.Text = _openFileDialog1.FileName
            End If
        End Sub
    End Class
End Namespace
