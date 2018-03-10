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

Namespace ConnectionWindow
	''' <summary>
	''' Interaction logic for AccessConnectionWindow.xaml
	''' </summary>
	Public Partial Class AccessConnectionWindow
		Public ConnectionString As String = ""

		Private ReadOnly _openFileDialog As New OpenFileDialog()

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub ButtonConnect_OnClick(sender As Object, e As RoutedEventArgs)
			ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
			ConnectionString += ";Data Source=" & textboxDatabase.Text
			ConnectionString += ";User Id=" & textboxUserName.Text
			ConnectionString += ";Password="

			If textboxPassword.Password.Length > 0 Then
				ConnectionString += textboxPassword.Password & ";"
			End If

			' check the connection

			Using connection = New OleDbConnection(ConnectionString)
				Mouse.OverrideCursor = Cursors.Wait

				Try
					connection.Open()
					DialogResult = True
					Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Connection Failure.")
				Finally
					Mouse.OverrideCursor = Nothing
				End Try
			End Using
		End Sub

		Private Sub ButtonBrowse_OnClick(sender As Object, e As RoutedEventArgs)
			If _openFileDialog.ShowDialog() = True Then
				textboxDatabase.Text = _openFileDialog.FileName
			End If
		End Sub

		Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
			Close()
		End Sub
	End Class
End Namespace
