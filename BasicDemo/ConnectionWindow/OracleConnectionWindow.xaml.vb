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
Imports System.Windows.Controls
Imports System.Windows.Input

Namespace ConnectionWindow
	''' <summary>
	''' Interaction logic for OracleConnectionWindow.xaml
	''' </summary>
	Public Partial Class OracleConnectionWindow
		Public ConnectionString As String = ""

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub ButtonConnect_OnClick(sender As Object, e As RoutedEventArgs)
            Dim builder = New OracleConnectionStringBuilder() With { _
                .DataSource = textBoxServerName.Text _
            }


			builder.UserID = textBoxLogin.Text
			builder.Password = textBoxPassword.Password

			' check the connection

			Using connection = New OracleConnection(builder.ConnectionString)
				Mouse.OverrideCursor = Cursors.Wait

				Try
					connection.Open()
					ConnectionString = builder.ConnectionString
					DialogResult = True
					Close()
				Catch ex As System.Exception
					MessageBox.Show(ex.Message, "Connection Failure.")
				Finally
					Mouse.OverrideCursor = Nothing
				End Try
			End Using
		End Sub

		Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
			Close()
		End Sub
	End Class
End Namespace
