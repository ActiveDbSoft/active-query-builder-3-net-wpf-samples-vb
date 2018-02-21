'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Data.OracleClient
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Namespace ConnectionWindows
	''' <summary>
	''' Interaction logic for OracleConnectionWindow.xaml
	''' </summary>
	Public Partial Class OracleConnectionWindow
		Public ConnectionString As String = ""

		Public Sub New()
			InitializeComponent()
				'1 - Oracle Server Authentication
			comboBoxAuthentication.SelectedIndex = 1
		End Sub

		Private Sub ButtonConnect_OnClick(sender As Object, e As RoutedEventArgs)
            Dim builder As New OracleConnectionStringBuilder() With { _
                .DataSource = textBoxServerName.Text _
            }

			If comboBoxAuthentication.SelectedIndex = 0 Then
				builder.IntegratedSecurity = True
			Else
				builder.IntegratedSecurity = False
				builder.UserID = textBoxLogin.Text
				builder.Password = textBoxPassword.Password
			End If

			' check the connection

			Using connection As New OracleConnection(builder.ConnectionString)
				Mouse.OverrideCursor = Cursors.Wait

				Try
					connection.Open()
					ConnectionString = builder.ConnectionString
				Catch ex As Exception
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

		Private Sub ComboBoxAuthentication_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
			If comboBoxAuthentication.SelectedIndex = 0 Then
				'Windows Authentication
				textBoxLogin.IsEnabled = False
				textBoxPassword.IsEnabled = False
			Else
				'Oracle Server Authentication
				textBoxLogin.IsEnabled = True
				textBoxPassword.IsEnabled = True
			End If
		End Sub
	End Class
End Namespace
