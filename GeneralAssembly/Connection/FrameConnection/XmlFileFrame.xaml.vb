'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System
Imports System.Reflection
Imports System.Windows
Imports Microsoft.Win32

Namespace Connection.FrameConnection
	''' <summary>
	''' Interaction logic for XmlFileFrame.xaml
	''' </summary>
	Partial Public Class XmlFileFrame
		Implements IConnectionFrame

		Private _openFileDialog1 As OpenFileDialog
		Public Sub New()
			InitializeComponent()
		End Sub

		Public Property ConnectionString() As String Implements IConnectionFrame.ConnectionString
			Get
				Return tbXmlFile.Text
			End Get
			Set(value As String)
				tbXmlFile.Text = value
			End Set
		End Property

		Public Event OnSyntaxProviderDetected As SyntaxProviderDetected Implements IConnectionFrame.OnSyntaxProviderDetected

		Public Sub SetServerType(serverType As String) Implements IConnectionFrame.SetServerType

		End Sub

		Public Sub New(xmlFilePath As String)
			InitializeComponent()

			tbXmlFile.Text = xmlFilePath
		End Sub

		Public Function TestConnection() As Boolean Implements IConnectionFrame.TestConnection
			If Not String.IsNullOrEmpty(ConnectionString) Then
				Return True
			End If

			MessageBox.Show("Invalid Xml file path.", System.Reflection.Assembly.GetEntryAssembly().GetName().Name)

			Return False
		End Function

		Private Sub BtnBrowse_Click(sender As Object, e As EventArgs)
			_openFileDialog1 = New OpenFileDialog With {.FileName = ConnectionString}

			If _openFileDialog1.ShowDialog().Equals(True) Then
				ConnectionString = _openFileDialog1.FileName
			End If
		End Sub
	End Class
End Namespace
