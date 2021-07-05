//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Data.Common
Imports System.Windows

Namespace Windows.QueryInformationWindows
	''' <summary>
	''' Interaction logic for QueryParametersWindow.xaml
	''' </summary>
	Partial Public Class QueryParametersWindow
		Private ReadOnly _command As DbCommand
		Private ReadOnly _source As ObservableCollection(Of DataGridItem)

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(command As DbCommand)
			_source = New ObservableCollection(Of DataGridItem)()
			_command = command

			InitializeComponent()

			For i = 0 To _command.Parameters.Count - 1
				Dim p = _command.Parameters(i)

				_source.Add(New DataGridItem(p.ParameterName, p.DbType, p.Value))
			Next i
			GridData.ItemsSource = _source
		End Sub

		Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
			For i = 0 To _command.Parameters.Count - 1
				_command.Parameters(i).Value = _source(i).Value
			Next i

			DialogResult = True
		End Sub

		Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
			DialogResult = False
		End Sub
	End Class

	Friend Class DataGridItem
		Public Property ParameterName() As String
		Public Property DataType() As DbType
		Public Property Value() As Object

		Public Sub New(parametrName As String, dataType As DbType, value As Object)
			ParameterName = parametrName
			Me.DataType = dataType
			Me.Value = value
		End Sub
	End Class
End Namespace
