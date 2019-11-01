'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.ObjectModel
Imports System.Data
Imports System.Data.Common
Imports System.Windows

Namespace Windows

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

            For i As Integer = 0 To _command.Parameters.Count - 1
                Dim p As DbParameter = _command.Parameters(i)

                _source.Add(New DataGridItem(p.ParameterName, p.DbType, ""))
            Next
            GridData.ItemsSource = _source
        End Sub

        Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
            For i As Integer = 0 To _command.Parameters.Count - 1
                _command.Parameters(i).Value = _source(i).Value
            Next
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub
    End Class

    Class DataGridItem
        Public Property ParameterName As String

        Public Property DataType As DbType

        Public Property Value As Object

        Public Sub New(parametrName As String, dataType_1 As DbType, value_2 As Object)
            ParameterName = parametrName
            DataType = dataType_1
            Value = value_2
        End Sub
    End Class
End Namespace
