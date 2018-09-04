'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows

Namespace Common
	Public Class HeaderData
		Implements INotifyPropertyChanged
		Private _sorting As Sorting
		Private _title As String
		Private _counter As Integer

		Private _upArrowVisible As Visibility
		Private _downArrowVisible As Visibility
		Private _showSortBlock As Visibility

		Public Property UpArrowVisible() As Visibility
			Get
				Return _upArrowVisible
			End Get
			Private Set
				_upArrowVisible = value
				OnPropertyChanged("UpArrowVisible")
			End Set
		End Property

		Public Property DownArrowVisible() As Visibility
			Get
				Return _downArrowVisible
			End Get
			Private Set
				_downArrowVisible = value
				OnPropertyChanged("DownArrowVisible")
			End Set
		End Property

		Public Property ShowSortBlock() As Visibility
			Get
				Return _showSortBlock
			End Get
			Private Set
				_showSortBlock = value
				OnPropertyChanged("ShowSortBlock")
			End Set
		End Property

		Public Property Sorting() As Sorting


			Get
				Return _sorting
			End Get
			Set
				_sorting = value
				Select Case value
					Case Sorting.Asc
						ShowSortBlock = Visibility.Visible
						UpArrowVisible = Visibility.Visible
						DownArrowVisible = Visibility.Hidden
						Exit Select
					Case Sorting.Desc
						ShowSortBlock = Visibility.Visible
						UpArrowVisible = Visibility.Hidden
						DownArrowVisible = Visibility.Visible
						Exit Select
					Case Sorting.None
						ShowSortBlock = Visibility.Hidden
						Exit Select
					Case Else
						Throw New ArgumentOutOfRangeException("value", value, Nothing)
				End Select
				OnPropertyChanged("Sorting")
			End Set
		End Property

		Public Property Title() As String
			Get
				Return _title
			End Get
			Set
				_title = value
				OnPropertyChanged("Title")
			End Set
		End Property

		Public Property Counter() As Integer
			Get
				Return _counter
			End Get
			Set
				_counter = value
				OnPropertyChanged("Counter")
			End Set
		End Property

		Public Sub New()
			Sorting = Sorting.None
			Title = "Empty"
			Counter = 0
			ShowSortBlock = Visibility.Hidden
		End Sub


		Protected Overridable Sub OnPropertyChanged(<CallerMemberName> Optional propertyName As String = Nothing)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
		End Sub

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    End Class

	Public Enum Sorting
		Asc
		Desc
		None
	End Enum
End Namespace
