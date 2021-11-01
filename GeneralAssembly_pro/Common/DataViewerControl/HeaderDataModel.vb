''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows

Namespace Common.DataViewerControl
	Public Class HeaderDataModel
		Implements INotifyPropertyChanged

		Private _sorting As Sorting
		Private _title As String
		Private _counter As Integer

		Private _upArrowVisible As Visibility
		Private _downArrowVisible As Visibility
		Private _showSortBlock As Visibility

		Public Property UpArrowVisible() As Visibility
			Private Set(value As Visibility)
				_upArrowVisible = value
				OnPropertyChanged("UpArrowVisible")
			End Set
			Get
				Return _upArrowVisible
			End Get
		End Property

		Public Property DownArrowVisible() As Visibility
			Private Set(value As Visibility)
				_downArrowVisible = value
				OnPropertyChanged("DownArrowVisible")
			End Set
			Get
				Return _downArrowVisible
			End Get
		End Property

		Public Property ShowSortBlock() As Visibility
			Private Set(value As Visibility)
				_showSortBlock = value
				OnPropertyChanged("ShowSortBlock")
			End Set
			Get
				Return _showSortBlock
			End Get
		End Property

		Public Property Sorting() As Sorting
			Set(value As Sorting)
				_sorting = value

				Select Case value
					Case Sorting.Asc
						ShowSortBlock = Visibility.Visible
						UpArrowVisible = Visibility.Visible
						DownArrowVisible = Visibility.Hidden
					Case Sorting.Desc
						ShowSortBlock = Visibility.Visible
						UpArrowVisible = Visibility.Hidden
						DownArrowVisible = Visibility.Visible
					Case Sorting.None
						ShowSortBlock = Visibility.Hidden
					Case Else
						Throw New ArgumentOutOfRangeException("value", value, Nothing)
				End Select

				OnPropertyChanged("Sorting")
			End Set
			Get
				Return _sorting
			End Get
		End Property

		Public Property Title() As String
			Set(value As String)
				_title = value
				OnPropertyChanged("Title")
			End Set
			Get
				Return _title
			End Get
		End Property

		Public Property Counter() As Integer
			Set(value As Integer)
				_counter = value
				OnPropertyChanged("Counter")
			End Set
			Get
				Return _counter
			End Get
		End Property

		Public Sub New()
			Sorting = Sorting.None
			Title = "Empty"
			Counter = 0
			ShowSortBlock = Visibility.Hidden
		End Sub

		Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Protected Overridable Overloads Sub OnPropertyChanged(<CallerMemberName> Optional propertyName As String = Nothing)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub
	End Class

	Public Enum Sorting
		Asc
		Desc
		None
	End Enum
End Namespace
