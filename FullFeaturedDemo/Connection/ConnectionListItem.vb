'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel

Namespace Connection
	Public Class ConnectionListItem
		Implements INotifyPropertyChanged
		Private _name As String
		Private _type As String
		Private _tag As Object
		Private _userQueries As String

		Public Event PropertyChanged As PropertyChangedEventHandler

		Public Property Name() As String
			Get
				Return _name
			End Get
			Set
				_name = value
				OnPropertyChanged("Name")
			End Set
		End Property

		Public Property Type() As String
			Get
				Return _type
			End Get
			Set
				_type = value
				OnPropertyChanged("Type")
			End Set
		End Property

		Public Property Tag() As Object
			Get
				Return _tag
			End Get
			Set
				_tag = value
				OnPropertyChanged("Tag")
			End Set
		End Property

		Public Property UserQueries() As String
			Get
				Return _userQueries
			End Get
			Set
				_userQueries = value
				OnPropertyChanged("UserQueries")
			End Set
		End Property

		Protected Overridable Sub OnPropertyChanged(propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
		End Sub

        Public Event INotifyPropertyChanged_PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    End Class
End Namespace
