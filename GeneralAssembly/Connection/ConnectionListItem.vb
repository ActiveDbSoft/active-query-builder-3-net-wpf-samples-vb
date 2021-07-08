''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Namespace Connection
	Public Class ConnectionListItem
		Implements INotifyPropertyChanged

		Private _name As String
		Private _type As String
		Private _tag As Object
		Private _userQueries As String

		Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

		Public Property Name() As String
			Set(value As String)
				_name = value
				OnPropertyChanged("Name")
			End Set
			Get
				Return _name
			End Get
		End Property

		Public Property Type() As String
			Set(value As String)
				_type = value
				OnPropertyChanged("Type")
			End Set
			Get
				Return _type
			End Get
		End Property

		Public Property Tag() As Object
			Set(value As Object)
				_tag = value
				OnPropertyChanged("Tag")
			End Set
			Get
				Return _tag
			End Get
		End Property

		Public Property UserQueries() As String
			Set(value As String)
				_userQueries = value
				OnPropertyChanged("UserQueries")
			End Set
			Get
				Return _userQueries
			End Get
		End Property

		Protected Overridable Sub OnPropertyChanged(propertyName As String)
			Dim handler = PropertyChangedEvent
			If handler IsNot Nothing Then
				handler(Me, New PropertyChangedEventArgs(propertyName))
			End If
		End Sub
	End Class
End Namespace
