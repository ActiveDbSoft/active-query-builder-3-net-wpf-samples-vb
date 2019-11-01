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
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Controls

Namespace Common
	''' <summary>
	''' Interaction logic for WindowMessage.xaml
	''' </summary>
	Public Partial Class WindowMessage
		Implements INotifyPropertyChanged
		Public Property ContetnAlignment() As HorizontalAlignment
			Get
				Return TextBlockContent.HorizontalAlignment
			End Get
			Set
				TextBlockContent.HorizontalAlignment = value
			End Set
		End Property
		Public Property Text() As String
			Get
				Return TextBlockContent.Text
			End Get
			Set
				TextBlockContent.Text = value
			End Set
		End Property
        Public Property Buttons As ObservableCollection(Of Button)
            Get
                Return m_Buttons
            End Get
            Private Set(value As ObservableCollection(Of Button))
                m_Buttons = Value
            End Set
        End Property
		Private m_Buttons As ObservableCollection(Of Button)

		Public Sub New()
			Buttons = New ObservableCollection(Of Button)()
			InitializeComponent()
			AddHandler Buttons.CollectionChanged, AddressOf Buttons_CollectionChanged
		End Sub

		Private Sub Buttons_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
			Select Case e.Action
				Case NotifyCollectionChangedAction.Add

					For Each button As Button In e.NewItems
						PlaceButtons.Children.Add(button)
					Next

					Exit Select
				Case NotifyCollectionChangedAction.Remove
					For Each button As Button In e.OldItems
						PlaceButtons.Children.Remove(button)
					Next
					Exit Select
				Case NotifyCollectionChangedAction.Reset
					PlaceButtons.Children.Clear()
                    For Each button As Button In Buttons
                        PlaceButtons.Children.Add(button)
                    Next
					Exit Select
				Case Else
					Throw New ArgumentOutOfRangeException()
			End Select

			OnPropertyChanged("Buttons")
		End Sub

		Public Shadows Function Show() As Boolean
            Return CBool(ShowDialog() = True)
        End Function

        Protected Overridable Overloads Sub OnPropertyChanged(<CallerMemberName> Optional propertyName As String = Nothing)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    End Class
End Namespace
