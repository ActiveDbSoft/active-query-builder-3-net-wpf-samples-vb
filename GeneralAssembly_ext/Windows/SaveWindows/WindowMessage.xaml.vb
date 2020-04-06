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
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Imaging

Namespace Windows.SaveWindows
	''' <summary>
	''' Interaction logic for WindowMessage.xaml
	''' </summary>
	Partial Public Class WindowMessage
		Implements INotifyPropertyChanged

		Public Property ContentAlignment() As HorizontalAlignment
			Get
				Return TextBlockContent.HorizontalAlignment
			End Get
			Set(value As HorizontalAlignment)
				TextBlockContent.HorizontalAlignment = value
			End Set
		End Property
		Public Property Text() As String
			Get
				Return TextBlockContent.Text
			End Get
			Set(value As String)
				TextBlockContent.Text = value
			End Set
		End Property

        Public Property Buttons As ObservableCollection(Of Button)

        Public Sub New()
			Buttons = New ObservableCollection(Of Button)()
			InitializeComponent()
			Icon = GetImageSource(My.Resources.disk)
			AddHandler Buttons.CollectionChanged, AddressOf Buttons_CollectionChanged
		End Sub

		Public Shared Function GetImageSource(image As Icon) As BitmapImage
			If image Is Nothing Then
				Return Nothing
			End If

			Using memory = New MemoryStream()
				image.Save(memory)
				memory.Position = 0

				Dim bitmapImage = New BitmapImage()
				bitmapImage.BeginInit()
				bitmapImage.StreamSource = memory
				bitmapImage.CacheOption = BitmapCacheOption.OnLoad
				bitmapImage.EndInit()

				Return bitmapImage
			End Using
		End Function

		Private Sub Buttons_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
			Select Case e.Action
				Case NotifyCollectionChangedAction.Add

					For Each button As Button In e.NewItems
						PlaceButtons.Children.Add(button)
					Next button

				Case NotifyCollectionChangedAction.Remove
					For Each button As Button In e.OldItems
						PlaceButtons.Children.Remove(button)
					Next button
				Case NotifyCollectionChangedAction.Reset
					PlaceButtons.Children.Clear()
					For Each button In Buttons
						PlaceButtons.Children.Add(button)
					Next button
				Case Else
					Throw New ArgumentOutOfRangeException()
			End Select

			OnPropertyChanged("Buttons")
		End Sub

		Public Shadows Function Show() As Boolean
			Return ShowDialog().Equals(True)
		End Function


		Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Protected Overridable Overloads Sub OnPropertyChanged(<CallerMemberName> Optional propertyName As String = Nothing)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub
	End Class
End Namespace
