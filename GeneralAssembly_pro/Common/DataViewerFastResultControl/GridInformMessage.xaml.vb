//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System.Windows
Imports System.Windows.Input

Namespace Common.DataViewerFastResultControl
	Partial Public Class GridInformMessage
		Public Shared ReadOnly MessageProperty As DependencyProperty = DependencyProperty.Register("Message", GetType(String), GetType(GridInformMessage), New PropertyMetadata(Nothing, AddressOf PropertyChanged))

		Private Shared Sub PropertyChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
			Dim obj As GridInformMessage = TryCast(d, GridInformMessage)
			If obj Is Nothing Then
				Return
			End If

			obj.BlockMessage.Text = TryCast(e.NewValue, String)

			obj.Visibility = If(String.IsNullOrEmpty(obj.BlockMessage.Text), Visibility.Collapsed, Visibility.Visible)
		End Sub

		Public Property Message() As String
			Get
				Return CStr(GetValue(MessageProperty))
			End Get
			Set(value As String)
				SetValue(MessageProperty, value)
			End Set
		End Property
		Public Sub New()
			InitializeComponent()
			Visibility = Visibility.Collapsed
		End Sub

		Private Sub Close_Click(sender As Object, e As MouseButtonEventArgs)
			Message = ""
		End Sub
	End Class
End Namespace
