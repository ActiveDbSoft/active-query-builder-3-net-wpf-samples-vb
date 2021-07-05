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
Imports System.Windows.Controls
Imports System.Windows.Media

Namespace Windows.SaveWindows
	Public Class SaveAsWindowDialog
		Inherits WindowMessage

		Public Enum ActionSave
			UserQuery
			File
			NotSave
			[Continue]
		End Enum

		Public Property Action() As ActionSave

		Private Const WIDTH_BUTTON As Integer = 120

		Public Sub New(nameQuery As String)
			Background = New SolidColorBrush(SystemColors.ControlColor)
			WindowStartupLocation = WindowStartupLocation.Manual
			SizeToContent = SizeToContent.WidthAndHeight

			ResizeMode = ResizeMode.NoResize
			ShowInTaskbar = False
			Title = "Save Dialog"

			Action = ActionSave.NotSave

			Dim border = New Border With {
				.BorderThickness = New Thickness(0),
				.BorderBrush = Brushes.Gray
			}

			Dim root = New StackPanel With {.Margin = New Thickness(10)}
			border.Child = root
			Content = border

			Dim message = New TextBlock With {
				.Text = "Save changes to the [" & nameQuery & "]?",
				.HorizontalAlignment = HorizontalAlignment.Center
			}

			Dim bottomStack = New StackPanel With {
				.Orientation = Orientation.Horizontal,
				.HorizontalAlignment = HorizontalAlignment.Center,
				.Margin = New Thickness(0, 10, 0, 0)
			}

			Dim buttonSaveFile = New Button With {
				.Width = WIDTH_BUTTON,
				.Content = "Save to file..."
			}
			AddHandler buttonSaveFile.Click, Sub()
				Action = ActionSave.File
				Close()
			End Sub

			Dim buttonSaveUserQuery = New Button With {
				.Width = WIDTH_BUTTON,
				.Content = "Save as User Query",
				.Margin = New Thickness(5, 0, 5, 0)
			}
			AddHandler buttonSaveUserQuery.Click, Sub()
				Action = ActionSave.UserQuery
				Close()
			End Sub

			Dim buttonNotSave = New Button With {
				.Width = WIDTH_BUTTON,
				.Content = "Don't save",
				.Margin = New Thickness(0, 0, 5, 0)
			}
			AddHandler buttonNotSave.Click, Sub()
				Action = ActionSave.NotSave
				Close()
			End Sub


			Dim buttonCancel = New Button With {
				.Width = WIDTH_BUTTON,
				.Content = "Cancel"
			}
			AddHandler buttonCancel.Click, Sub()
				Action = ActionSave.Continue
				Close()
			End Sub

			root.Children.Add(message)
			root.Children.Add(bottomStack)

			bottomStack.Children.Add(buttonSaveFile)
			bottomStack.Children.Add(buttonSaveUserQuery)
			bottomStack.Children.Add(buttonNotSave)
			bottomStack.Children.Add(buttonCancel)
		End Sub
	End Class
End Namespace
