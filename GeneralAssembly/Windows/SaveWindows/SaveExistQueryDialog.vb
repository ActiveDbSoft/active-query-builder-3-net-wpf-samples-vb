''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Windows
Imports System.Windows.Controls

Namespace Windows.SaveWindows
    Public Class SaveExistQueryDialog
        Inherits WindowMessage

        Public Result? As Boolean

        Private Const WIDTH_BUTTON As Integer = 90

        Public Sub New()
            Title = "Save dialog"
            Text = "Save changes?"
            ContentAlignment = HorizontalAlignment.Center
            Result = Nothing

            Dim saveButton = New Button With {
                .Content = "Save",
                .Width = WIDTH_BUTTON,
                .IsDefault = True
            }
            AddHandler saveButton.Click, Sub()
                Result = True
                Close()
            End Sub

            Dim notSaveButton = New Button With {
                .Margin = New Thickness(5, 0, 5, 0),
                .Content = "Don't save",
                .Width = WIDTH_BUTTON,
                .IsCancel = True
            }
            AddHandler notSaveButton.Click, Sub()
                Result = False
                Close()
            End Sub

            Dim continueButton = New Button With {
                .Content = "Continue edit",
                .Width = WIDTH_BUTTON
            }
            AddHandler continueButton.Click, Sub()
                Close()
            End Sub

            Buttons.Add(saveButton)
            Buttons.Add(notSaveButton)
            Buttons.Add(continueButton)
        End Sub
    End Class
End Namespace
