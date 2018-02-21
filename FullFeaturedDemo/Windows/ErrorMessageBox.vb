'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Windows
    Public Class ErrorMessageBox
        Private Shared _instance As ErrorMessageBox
        Private ReadOnly _popup As Popup
        Private ReadOnly _timerClose As Timer
        Private ReadOnly _textBlock As TextBlock

        Public Sub New()
            _textBlock = New TextBlock() With {
                .Background = Brushes.LightPink,
                .Padding = New Thickness(10)
            }

            _popup = New Popup() With {
                .StaysOpen = True,
                .Child = _textBlock,
                .Placement = PlacementMode.Relative,
                .AllowsTransparency = True,
                .IsOpen = True
            }
            AddHandler _popup.Opened, AddressOf PopupOnOpened
            _timerClose = New Timer(AddressOf ElepsedTime)
        End Sub

        Private Sub PopupOnOpened(sender As Object, e As EventArgs)
            Dim control As FrameworkElement = TryCast(_popup.PlacementTarget, FrameworkElement)

            If control Is Nothing Then
                Return
            End If

            _textBlock.Measure(New Size(Double.MaxValue, Double.MaxValue))
            _textBlock.Arrange(New Rect(New Point(0, 0), _instance._textBlock.DesiredSize))

            _popup.HorizontalOffset = -_instance._textBlock.ActualWidth - 6
            _popup.VerticalOffset = 6
        End Sub

        Private Sub ElepsedTime(state As Object)
            _popup.Dispatcher.Invoke(DispatcherPriority.Normal, New Action(Function()
                                                                               _popup.IsOpen = False
                                                                           End Function))
        End Sub

        Public Shared Sub Show(control As FrameworkElement, message As String)
            If _instance Is Nothing Then
                _instance = New ErrorMessageBox()
            End If

            If String.IsNullOrEmpty(message) Then
                _instance._timerClose.Change(Timeout.Infinite, Timeout.Infinite)
                _instance._popup.IsOpen = False
                Return
            End If

            If _instance._popup.IsOpen Then
                _instance._timerClose.Change(Timeout.Infinite, Timeout.Infinite)
                _instance._popup.IsOpen = False
            End If

        End Sub

    End Class
End Namespace
