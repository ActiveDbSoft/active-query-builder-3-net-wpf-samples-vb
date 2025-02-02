''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Namespace Common
    Public Class AutoGreyableImage
        Inherits Image

        Shared Sub New()
            IsEnabledProperty.OverrideMetadata(GetType(AutoGreyableImage), New FrameworkPropertyMetadata(True, AddressOf OnAutoGreyScaleImageIsEnabledPropertyChanged))
        End Sub

        Private Shared Sub OnAutoGreyScaleImageIsEnabledPropertyChanged(source As DependencyObject, args As DependencyPropertyChangedEventArgs)
            Dim autoGreyScaleImg = TryCast(source, AutoGreyableImage)
            Dim isEnable = Convert.ToBoolean(args.NewValue)
            If autoGreyScaleImg Is Nothing Then
                Return
            End If

            If Not isEnable Then
                Dim bitmapImage = New BitmapImage(New Uri(autoGreyScaleImg.Source.ToString()))
                autoGreyScaleImg.Source = New FormatConvertedBitmap(bitmapImage, PixelFormats.Gray32Float, Nothing, 0)
                autoGreyScaleImg.OpacityMask = New ImageBrush(bitmapImage)
            Else
                autoGreyScaleImg.Source = CType(autoGreyScaleImg.Source, FormatConvertedBitmap).Source
                autoGreyScaleImg.OpacityMask = Nothing
            End If
        End Sub
    End Class
End Namespace
