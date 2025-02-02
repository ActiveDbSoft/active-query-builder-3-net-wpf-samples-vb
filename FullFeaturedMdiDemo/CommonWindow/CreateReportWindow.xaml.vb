''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports FullFeaturedMdiDemo.Common

Namespace CommonWindow
    Partial Public Class CreateReportWindow
        Public Shared ReadOnly SelectedReportTypeProperty As DependencyProperty = DependencyProperty.Register("SelectedReportType", GetType(ReportType?), GetType(CreateReportWindow), New PropertyMetadata(Nothing))

        Public Property SelectedReportType() As ReportType?
            Get
                Return DirectCast(GetValue(SelectedReportTypeProperty), ReportType?)
            End Get
            Set(value? As ReportType)
                SetValue(SelectedReportTypeProperty, value)
            End Set
        End Property
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub ReportButton_OnChecked(sender As Object, e As RoutedEventArgs)
            If RadioButtonActiveReports.IsChecked = True Then
                SelectedReportType = ReportType.ActiveReports14
                Return
            End If

            If RadioButtonStimulsoft.IsChecked = True Then
                SelectedReportType = ReportType.Stimulsoft
                Return
            End If

            If RadioButtonFastReport.IsChecked = True Then
                SelectedReportType = ReportType.FastReport
            End If
        End Sub

        Private Sub ButtonBase_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
        End Sub
    End Class
End Namespace
