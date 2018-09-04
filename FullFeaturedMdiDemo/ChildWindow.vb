'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

#Region "Usings"

Imports System.Windows
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.View.QueryView
Imports FullFeaturedMdiDemo.Common
Imports FullFeaturedMdiDemo.MdiControl
Imports Helpers = FullFeaturedMdiDemo.Common.Helpers

#End Region

Public Class ChildWindow
    Inherits MdiChildWindow

    Public Event SaveQueryEvent As EventHandler   

    Public Event SaveAsInFileEvent As EventHandler
      
    Public Event SaveAsNewUserQueryEvent As EventHandler

    Public ReadOnly Property QueryView() As IQueryView
        Get
            Return _content.QueryView
        End Get
    End Property

    Public ReadOnly Property SqlQuery() As SQLQuery
        Get
            Return _content.SqlQuery
        End Get
    End Property

    Public Property SqlFormattingOptions() As SQLFormattingOptions
        Get
            Return _content.SqlFormattingOptions
        End Get
        Set
            _content.SqlFormattingOptions = Value
        End Set
    End Property

    Public Property SqlGenerationOptions() As SQLGenerationOptions
        Get
            Return _content.SqlGenerationOptions
        End Get
        Set
            _content.SqlGenerationOptions = Value
        End Set
    End Property

    Public Property FileSourceUrl() As String
        Get
            Return _content.FileSourceUrl
        End Get
        Set
            _content.FileSourceUrl = Value
        End Set
    End Property

    Public Property QueryText() As String
        Get
            Return _content.QueryText
        End Get
        Set
            _content.QueryText = Value
        End Set
    End Property

    Public Property SqlSourceType() As Helpers.SourceType
        Get
            Return _content.SqlSourceType
        End Get
        Set
            _content.SqlSourceType = Value
        End Set
    End Property

    Public Property UserMetadataStructureItem() As MetadataStructureItem
        Get
            Return _content.UserMetadataStructureItem
        End Get
        Set
            _content.UserMetadataStructureItem = Value
        End Set
    End Property

    Public Property IsNeedClose() As Boolean
        Get
            Return _content.IsNeedClose
        End Get
        Set
            _content.IsNeedClose = Value
        End Set
    End Property

    Public ReadOnly Property FormattedQueryText() As String
        Get
            Return _content.FormattedQueryText
        End Get
    End Property

    Public Property IsModified() As Boolean
        Get
            Return _content.IsModified
        End Get
        Set
            _content.IsModified = Value
        End Set
    End Property

    Private ReadOnly _content As ContentWindowChild

    Public Sub New(sqlContext As SQLContext)
        _content = New ContentWindowChild(sqlContext)
        Children.Add(_content)

        AddHandler Loaded, Sub()
                               If Double.IsNaN(Width) Then
                                   Width = ActualWidth
                               End If
                               If Double.IsNaN(Height) Then
                                   Height = ActualHeight
                               End If

                           End Sub

        AddHandler _content.SaveQueryEvent, AddressOf  RaiseSaveQueryEvent
        AddHandler _content.SaveAsInFileEvent,AddressOf  RaiseSaveAsInFileEvent
        AddHandler _content.SaveAsNewUserQueryEvent,AddressOf  RaiseSaveAsNewUserQueryEvent
    End Sub

    Private Function RaiseSaveAsNewUserQueryEvent() As EventHandler
        RaiseEvent SaveAsNewUserQueryEvent(_content, Nothing)
    End Function

    Private Function RaiseSaveAsInFileEvent() As EventHandler
        RaiseEvent SaveAsInFileEvent(_content, Nothing)
    End Function

    Private Function RaiseSaveQueryEvent() As EventHandler
       RaiseEvent SaveQueryEvent(_content, Nothing)
    End Function


    Public Function CanRedo() As Boolean
        Return _content.CanRedo()
    End Function

    Public Function CanCopy() As Boolean
        Return _content.CanCopy()
    End Function

    Public Function CanPaste() As Boolean
        Return _content.CanPaste()
    End Function

    Public Function CanCut() As Boolean
        Return _content.CanCut()
    End Function

    Public Function CanSelectAll() As Boolean
        Return _content.CanSelectAll()
    End Function

    Public Function CanAddDerivedTable() As Boolean
        Return _content.CanAddDerivedTable()
    End Function

    Public Function CanCopyUnionSubQuery() As Boolean
        Return _content.CanCopyUnionSubQuery()
    End Function

    Public Function CanAddUnionSubQuery() As Boolean
        Return _content.CanAddUnionSubQuery()
    End Function

    Public Function CanShowProperties() As Boolean
        Return _content.CanShowProperties()
    End Function

    Public Function CanAddObject() As Boolean
        Return _content.CanAddObject()
    End Function

    Public Sub AddDerivedTable()
        _content.AddDerivedTable()
    End Sub

    Public Sub CopyUnionSubQuery()
        _content.CopyUnionSubQuery()
    End Sub

    Public Sub AddUnionSubQuery()
        _content.AddUnionSubQuery()
    End Sub

    Public Sub PropertiesQuery()
        _content.PropertiesQuery()
    End Sub

    Public Sub AddObject()
        _content.AddObject()
    End Sub

    Public Sub ShowQueryStatistics()
        _content.ShowQueryStatistics()
    End Sub

    Public Sub Undo()
        _content.Undo()
    End Sub

    Public Sub Redo()
        _content.Redo()
    End Sub

    Public Sub Copy()
        _content.Copy()
    End Sub

    Public Sub Paste()
        _content.Paste()
    End Sub

    Public Sub Cut()
        _content.Cut()
    End Sub

    Public Sub SelectAll()
        _content.SelectAll()
    End Sub

    Public Sub OpenExecuteTab()
        _content.OpenExecuteTab()
    End Sub

    Public Sub ForceClose()
        MyBase.Close()
    End Sub

    Public Overrides Sub Close()
        If IsNeedClose OrElse Not IsModified Then
            MyBase.Close()
            Return
        End If

        Dim point As Point = PointToScreen(New Point(0, 0))

        If SqlSourceType = Helpers.SourceType.[New] Then
            IsNeedClose = True

            Dim dialog As SaveAsWindowDialog = New SaveAsWindowDialog(Title) With {
                .Left = point.X,
                .Top = point.Y
            }

            AddHandler dialog.Loaded, AddressOf Dialog_Loaded
            dialog.ShowDialog()

            Select Case dialog.Action
                Case SaveAsWindowDialog.ActionSave.UserQuery
                    _content.OnSaveAsNewUserQueryEvent()
                    Exit Select
                Case SaveAsWindowDialog.ActionSave.File
                    _content.OnSaveAsInFileEvent()
                    Exit Select
                Case SaveAsWindowDialog.ActionSave.NotSave
                    MyBase.Close()
                    Exit Select
                Case SaveAsWindowDialog.ActionSave.[Continue]
                    IsNeedClose = False
                    Exit Select
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select
        Else
            Dim saveDialog As SaveExistQueryDialog = New SaveExistQueryDialog() With {
                .Left = point.X,
                .Top = point.Y
            }

            AddHandler saveDialog.Loaded, AddressOf Dialog_Loaded
            saveDialog.ShowDialog()

            If Not saveDialog.Result.HasValue Then
                Return
            End If

            IsNeedClose = True

            If saveDialog.Result = True Then
                _content.OnSaveQueryEvent()
                Return
            End If
        End If

        MyBase.Close()
    End Sub

    Private Sub Dialog_Loaded(sender As Object, e As RoutedEventArgs)
        Dim dialog As Window = TryCast(sender, Window)

        If dialog Is Nothing Then
            Return
        End If

        RemoveHandler dialog.Loaded, AddressOf Dialog_Loaded

        dialog.Left += (ActualWidth / 2 - dialog.ActualWidth / 2)
        dialog.Top += (ActualHeight / 2 - dialog.ActualHeight / 2)
    End Sub

    Public Function CanUndo() As Boolean
        Return _content.CanUndo()
    End Function
End Class
