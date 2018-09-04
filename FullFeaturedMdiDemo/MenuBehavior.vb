'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls

Public NotInheritable Class MenuBehavior
    Private Sub New()
    End Sub
    <AttachedPropertyBrowsableForType(GetType(MenuItem))> _
    Public Shared Function GetOptionGroupName(obj As MenuItem) As String
        Return DirectCast(obj.GetValue(OptionGroupNameProperty), String)
    End Function

    Public Shared Sub SetOptionGroupName(obj As MenuItem, value As String)
        obj.SetValue(OptionGroupNameProperty, value)
    End Sub

    Public Shared ReadOnly OptionGroupNameProperty As DependencyProperty = DependencyProperty.RegisterAttached("OptionGroupName", GetType(String), GetType(MenuBehavior), New UIPropertyMetadata(Nothing, AddressOf OptionGroupNameChanged))

    Private Shared Sub OptionGroupNameChanged(o As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim menuItem As MenuItem = TryCast(o, MenuItem)
        If menuItem Is Nothing Then
            Return
        End If

        Dim oldValue As String = DirectCast(e.OldValue, String)
        Dim newValue As String = DirectCast(e.NewValue, String)

        If Not String.IsNullOrEmpty(oldValue) Then
            RemoveFromOptionGroup(menuItem)
        End If
        If Not String.IsNullOrEmpty(newValue) Then
            AddToOptionGroup(menuItem)
        End If
    End Sub

    Private Shared Function GetOptionGroups(obj As DependencyObject) As Dictionary(Of String, HashSet(Of MenuItem))
        Return DirectCast(obj.GetValue(OptionGroupsPropertyKey.DependencyProperty), Dictionary(Of String, HashSet(Of MenuItem)))
    End Function

    Private Shared Sub SetOptionGroups(obj As DependencyObject, value As Dictionary(Of String, HashSet(Of MenuItem)))
        obj.SetValue(OptionGroupsPropertyKey, value)
    End Sub

    Private Shared ReadOnly OptionGroupsPropertyKey As DependencyPropertyKey = DependencyProperty.RegisterAttachedReadOnly("OptionGroups", GetType(Dictionary(Of String, HashSet(Of MenuItem))), GetType(MenuBehavior), New UIPropertyMetadata(Nothing))

    Private Shared Function GetOptionGroup(menuItem As MenuItem, create As Boolean) As HashSet(Of MenuItem)
        Dim groupName As String = GetOptionGroupName(menuItem)
        If groupName Is Nothing Then
            Return Nothing
        End If

        If menuItem.Parent Is Nothing Then
            Return Nothing
        End If

        Dim optionGroups = GetOptionGroups(menuItem.Parent)
        If optionGroups Is Nothing Then
            If create Then
                optionGroups = New Dictionary(Of String, HashSet(Of MenuItem))()
                SetOptionGroups(menuItem.Parent, optionGroups)
            Else
                Return Nothing
            End If
        End If

        Dim group As HashSet(Of MenuItem)
        If Not optionGroups.TryGetValue(groupName, group) AndAlso create Then
            group = New HashSet(Of MenuItem)()
            optionGroups(groupName) = group
        End If
        Return group
    End Function

    Private Shared Sub AddToOptionGroup(menuItem As MenuItem)
        Dim group = GetOptionGroup(menuItem, True)
        If group Is Nothing Then
            Return
        End If

        If Not group.Add(menuItem) Then
            Return
        End If

        AddHandler menuItem.Checked, AddressOf menuItem_Checked
        AddHandler menuItem.Unchecked, AddressOf menuItem_Unchecked
    End Sub

    Private Shared Sub RemoveFromOptionGroup(menuItem As MenuItem)
        Dim group = GetOptionGroup(menuItem, False)
        If group Is Nothing Then
            Return
        End If

        If group.Remove(menuItem) Then
            RemoveHandler menuItem.Checked, AddressOf menuItem_Checked
            RemoveHandler menuItem.Unchecked, AddressOf menuItem_Unchecked
        End If
    End Sub

    Private Shared Sub menuItem_Checked(sender As Object, e As RoutedEventArgs)
        Dim menuItem As MenuItem = TryCast(sender, MenuItem)
        If menuItem Is Nothing Then
            Return
        End If

        Dim groupName As String = GetOptionGroupName(menuItem)
        If groupName Is Nothing Then
            Return
        End If

        ' More than 1 checked option is allowed
        If groupName.EndsWith("*") OrElse groupName.EndsWith("+") Then
            Return
        End If

        Dim group = GetOptionGroup(menuItem, False)
        If group Is Nothing Then
            Return
        End If

        For Each item As MenuItem In group
            If CStr(item Is menuItem) Then Continue For

            item.IsChecked = False
        Next
    End Sub

    Private Shared Sub menuItem_Unchecked(sender As Object, e As RoutedEventArgs)
        Dim menuItem As MenuItem = TryCast(sender, MenuItem)
        If menuItem Is Nothing Then
            Return
        End If

        Dim groupName As String = GetOptionGroupName(menuItem)
        If groupName Is Nothing Then
            Return
        End If

        ' 0 checked option is allowed
        If groupName.EndsWith("*") OrElse groupName.EndsWith("?") Then
            Return
        End If

        Dim group = GetOptionGroup(menuItem, False)
        If group Is Nothing Then
            Return
        End If

        If Not Enumerable.Cast(Of MenuItem)(group).Any(Function(item As MenuItem) item.IsChecked = True) Then
            menuItem.IsChecked = True
        End If
    End Sub
End Class
