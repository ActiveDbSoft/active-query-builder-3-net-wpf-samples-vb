'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Linq
Imports System.Windows.Controls

Namespace Common
    Partial Friend Class CheckableCombobox
        Public Event ItemCheckStateChanged As EventHandler
        Public Overloads ReadOnly Property Items As ObservableCollection(Of SelectableItem)

        Public Sub New()
            Items = New ObservableCollection(Of SelectableItem)()
            InitializeComponent()
            ItemsSource = Items
            AddHandler Items.CollectionChanged, AddressOf Items_CollectionChanged
        End Sub

        Protected Overrides Sub OnSelectionChanged(e As SelectionChangedEventArgs)
            SelectedIndex = -1
        End Sub

        Private Sub Items_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add

                    For Each eNewItem As SelectableItem In e.NewItems
                        AddHandler eNewItem.INotifyPropertyChanged_PropertyChanged, AddressOf ItemPropertyChanged
                    Next

                Case NotifyCollectionChangedAction.Remove

                    For Each eNewItem As SelectableItem In e.OldItems
                        RemoveHandler eNewItem.INotifyPropertyChanged_PropertyChanged, AddressOf ItemPropertyChanged
                    Next

                Case NotifyCollectionChangedAction.Replace

                    For Each eNewItem As SelectableItem In e.NewItems
                        AddHandler eNewItem.INotifyPropertyChanged_PropertyChanged, AddressOf ItemPropertyChanged
                    Next

                    For Each eNewItem As SelectableItem In e.OldItems
                        RemoveHandler eNewItem.INotifyPropertyChanged_PropertyChanged, AddressOf ItemPropertyChanged
                    Next

                Case NotifyCollectionChangedAction.Move, NotifyCollectionChangedAction.Reset
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select
        End Sub

        Private Sub ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(SelectableItem.IsChecked) Then Return
            UpdateText()
            RaiseEvent ItemCheckStateChanged(Me, New EventArgs())
        End Sub

        Public Function IsItemChecked(i As Integer) As Boolean
            Return Items(i).IsChecked
        End Function

        Public Sub ClearCheckedItems()
            Dim checkableComboboxItem As SelectableItem
            For Each checkableComboboxItem In Items
                checkableComboboxItem.IsChecked = False
            Next
        End Sub

        Public Sub SetItemChecked(i As Integer, b As Boolean)
            Items(i).IsChecked = b
        End Sub

        Private Sub UpdateText()
            Dim list As List(Of SelectableItem) = Items.Where(Function(x) x.IsChecked).ToList()
            Text = String.Empty

            Dim i As Integer
            For i = 0 To list.Count - 1
                Text += (If(i = 0, "", ", ")) & list(i).Content.ToString()
            Next

            ToolTip = If(String.IsNullOrEmpty(Text), Nothing, Text)
        End Sub
    End Class

    Friend Class SelectableItem
        Implements INotifyPropertyChanged
        Private _isChecked As Boolean
        Private _content As Object

        Public Property IsChecked As Boolean
            Get
                Return _isChecked
            End Get
            Set(value As Boolean)
                _isChecked = value
                RaiseEvent INotifyPropertyChanged_PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(IsChecked)))
            End Set
        End Property

        Public Property Content As Object
            Get
                Return _content
            End Get
            Set
                _content = Value
                RaiseEvent INotifyPropertyChanged_PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Content)))
            End Set
        End Property

        Public Sub New()
        End Sub

        Public Sub New(content As Object)
            Me.Content = content
        End Sub

        Public Sub New(content As Object, isChecked As Boolean)
            Me.New(content)
            Me.IsChecked = isChecked
        End Sub

        Public Event INotifyPropertyChanged_PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    End Class
End Namespace
