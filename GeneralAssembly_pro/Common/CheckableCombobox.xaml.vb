''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel

Namespace Common
    ''' <summary>
    ''' Interaction logic for CheckableCombobox.xaml
    ''' </summary>
    Partial Public Class CheckableCombobox
        Public Event ItemCheckStateChanged As EventHandler
        Public Shadows Property Items As ObservableCollection(Of SelectableItem)

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
                        AddHandler eNewItem.PropertyChanged, AddressOf ItemPropertyChanged
                    Next eNewItem

                Case NotifyCollectionChangedAction.Remove
                    For Each eNewItem As SelectableItem In e.OldItems
                        RemoveHandler eNewItem.PropertyChanged, AddressOf ItemPropertyChanged
                    Next eNewItem
                Case NotifyCollectionChangedAction.Replace
                    For Each eNewItem As SelectableItem In e.NewItems
                        AddHandler eNewItem.PropertyChanged, AddressOf ItemPropertyChanged
                    Next eNewItem
                    For Each eNewItem As SelectableItem In e.OldItems
                        RemoveHandler eNewItem.PropertyChanged, AddressOf ItemPropertyChanged
                    Next eNewItem
                Case NotifyCollectionChangedAction.Move, NotifyCollectionChangedAction.Reset
                Case Else
                    Throw New ArgumentOutOfRangeException()
            End Select

            Text = String.Empty

            For Each item In Items
                If Not item.IsChecked Then
                    Continue For
                End If

                If Not String.IsNullOrEmpty(Text) Then
                    Text &= ", "
                End If
                Text &= item.Content.ToString()
            Next item
        End Sub

        Private Sub ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(SelectableItem.IsChecked) Then
                Return
            End If
            UpdateText()

            OnItemCheckStateChanged()
        End Sub

        Protected Overridable Sub OnItemCheckStateChanged()
            ItemCheckStateChangedEvent?.Invoke(Me, EventArgs.Empty)
        End Sub

        Public Function IsItemChecked(i As Integer) As Boolean
            Return Items(i).IsChecked
        End Function

        Public Sub ClearCheckedItems()
            For Each checkableComboboxItem In Items
                checkableComboboxItem.IsChecked = False
            Next checkableComboboxItem
        End Sub

        Public Sub SetItemChecked(i As Integer, b As Boolean)
            Items(i).IsChecked = b
        End Sub

        Private Sub UpdateText()
            Dim list = Items.Where(Function(x) x.IsChecked).ToList()

            Text = String.Empty

            For i = 0 To list.Count - 1
                Text &= (If(i = 0, "", ", ")) + list(i).Content.ToString()
            Next i

            ToolTip = If(String.IsNullOrEmpty(Text), Nothing, Text)
        End Sub
    End Class

    Public Class SelectableItem
        Implements INotifyPropertyChanged

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Private _isChecked As Boolean
        Private _content As Object

        Public Property IsChecked() As Boolean
            Get
                Return _isChecked
            End Get
            Set(value As Boolean)
                _isChecked = value
                OnPropertyChanged(NameOf(IsChecked))
            End Set
        End Property

        Public Property Content() As Object
            Get
                Return _content
            End Get
            Set(value As Object)
                _content = value
                OnPropertyChanged(NameOf(Content))
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

        Protected Overridable Sub OnPropertyChanged(propertyName As String)
            PropertyChangedEvent?.Invoke(Me, New PropertyChangedEventArgs(propertyName))
        End Sub
    End Class
End Namespace
