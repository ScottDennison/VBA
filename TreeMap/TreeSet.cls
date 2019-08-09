VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private map As New TreeMap

Private Sub Class_Initialize()
    map.duplication_mode = TreeMapDuplicationMode_Overwrite
End Sub

Public Function _
add(entry As Variant) As Boolean
    add = map.add(entry, Empty) Is Nothing
End Function

Public Function _
find(entry As Variant) As Boolean
    find = Not (map.find(entry) Is Nothing)
End Function

Public Property Get _
count() As Long
    count = map.count
End Property

Public Property Get _
entry_comparator() As IVariantComparator
    Set entry_comparator = map.key_comparator
End Property

Public Property Set _
entry_comparator(a_entry_comparator As IVariantComparator)
    Set map.key_comparator = a_entry_comparator
End Property

Public Function _
remove(entry As Variant) As Boolean
    remove = Not (map.remove(entry) Is Nothing)
End Function

Public Function _
create_in_order_cursor(Optional from_entry As Variant = Nothing) As TreeSetInOrderCursor
    Dim inorder As New TreeSetInOrderCursor
    inorder.start_using (map.create_in_order_cursor(from_entry))
    Set create_in_order_cursor = inorder
End Function

Public Sub _
dump(path As String, Optional N As Variant)
    Dim Nnode As TreeMapNode
    If N Is Nothing Then
        map.dump path
    Else
        Set Nnode = map.find(N)
        If Nnode Is Nothing Then
            Err.Raise 10, "TreeSet.dump()", "No such key"
        End If
        map.dump path, Nnode
    End If
End Sub
