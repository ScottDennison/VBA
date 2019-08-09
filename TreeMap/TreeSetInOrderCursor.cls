VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeSetInOrderCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private map_cursor As TreeMapInOrderCursor

Public Function _
next_entry() As Variant
    Dim next_node As TreeMapNode
    Set next_node = map_cursor.next_node()
    If next_node Is Nothing Then
        Set next_entry = Nothing
    ElseIf (IsObject(next_node.key)) Then
        Set next_entry = next_node.key
    Else
        next_entry = next_node.key
    End If
End Function

Public Function _
prev_entry() As Variant
    Dim prev_node As TreeMapNode
    Set prev_node = map_cursor.prev_node()
    If prev_node Is Nothing Then
        Set prev_entry = Nothing
    ElseIf (IsObject(prev_node.key)) Then
        Set prev_entry = prev_node.key
    Else
        prev_entry = prev_node.key
    End If
End Function

Public Sub _
start_using(map_cursor_ As TreeMapInOrderCursor)
    Set map_cursor = map_cursor_
End Sub

Public Function _
first_entry() As Variant
    Dim first_node As TreeMapNode
    Set first_node = map_cursor.first_node()
    If first_node Is Nothing Then
        Set first_entry = Nothing
    ElseIf (IsObject(first_node.key)) Then
        Set first_entry = first_node.key
    Else
        first_entry = first_node.key
    End If
End Function

Public Function _
last_entry() As Variant
    Dim last_node As TreeMapNode
    Set last_node = map_cursor.last_node()
    If last_node Is Nothing Then
        Set last_entry = Nothing
    ElseIf (IsObject(last_node.key)) Then
        Set last_entry = last_node.key
    Else
        last_entry = last_node.key
    End If
End Function
