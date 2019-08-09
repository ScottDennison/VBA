VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeMapInOrderCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cursor As TreeMapNode

Public Function _
next_node() As TreeMapNode
    If (cursor Is Nothing) Then
        Set next_node = Nothing
        Exit Function
    End If
    If (Not cursor.right Is Nothing) Then
        Set cursor = leftmost(cursor.right)
    Else
        If (cursor.parent Is Nothing) Then
            Set cursor = Nothing
            Exit Function
        End If
        Do While (cursor Is cursor.parent.right)
            Set cursor = cursor.parent
            If (cursor.parent Is Nothing) Then Exit Do
        Loop
        Set cursor = cursor.parent
    End If
    Set next_node = cursor
End Function

Public Function _
prev_node() As TreeMapNode
    If (cursor Is Nothing) Then
        Set prev_node = Nothing
        Exit Function
    End If
    If (Not cursor.left Is Nothing) Then
        Set cursor = rightmost(cursor.left)
    Else
        If (cursor.parent Is Nothing) Then
            Set cursor = Nothing
            Exit Function
        End If
        Do While (cursor Is cursor.parent.left)
            Set cursor = cursor.parent
            If (cursor.parent Is Nothing) Then Exit Do
        Loop
        Set cursor = cursor.parent
    End If
    Set prev_node = cursor
End Function

Public Sub _
start(start_at As TreeMapNode)
    Set cursor = start_at
End Sub

Public Function _
first_node() As TreeMapNode
    Set cursor = leftmost(root(cursor))
    Set first_node = cursor
End Function

Public Function _
last_node() As TreeMapNode
    Set cursor = rightmost(root(cursor))
    Set last_node = cursor
End Function

Private Function _
root(N As TreeMapNode) As TreeMapNode
    Do While (Not N.parent Is Nothing)
        Set N = N.parent
    Loop
    Set root = N
End Function

Private Function _
leftmost(N As TreeMapNode) As TreeMapNode
    Do While (Not N.left Is Nothing)
        Set N = N.left
    Loop
    Set leftmost = N
End Function

Private Function _
rightmost(N As TreeMapNode) As TreeMapNode
    Do While (Not N.right Is Nothing)
        Set N = N.right
    Loop
    Set rightmost = N
End Function

