VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TreeMapDuplicationMode
    TreeMapDuplicationMode_Ignore
    TreeMapDuplicationMode_RaiseError
    TreeMapDuplicationMode_Overwrite
End Enum

Private m_root As TreeMapNode
Private m_duplication_mode As TreeMapDuplicationMode
Private m_count As Long
Private m_key_comparator As IVariantComparator

Private Sub Class_Initialize()
    m_duplication_mode = TreeMapDuplicationMode_RaiseError
    m_count = 0
    Set m_key_comparator = New SimpleVariantComparator
End Sub

Public Function _
add(key As Variant, value As Variant) As TreeMapNode
    Dim N As TreeMapNode
    'Debug.Print "add key " & key
    Set N = rb_insert(key)
    If Not (N Is Nothing) Then
        m_count = m_count + 1
        If (IsObject(value)) Then
            Set N.payload = value
        Else
            N.payload = value
        End If
        Set add = N
    End If
End Function

Public Function _
find(key As Variant) As TreeMapNode
    Dim at As TreeMapNode
    Set at = m_root
    Do While Not (at Is Nothing)
        Select Case m_key_comparator.compare(key, at.key)
            Case 0:
                Set find = at
                Exit Function
            Case -1:
                Set at = at.left
            Case 1:
                Set at = at.right
        End Select
    Loop
End Function

Public Property Get _
count() As Long
    count = m_count
End Property

Public Property Get _
duplication_mode() As TreeMapDuplicationMode
    duplication_mode = m_duplication_mode
End Property

Public Property Let _
duplication_mode(a_duplicaiton_mode As TreeMapDuplicationMode)
    m_duplication_mode = a_duplicaiton_mode
End Property

Public Property Get _
key_comparator() As IVariantComparator
    Set key_comparator = m_key_comparator
End Property

Public Property Set _
key_comparator(a_key_comparator As IVariantComparator)
    Set m_key_comparator = a_key_comparator
End Property

Public Function _
remove(key As Variant) As TreeMapNode
    Dim N As TreeMapNode
    Set N = find(key)
    If Not (N Is Nothing) Then
        remove_node N
        m_count = m_count - 1
    End If
    Set remove = N
End Function

Public Function _
create_in_order_cursor(Optional from_key As Variant = Nothing) As TreeMapInOrderCursor
    Dim inorder As New TreeMapInOrderCursor
    If Not (from_key Is Nothing) Then
        inorder.start find(from_key)
    Else
        inorder.start m_root
    End If
    Set create_in_order_cursor = inorder
End Function

Public Sub _
dump(path As String, Optional N As TreeMapNode)
    Open path For Output As #1
    do_dump N, 0
    Close #1
End Sub

Private Sub _
do_dump(Optional N As TreeMapNode, Optional depth As Long = 0)
    Dim spacer As String, M As TreeMapNode, last As Long, keystr As String
    Dim blacks As Long
    If (N Is Nothing) Then Set N = m_root
    If Not (N.left Is Nothing) Then
        do_dump N.left, depth + 1
    End If
    If (N.parent Is Nothing) Then
        spacer = "ROOT+ "
    ElseIf (N Is N.parent.left) Then
        last = 0
        spacer = "/---+ "
    Else
        last = 1
        spacer = "\---+ "
    End If
    Set M = N.parent
    blacks = IIf(N.color = 0, 3, 2)
    If Not (M Is Nothing) Then
        Do While Not M.parent Is Nothing:
            If (last = 0) <> (M Is M.parent.left) Then
                spacer = "|   " & spacer
            Else
                spacer = "    " & spacer
            End If
            last = IIf(M Is M.parent.left, 0, 1)
            If (M.color = 0) Then blacks = blacks + 1
            Set M = M.parent
        Loop
        If (N.left Is Nothing) And (N.right Is Nothing) Then
            spacer = Format(blacks, "0000") & spacer
        Else
            spacer = "----" & spacer
        End If
    End If
    If IsArray(N.key) Then
        keystr = "(" & N.key(LBound(N.key))
        For last = LBound(N.key) + 1 To UBound(N.key)
            keystr = keystr & "," & N.key(last)
        Next last
        keystr = keystr & ")"
    ElseIf IsObject(N.key) Then
        keystr = ObjPtr(N.key)
    Else
        keystr = CStr(N.key)
    End If
    Write #1, spacer & IIf(N.color = 0, "blk:", "red:") & keystr
    If Not (N.right Is Nothing) Then
        do_dump N.right, depth + 1
    End If
End Sub


' left rotation, ie.:
' (N N.left (M M.left M.right)) -> (M (N N.left M.left) M.right)
Private Sub left_rotate(N As TreeMapNode)
    Dim M As TreeMapNode
    Set M = N.right
    Set N.right = M.left
    If Not (M.left Is Nothing) Then
        Set M.left.parent = N
    End If
    Set M.parent = N.parent
    If (N.parent Is Nothing) Then
        Set m_root = M
    Else
        If (N Is N.parent.left) Then
            Set N.parent.left = M
        Else
            Set N.parent.right = M
        End If
    End If
    Set M.left = N
    Set N.parent = M
End Sub

' right rotation, ie.:
' (N (M M.left M.right) N.right) -> (M M.left (N M.right N.right))
Private Sub right_rotate(N As TreeMapNode)
    Dim M As TreeMapNode
    Set M = N.left
    Set N.left = M.right
    If Not (M.right Is Nothing) Then
        Set M.right.parent = N
    End If
    Set M.parent = N.parent
    If (N.parent Is Nothing) Then
        Set m_root = M
    Else
        If (N Is N.parent.left) Then
            Set N.parent.left = M
        Else
            Set N.parent.right = M
        End If
    End If
    Set M.right = N
    Set N.parent = M
End Sub

Private Function _
tree_insert(key As Variant) As TreeMapNode
    Dim at As TreeMapNode, last As TreeMapNode
    Set at = m_root: Set last = Nothing
    Do While Not (at Is Nothing)
        Set last = at
        Select Case m_key_comparator.compare(key, at.key)
            Case -1: Set at = at.left
            Case 1: Set at = at.right
            Case 0:
                Select Case (m_duplication_mode)
                    Case TreeMapDuplicationMode_RaiseError:
                        Err.Raise 9, "TreeMap.tree_insert()", "Key already exists"
                    Case TreeMapDuplicationMode_Ignore:
                        Exit Function
                    Case TreeMapDuplicationMode_Overwrite:
                        Set tree_insert = at
                        Exit Function
                End Select
        End Select
    Loop
    If (last Is Nothing) Then
        Set m_root = New TreeMapNode
        m_root.key = key
        Set tree_insert = m_root
        Exit Function
    Else
        Set at = New TreeMapNode
        Set at.parent = last
        at.key = key
        Set tree_insert = at
        Select Case m_key_comparator.compare(key, last.key)
            Case -1: Set last.left = at
            Case 1: Set last.right = at
            Case Else: Stop
        End Select
    End If
End Function

Private Function _
grandparent(N As TreeMapNode) As TreeMapNode
    Set grandparent = Nothing
    If Not (N.parent Is Nothing) Then
        Set grandparent = N.parent.parent
    End If
End Function

Private Function _
uncle(N As TreeMapNode) As TreeMapNode
    Dim gp As TreeMapNode
    Set gp = grandparent(N)
    Set uncle = Nothing
    If Not (gp Is Nothing) Then
        If (N.parent Is gp.left) Then
            Set uncle = gp.right
        Else
            Set uncle = gp.left
        End If
    End If
End Function

Private Function _
sibling(N As TreeMapNode) As TreeMapNode
    If (N Is N.parent.left) Then
        Set sibling = N.parent.right
    Else
        Set sibling = N.parent.left
    End If
End Function

Private Function _
rb_insert(key As Variant) As TreeMapNode
    Dim N As TreeMapNode, U As TreeMapNode, G As TreeMapNode

    Set N = tree_insert(key)
    If (N Is Nothing) Then Exit Function
    Set rb_insert = N

    ' node already existed and payload was overwritten by tree_insert()
    If (Not N.left Is Nothing) Or (Not N.right Is Nothing) Then Exit Function

    N.color = 1
    Do While True
        If (N Is m_root) Then Exit Do
        If (N.parent.color <> 1) Then Exit Do
        Set U = uncle(N)
        Set G = grandparent(N)
        Select Case True:
            Case U Is Nothing, U.color = 0:
                ' U in black
                If (N Is N.parent.right) And (N.parent Is G.left) Then
                    left_rotate N.parent
                    Set N = N.left
                ElseIf (N Is N.parent.left) And (N.parent Is G.right) Then
                    right_rotate N.parent
                    Set N = N.right
                End If
                N.parent.color = 0
                G.color = 1
                If (N Is N.parent.left) Then
                    right_rotate G
                Else
                    left_rotate G
                End If
            Case Else:
                ' U is red
                N.parent.color = 0
                U.color = 0
                G.color = 1
                Set N = G
        End Select
    Loop
    m_root.color = 0
End Function

Private Sub _
replace_node(dest As TreeMapNode, src As TreeMapNode)
    If (src Is Nothing) Then
        dest.key = Empty
        dest.payload = Empty
    Else
        dest.key = src.key
        dest.payload = Empty
        If (IsObject(src.payload)) Then
            Set dest.payload = src.payload
        Else
            dest.payload = src.payload
        End If
    End If
End Sub

Private Sub _
unlink_node(N As TreeMapNode)
    If (N Is Nothing) Then Exit Sub
    If (N.parent Is Nothing) Then
        Set m_root = Nothing
        Exit Sub
    End If
    If (N Is N.parent.left) Then
        Set N.parent.left = Nothing
    Else
        Set N.parent.right = Nothing
    End If
    Set N.parent = Nothing
End Sub

Private Function _
get_color(N As TreeMapNode) As Long
    If (N Is Nothing) Then
        get_color = 0
    Else
        get_color = N.color
    End If
End Function

Private Sub _
set_color(N As TreeMapNode, color As Long)
    If (N Is Nothing) Then
        If (color = 0) Then Exit Sub Else Stop 'trying to repaint a leaf red
    End If
    N.color = color
End Sub

Private Sub _
remove_node(N As TreeMapNode)
    Dim C As New TreeMapInOrderCursor, child As TreeMapNode, S As TreeMapNode

    If ((Not N.left Is Nothing) And (Not N.right Is Nothing)) Then
        Set C = New TreeMapInOrderCursor
        C.start N
        Set child = C.prev_node()
        replace_node N, child
        remove_node child
    Else
        'Me.dump
        If (N.left Is Nothing) Then Set child = N.right Else Set child = N.left
        replace_node N, child
        unlink_node child
        If (get_color(N) = 0) Then
            If (get_color(child) = 1) Then
                set_color child, 0
                If (IsEmpty(N.key)) Then unlink_node N
                Exit Sub
            Else
                ' by wikipedia :P http://en.wikipedia.org/wiki/Red%E2%80%93black_tree
                ' case #1: N is the new root -> exit sub
                Do While (Not N.parent Is Nothing)
                    ' case #2: S is red (-> N.parent, S.left, S.right are black)
                    Set S = sibling(N)
                    If (S.color = 1) Then
                        N.parent.color = 1
                        S.color = 0
                        If (N Is N.parent.left) Then
                            left_rotate N.parent
                        Else
                            right_rotate N.parent
                        End If
                    End If
                    ' case #3: S, Sl, Sr, P are black
                    Set S = sibling(N)
                    If ((N.parent.color = 0) And _
                        (S.color = 0) And _
                        (get_color(S.left) = 0) And _
                        (get_color(S.right) = 0)) Then
                        S.color = 1
                        Set child = N
                        Set N = N.parent
                        If (IsEmpty(child.key)) Then unlink_node child
                        'Stop ' will loop to case #1
                    Else
                        ' case #4
                        'Set S = sibling(N)
                        If ((N.parent.color = 1) And _
                            (S.color = 0) And _
                            (get_color(S.left) = 0) And _
                            (get_color(S.right) = 0)) Then
                            S.color = 1
                            N.parent.color = 0
                            If (IsEmpty(N.key)) Then unlink_node N
                            Exit Sub
                        End If
                        ' case #5
                        If (S.color = 0) Then
                            If ((N Is N.parent.left) And _
                                (get_color(S.right) = 0) And _
                                (get_color(S.left) = 1)) Then
                                S.color = 1
                                set_color S.left, 0
                                right_rotate S
                            ElseIf ((N Is N.parent.right) And _
                                (get_color(S.left) = 0) And _
                                (get_color(S.right) = 1)) Then
                                S.color = 1
                                set_color S.right, 0
                                left_rotate S
                            End If
                        End If
                        ' case #6
                        Set S = sibling(N)
                        S.color = N.parent.color
                        N.parent.color = 0
                        If (N Is N.parent.left) Then
                            set_color S.right, 0
                            left_rotate N.parent
                        Else
                            set_color S.left, 0
                            right_rotate N.parent
                        End If
                        If (IsEmpty(N.key)) Then unlink_node N
                        Exit Sub
                    End If
                Loop
            End If
            'If (IsEmpty(N.key)) Then unlink_node N
        Else
            'If (IsEmpty(N.key)) Then unlink_node N
        End If
        If (IsEmpty(N.key)) Then unlink_node N
    End If
End Sub
