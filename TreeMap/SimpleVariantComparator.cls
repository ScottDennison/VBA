VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleVariantComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVariantComparator

Private Function min_long(val1 As Long, val2 As Long) As Long
    If val1 < val2 Then
        min_long = val1
    Else
        min_long = val2
    End If
End Function

Public Function _
IVariantComparator_compare(v1 As Variant, v2 As Variant) As Long
    Dim i As Long, offset As Long
    If (IsArray(v1)) Then
        If (IsArray(v2)) Then
            offset = LBound(v2) - LBound(v1)
            For i = LBound(v1) To min_long(UBound(v1), UBound(v2) + offset)
                If (v1(i) < v2(i + offset)) Then
                    IVariantComparator_compare = -1
                    Exit Function
                ElseIf (v1(i) > v2(i + offset)) Then
                    IVariantComparator_compare = 1
                    Exit Function
                End If
            Next i
            If (UBound(v1) < UBound(v2) + offset) Then
                IVariantComparator_compare = -1
            ElseIf (UBound(v1) > UBound(v2) + offset) Then
                IVariantComparator_compare = 1
            Else
                IVariantComparator_compare = 0
            End If
        Else
            IVariantComparator_compare = -1
        End If
    Else
        If (IsArray(v2)) Then
            IVariantComparator_compare = 1
        Else
            If (v1 < v2) Then
                IVariantComparator_compare = -1
            ElseIf (v1 = v2) Then
                IVariantComparator_compare = 0
            Else
                IVariantComparator_compare = 1
            End If
        End If
    End If
End Function
