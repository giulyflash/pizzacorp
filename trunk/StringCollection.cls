VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Simple String collection Class Module
'Jeff Erbrecht, March 30th, 2007

Private S() As String 'Internal string array

Public Sub Add(Item As String)
    'An error will occur if the S array has not been dimensioned yet
    On Error GoTo NotInitializedYet
    
    'Resize S() and store Item to its latest index
    ReDim Preserve S(0 To UBound(S) + 1)
    S(UBound(S)) = Item
    Exit Sub
    
NotInitializedYet:
    'Dimension the S array
    ReDim S(0 To 0)
    S(0) = Item 'Store Item to the first index of S
    
End Sub

Public Sub Remove(Index As Integer)
    On Error GoTo ErrHandler 'Avoid any error messages
    
    Dim I As Integer
    
    'Bump everything down
    If Index < UBound(S) Then
        For I = Index + 1 To UBound(S)
            S(I - 1) = S(I)
        Next I
    End If
    
    'Re-dimension the array
    ReDim Preserve S(0 To UBound(S) - 1)
    
    Exit Sub
    
ErrHandler:
    
End Sub

Public Function Count() As Integer
    On Error GoTo ErrHandler 'Avoid any error messages
    
    'Return the size of S
    Count = UBound(S) + 1
    Exit Function
    
ErrHandler:
    Count = 0
    
End Function

Public Function Item(Index As Integer) As String
    On Error GoTo ErrHandler 'Avoid any error messages
    
    'Return the string at Index in the S() array
    Item = S(Index)
    Exit Function
    
ErrHandler:
    Item = ""
    
End Function
