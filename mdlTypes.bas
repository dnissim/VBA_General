Attribute VB_Name = "mdlTypes"
'Author: David Nissim
Option Explicit

Type Transaction
    Source As String
    Location As String
    Category As String
End Type

Type Guess
    strGuess As String
    Matches As Long
    Proportion As Single
End Type

Type Coordinates2D
    Row As Long
    Column As Long
    Found As Boolean
End Type

Type CollectionItem
    Index As Long
    Value As Variant
    Found As Boolean
End Type
