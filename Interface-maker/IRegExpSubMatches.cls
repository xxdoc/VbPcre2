VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRegExpSubMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InitClass()
Attribute InitClass.VB_MemberFlags = "40"
End Sub

Public Property Get Count() As Long
End Property

Public Property Get Item(Index As Long) As String
Attribute Item.VB_UserMemId = 0
End Property

Public Function Enumerator() As IEnumVARIANT
Attribute Enumerator.VB_UserMemId = -4
Attribute Enumerator.VB_MemberFlags = "40"
End Function

Public Sub Add(sValue As String)
Attribute Add.VB_MemberFlags = "40"
End Sub
