VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRegExpMatchCollection"
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

Public Property Get Item(Index As Long) As IRegExpMatch
End Property

Public Function Enumerator() As IEnumVARIANT
Attribute Enumerator.VB_UserMemId = -4
Attribute Enumerator.VB_MemberFlags = "40"
End Function

Public Sub Add(FirstIndex As Long, Value As String, SubMatches As IRegExpSubMatches)
Attribute Add.VB_MemberFlags = "40"
End Sub
