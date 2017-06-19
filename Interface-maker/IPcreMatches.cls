VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPcreMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InitClass()
End Sub

Public Property Get Count() As Long
End Property

Public Property Get Item(ByVal p_ZeroBasedIndex As Long) As IPcreMatch
Attribute Item.VB_UserMemId = 0
End Property

Public Function Enumerator() As IEnumVARIANT
Attribute Enumerator.VB_UserMemId = -4
Attribute Enumerator.VB_MemberFlags = "40"
End Function

Public Property Get Text() As String
End Property

Public Sub AppendText(ByVal p_Text As String)
End Sub

Public Sub Add(ByVal p_TextToSearch As String, pa_OvectorArray() As Long, ByVal p_MatchCount As Long)
End Sub
