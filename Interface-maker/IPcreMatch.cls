VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPcreMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get SearchedText() As String
End Property

Public Property Get MatchedText() As String
End Property

Public Property Get SubMatchCount() As Long
End Property

Public Property Get SubMatchValue(ByVal p_ZeroBasedSubMatchIndex As Long) As String
End Property

Public Sub Init(ByVal p_SearchedText As String, pa_OvectorArray() As Long, ByVal p_MatchCount As Long)
End Sub

