VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRegExpMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Init(FirstIndex As Long, Value As String, SubMatches As IRegExpSubMatches)
Attribute Init.VB_MemberFlags = "40"
End Sub

Public Property Get FirstIndex() As Long
End Property

Public Property Get Length() As Long
End Property

Public Property Get SubMatches() As IRegExpSubMatches
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
End Property

