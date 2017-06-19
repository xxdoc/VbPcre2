VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPcreOptionsMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InitClass()
End Sub

Public Function OptionFlags() As Long
End Function

Public Property Get Anchored() As Boolean
End Property

Public Property Let Anchored(ByVal p_Enabled As Boolean)
End Property

Public Property Get SubjectIsBeginningOfLine() As Boolean
End Property

Public Property Let SubjectIsBeginningOfLine(ByVal p_Enabled As Boolean)
End Property

Public Property Get SubjectIsEndOfLine() As Boolean
End Property

Public Property Let SubjectIsEndOfLine(ByVal p_Enabled As Boolean)
End Property

Public Property Get MatchIfEmpty() As Boolean
End Property

Public Property Let MatchIfEmpty(ByVal p_Enabled As Boolean)
End Property

Public Property Get MatchIfEmptyAtStart() As Boolean
End Property

Public Property Let MatchIfEmptyAtStart(ByVal p_Enabled As Boolean)
End Property

Public Property Get PerformUtfCheck() As Boolean
End Property

Public Property Let PerformUtfCheck(ByVal p_Enabled As Boolean)
End Property

Public Property Get MatchedEventEnabled() As Boolean
End Property
   
Public Property Let MatchedEventEnabled(ByVal p_Enabled As Boolean)
End Property
   
Public Property Get MatchedEventSubstitutionCacheSize() As Long
End Property

Public Property Let MatchedEventSubstitutionCacheSize(ByVal p_MaxCount As Long)
End Property

