VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPcreOptionsReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InitClass()
Attribute InitClass.VB_MemberFlags = "40"
End Sub

Public Function OptionFlags() As Long
Attribute OptionFlags.VB_MemberFlags = "40"
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

Public Property Get ErrorOnUnknownCapturingGroups() As Boolean
End Property

Public Property Let ErrorOnUnknownCapturingGroups(ByVal p_Enabled As Boolean)
End Property

Public Property Get TreatUnknownCapturingGroupsAsEmptyStrings() As Boolean
End Property

Public Property Let TreatUnknownCapturingGroupsAsEmptyStrings(ByVal p_Enabled As Boolean)
End Property

Public Property Get ExtendedReplacement() As Boolean
End Property

Public Property Let ExtendedReplacement(ByVal p_Enabled As Boolean)
End Property

