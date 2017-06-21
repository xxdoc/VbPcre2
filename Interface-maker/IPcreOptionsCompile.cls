VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPcreOptionsCompile"
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

Public Property Get AllowDuplicateSubpatternNames() As Boolean
End Property

Public Property Let AllowDuplicateSubpatternNames(ByVal p_Enabled As Boolean)
End Property

Public Property Get CaseSensitive() As Boolean
End Property

Public Property Let CaseSensitive(ByVal p_Enabled As Boolean)
End Property

Public Property Get Anchored() As Boolean
End Property

Public Property Let Anchored(ByVal p_Enabled As Boolean)
End Property

Public Property Get AutomaticallyInsertCallouts() As Boolean
End Property

Public Property Let AutomaticallyInsertCallouts(ByVal p_Enabled As Boolean)
End Property

Public Property Get IgnorePatternWhitspaceAndComments() As Boolean
End Property

Public Property Let IgnorePatternWhitspaceAndComments(ByVal p_Enabled As Boolean)
End Property

Public Property Get DotMatchesAllCharacters() As Boolean
End Property

Public Property Let DotMatchesAllCharacters(ByVal p_Enabled As Boolean)
End Property

Public Property Get Greedy() As Boolean
End Property

Public Property Let Greedy(ByVal p_Enabled As Boolean)
End Property

Public Property Get CheckUtfValidity() As Boolean
End Property

Public Property Let CheckUtfValidity(ByVal p_Enabled As Boolean)
End Property

Public Property Get Multiline() As Boolean
End Property

Public Property Let Multiline(ByVal p_Enabled As Boolean)
End Property

Public Property Get AllowEmptyClass() As Boolean
End Property

Public Property Let AllowEmptyClass(ByVal p_Enabled As Boolean)
End Property

Public Property Get AlternateBsuxHandling() As Boolean
End Property

Public Property Let AlternateBsuxHandling(ByVal p_Enabled As Boolean)
End Property

Public Property Get DollarMatchesEndOfStringOnly() As Boolean
End Property

Public Property Let DollarMatchesEndOfStringOnly(ByVal p_Enabled As Boolean)
End Property

Public Property Get Utf() As Boolean
End Property

Public Property Let Utf(ByVal p_Enabled As Boolean)
End Property

Public Property Get AlternateCircumflexHandling() As Boolean
End Property

Public Property Let AlternateCircumflexHandling(ByVal p_Enabled As Boolean)
End Property

