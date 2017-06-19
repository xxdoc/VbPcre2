VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPcre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Matched(ByRef p_MatchedText As String, ByRef p_SubstitutionAction As e_SubstitutionAction, ByRef p_Cancel As Boolean)

Public Sub InitClass()
End Sub

Public Sub TerminateClass()
End Sub

Public Function Options() As IPcreOptions
End Function

Public Function Replace(ByVal p_TextToSearch As String, ByVal p_ReplaceWithText As String, Optional ByVal p_RegexToMatch As String) As String
End Function

Public Function Test(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As Boolean
End Function

Public Function Execute(ByVal p_TextToSearch As String, Optional ByVal p_RegexToMatch As String) As IPcreMatches
End Function

Public Property Let GlobalSearch(ByVal p_Enable As Boolean)
End Property

Public Property Get GlobalSearch() As Boolean
End Property

Public Property Let IgnoreCase(ByVal p_Enable As Boolean)
End Property

Public Property Get IgnoreCase() As Boolean
End Property

Public Property Get Pattern() As String
End Property

Public Property Let Pattern(ByVal p_RegexPattern As String)
End Property

