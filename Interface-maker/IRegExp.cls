VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRegExp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get UsePcre() As Boolean
End Property

Public Property Let UsePcre(bValue As Boolean)
End Property

Public Property Get GlobalSearch() As Boolean
End Property

Public Property Let GlobalSearch(bValue As Boolean)
End Property

Public Property Get IgnoreCase() As Boolean
End Property

Public Property Let IgnoreCase(bValue As Boolean)
End Property

Public Property Get Multiline() As Boolean
End Property

Public Property Let Multiline(bValue As Boolean)
End Property

Public Property Get Pattern() As String
End Property

Public Property Let Pattern(sValue As String)
End Property

Public Function Replace(sourceString As String, replaceVar As Variant) As String
End Function

Public Function Test(sourceString As String) As Boolean
End Function

Public Function Execute(sourceString As String) As IRegExpMatchCollection
End Function

Public Property Get PCRE2() As IPcre
End Property

