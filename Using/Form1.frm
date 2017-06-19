VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VbPcre2 Test"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    TestRegexProxy bUsePcre:=True       'test with PCRE2
    TestRegexProxy bUsePcre:=False      'test with VBScript.Regexp
    Unload Me
End Sub

Sub TestRegexProxy(bUsePcre As Boolean)
   
   Debug.Print "Proxy Test - " & IIf(bUsePcre, "PCRE2", "VBScript")
   
   Dim lo_RegEx         As IRegExp
   Dim lo_Matches       As Object
   Dim lo_Match         As Object
   Dim lo_Submatches    As Object
   Dim lo_SubMatch      As Variant

'   'alternate declaration
'   Dim lo_RegEx         As IRegExp                  'VBScript_RegExp_55.RegExp
'   Dim lo_Matches       As IRegExpMatchCollection   'VBScript_RegExp_55.MatchCollection
'   Dim lo_Match         As IRegExpMatch             'VBScript_RegExp_55.Match
'   Dim lo_Submatches    As IRegExpSubMatches        'VBScript_RegExp_55.SubMatches
'   Dim lo_SubMatch      As Variant
   
   Dim l_SubjectText As String
   Dim l_Regex As String
   
   Dim ii As Long
   Dim jj As Long
   
   l_SubjectText = "File1.zip.exe" & vbCrLf & "File2.com" & vbCrLf & "File 3"
   l_Regex = "[\w ]+(\.\S+?)*$"
   
   ' creating an instance
   Set lo_RegEx = New cRegExp
   
   ' settings
   With lo_RegEx
      .IgnoreCase = True
      .Global = True
      .MultiLine = True
      .Pattern = l_Regex
      .UsePcre = bUsePcre ' set whether we want to use PCRE2 or VBScript.Regexp version
   End With
   
   Set lo_Matches = lo_RegEx.Execute(l_SubjectText)
   
   Debug.Print "Match Count: " & lo_Matches.Count
    
   For Each lo_Match In lo_Matches
   
      Set lo_Submatches = lo_Match.SubMatches
    
      ii = ii + 1
      Debug.Print "Match #" & ii & ": " & lo_Match.Value
      Debug.Print "Sub Match Count: " & lo_Submatches.Count
      
      'iterating submatches
      jj = 0
      For Each lo_SubMatch In lo_Submatches
        jj = jj + 1
        Debug.Print "SubMatch # " & jj & ": " & lo_SubMatch
      Next
      
'      'alternate
'      For jj = 0 To lo_Submatches.Count - 1
'        Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Submatches.Item(jj)  'alternate 1
'        Debug.Print "SubMatch # " & jj + 1 & ": " & lo_Submatches(jj)       'alternate 2
'      Next
   Next
   
   ' destroy an instance of class
   Set lo_RegEx = Nothing
   
   Debug.Print
End Sub
