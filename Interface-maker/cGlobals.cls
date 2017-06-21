VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ----------------------------------------------------------------------------
'          VBScript.RegExp - PCRE2 portable proxy wrapper class
' ----------------------------------------------------------------------------
'
' (c) Made by Polshyn Stanislav (Dragokas)
'
' (c) Based on PCRE2 wrapper by Jason Peter Brown (jpbro) <jason@bitspaces.com> - https://github.com/jpbro/VbPcre2
'
' Credits to all, who participated in project: Tanner_H, oumba, DEXWERX, dilettante (vbforums.com)
'
' Supporting threads:
' https://github.com/dragokas/VbPcre2
' https://github.com/jpbro/VbPcre2
' http://www.vbforums.com/showthread.php?847611
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' Some documentation from http://www.pcre.org/current/doc/html/pcre2api.html
' Some additional documentation for PCRE2 Regex options by Seongmyun Cho taken from https://github.com/xnsystems/kpcre/wiki
' Some additional documentation from http://manpages.ubuntu.com/manpages/zesty/man3/pcre2api.3.html

Public Type REGEXP_STATE
    Global      As Boolean
    IgnoreCase  As Boolean
    Multiline   As Boolean
    Pattern     As String
    UsePcre     As Boolean
    Init        As Boolean
    hLib        As Long
End Type

Public Type REGEXP_MATCH
    FirstIndex  As Long
    Length      As Long
    SubMatches  As IRegExpSubMatches 'cRegExpSubMatches
    Value       As String
End Type

Public Enum e_SubstitutionAction
   subaction_None ' Perform no substitution
   subaction_Replace ' Replace the match with the supplied text
   subaction_ReplaceAndCache  ' Replace the match with the supplied text. If a subsequent exact match is found, replace with the same cached value and skip the Matched event.
End Enum

Public Enum PCRE_ReturnCode
   PCRE_RC_OK = 0

   'Error codes: no match and partial match are "expected" errors.
   PCRE_RC_ERROR_NOMATCH = -1
   PCRE_RC_ERROR_PARTIAL = -2

   'Error codes for UTF-8 validity checks
   PCRE_RC_ERROR_UTF8_ERR1 = -3
   PCRE_RC_ERROR_UTF8_ERR2 = -4
   PCRE_RC_ERROR_UTF8_ERR3 = -5
   PCRE_RC_ERROR_UTF8_ERR4 = -6
   PCRE_RC_ERROR_UTF8_ERR5 = -7
   PCRE_RC_ERROR_UTF8_ERR6 = -8
   PCRE_RC_ERROR_UTF8_ERR7 = -9
   PCRE_RC_ERROR_UTF8_ERR8 = -10
   PCRE_RC_ERROR_UTF8_ERR9 = -11
   PCRE_RC_ERROR_UTF8_ERR10 = -12
   PCRE_RC_ERROR_UTF8_ERR11 = -13
   PCRE_RC_ERROR_UTF8_ERR12 = -14
   PCRE_RC_ERROR_UTF8_ERR13 = -15
   PCRE_RC_ERROR_UTF8_ERR14 = -16
   PCRE_RC_ERROR_UTF8_ERR15 = -17
   PCRE_RC_ERROR_UTF8_ERR16 = -18
   PCRE_RC_ERROR_UTF8_ERR17 = -19
   PCRE_RC_ERROR_UTF8_ERR18 = -20
   PCRE_RC_ERROR_UTF8_ERR19 = -21
   PCRE_RC_ERROR_UTF8_ERR20 = -22
   PCRE_RC_ERROR_UTF8_ERR21 = -23

   'Error codes for UTF-16 validity checks
   PCRE_RC_ERROR_UTF16_ERR1 = -24
   PCRE_RC_ERROR_UTF16_ERR2 = -25
   PCRE_RC_ERROR_UTF16_ERR3 = -26

   'Error codes for UTF-32 validity checks
   PCRE_RC_ERROR_UTF32_ERR1 = -27
   PCRE_RC_ERROR_UTF32_ERR2 = -28

   'Error codes for pcre2[_dfa]_match= , substring extraction functions, context
   ' functions, and serializing functions. They are in numerical order. Originally
   ' they were in alphabetical order too, but now that PCRE2 is released, the
   ' numbers must not be changed.
   PCRE_RC_ERROR_BADDATA = -29
   PCRE_RC_ERROR_MIXEDTABLES = -30         ' Name was changed
   PCRE_RC_ERROR_BADMAGIC = -31
   PCRE_RC_ERROR_BADMODE = -32
   PCRE_RC_ERROR_BADOFFSET = -33
   PCRE_RC_ERROR_BADOPTION = -34
   PCRE_RC_ERROR_BADREPLACEMENT = -35
   PCRE_RC_ERROR_BADUTFOFFSET = -36
   PCRE_RC_ERROR_CALLOUT = -37             ' Never used by PCRE2 itself
   PCRE_RC_ERROR_DFA_BADRESTART = -38
   PCRE_RC_ERROR_DFA_RECURSE = -39
   PCRE_RC_ERROR_DFA_UCOND = -40
   PCRE_RC_ERROR_DFA_UFUNC = -41
   PCRE_RC_ERROR_DFA_UITEM = -42
   PCRE_RC_ERROR_DFA_WSSIZE = -43
   PCRE_RC_ERROR_INTERNAL = -44
   PCRE_RC_ERROR_JIT_BADOPTION = -45
   PCRE_RC_ERROR_JIT_STACKLIMIT = -46
   PCRE_RC_ERROR_MATCHLIMIT = -47
   PCRE_RC_ERROR_NOMEMORY = -48
   PCRE_RC_ERROR_NOSUBSTRING = -49
   PCRE_RC_ERROR_NOUNIQUESUBSTRING = -50
   PCRE_RC_ERROR_NULL = -51
   PCRE_RC_ERROR_RECURSELOOP = -52
   PCRE_RC_ERROR_RECURSIONLIMIT = -53
   PCRE_RC_ERROR_UNAVAILABLE = -54
   PCRE_RC_ERROR_UNSET = -55
   PCRE_RC_ERROR_BADOFFSETLIMIT = -56
   PCRE_RC_ERROR_BADREPESCAPE = -57
   PCRE_RC_ERROR_REPMISSINGBRACE = -58
   PCRE_RC_ERROR_BADSUBSTITUTION = -59
   PCRE_RC_ERROR_BADSUBSPATTERN = -60
   PCRE_RC_ERROR_TOOMANYREPLACE = -61
   PCRE_RC_ERROR_BADSERIALIZEDDATA = -62

   [_PCRE_RC_ERROR_FIRST] = -1
   [_PCRE_RC_ERROR_LAST] = -62   ' If you add more PCRE2 error codes, make sure to update this value!
End Enum
'
'The following option bits can be passed only to pcre2_compile(). However,
' they may affect compilation, JIT compilation, and/or interpretive execution.
' The following tags indicate which:
'
' C   alters what is compiled by pcre2_compile()
' J   alters what is compiled by pcre2_jit_compile()
' M   is inspected during pcre2_match() execution
' D   is inspected during pcre2_dfa_match() execution
Public Enum PCRE_CompileOptions
   PCRE_CO_ALLOW_EMPTY_CLASS = &H1&            ' C
   PCRE_CO_ALT_BSUX = &H2&                     ' C
   PCRE_CO_AUTO_CALLOUT = &H4&                 ' C
   PCRE_CO_CASELESS = &H8&                     ' C
   PCRE_CO_DOLLAR_ENDONLY = &H10&              '   J M D
   PCRE_CO_DOTALL = &H20&                      ' C
   PCRE_CO_DUPNAMES = &H40&                    ' C
   PCRE_CO_EXTENDED = &H80&                    ' C
   PCRE_CO_FIRSTLINE = &H100&                  '   J M D
   PCRE_CO_MATCH_UNSET_BACKREF = &H200&        ' C J M
   PCRE_CO_MULTILINE = &H400&                  ' C
   PCRE_CO_NEVER_UCP = &H800&                  ' C
   PCRE_CO_NEVER_UTF = &H1000&                 ' C
   PCRE_CO_NO_AUTO_CAPTURE = &H2000&           ' C
   PCRE_CO_NO_AUTO_POSSESS = &H4000&           ' C
   PCRE_CO_NO_DOTSTAR_ANCHOR = &H8000&         ' C
   PCRE_CO_NO_START_OPTIMIZE = &H10000         '   J M D
   PCRE_CO_UCP = &H20000                       ' C J M D
   PCRE_CO_UNGREEDY = &H40000                  ' C
   PCRE_CO_UTF = &H80000                       ' C J M D
   PCRE_CO_NEVER_BACKSLASH_C = &H100000        ' C
   PCRE_CO_ALT_CIRCUMFLEX = &H200000           '   J M D
   PCRE_CO_ALT_VERBNAMES = &H400000            ' C
   PCRE_CO_USE_OFFSET_LIMIT = &H800000         '   J M D
End Enum

Public Enum PCRE_GlobalConst
   PCRE2_ERROR_NOMATCH = -1

   PCRE2_SUBSTITUTE_GLOBAL = &H100

   PCRE2_ANCHORED = &H80000000
   PCRE2_NO_UTF_CHECK = &H40000000
   PCRE2_NOTBOL = &H1
   PCRE2_NOTEOL = &H2
   PCRE2_NOTEMPTY = &H4
   PCRE2_NOTEMPTY_ATSTART = &H8
   PCRE2_PARTIAL_SOFT = &H10
   PCRE2_PARTIAL_HARD = &H20

   PCRE2_ALLOW_EMPTY_CLASS = &H1
   PCRE2_ALT_BSUX = &H2
   PCRE2_CASELESS = &H8
   PCRE2_DOLLAR_ENDONLY = &H10
   PCRE2_DOTALL = &H20
   PCRE2_EXTENDED = &H80
   PCRE2_FIRSTLINE = &H100
   PCRE2_MULTILINE = &H400
   PCRE2_UNGREEDY = &H40000
   PCRE2_UTF = &H80000
   PCRE2_NEVER_BACKSLASH_C = &H100000
   PCRE2_ALT_CIRCUMFLEX = &H200000
   PCRE2_AUTO_CALLOUT = &H4           ' TODO: Callouts not implemented
   PCRE2_DUPNAMES = &H40

' UNIMPLEMENTED
   PCRE2_MATCH_UNSET_BACKREF = &H200
   PCRE2_NEVER_UCP = &H800
   PCRE2_NEVER_UTF = &H1000
   PCRE2_NO_AUTO_CAPTURE = &H2000
   PCRE2_NO_AUTO_POSSESS = &H4000
   PCRE2_NO_DOTSTAR_ANCHOR = &H8000
   PCRE2_NO_START_OPTIMIZE = &H10000
   PCRE2_UCP = &H20000

   PCRE2_SUBSTITUTE_EXTENDED = &H200
   PCRE2_SUBSTITUTE_UNSET_EMPTY = &H400
   PCRE2_SUBSTITUTE_UNKNOWN_UNSET = &H800
End Enum

Public Enum REGEXP_File_Const
   MAX_PATH = 260&
   FILE_SHARE_READ = &H1&
   FILE_SHARE_WRITE = &H2&
   FILE_SHARE_DELETE = 4&
   FILE_READ_ATTRIBUTES = &H80&
   OPEN_EXISTING = 3&
   CREATE_ALWAYS = 2&
   GENERIC_READ = &H80000000
   GENERIC_WRITE = &H40000000
   FILE_ATTRIBUTE_DIRECTORY = &H10&
   INVALID_HANDLE_VALUE = &HFFFFFFFF
   ERROR_SUCCESS = 0&
   INVALID_FILE_ATTRIBUTES = -1&
   NO_ERROR = 0&
   FILE_BEGIN = 0&
   FILE_CURRENT = 1&
   FILE_END = 2&
   INVALID_SET_FILE_POINTER = &HFFFFFFFF
End Enum

