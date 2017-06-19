# TLB file maker

Intended for using as a part of "VBScript.Regexp - PCRE2 VB6 proxy wrapper"

## How to build
* compile project as IRegexp.dll
* run DLLUnreg.vbs
* use OleView to extract .odl (File -> View TypeLib... -> IRegexp.dll), save text as IRegexp.odl
* edit IRegexp.odl file:
1. Rename library Regexp_DLL by Regexp_Interface
1. Remove lines like [restricted] void Missing7();
2. Move e_SubstitutionAction typedef scope to the beginning of library Regexp { scope.
3. Replace "GlobalSearch" by "Global" in these lines:
 - HRESULT GlobalSearch([out, retval] VARIANT_BOOL* );
 - HRESULT GlobalSearch([in, out] VARIANT_BOOL* );
* install Microsoft Visual Studio
* change code of "RunMidl - VS 2017.cmd" file to point to actual path of "Developer Command Prompt for VS" .cmd file (see in Start menu -> Visual Studio -> Visual Studio Tools).
* Execute "RunMidl - VS 2017.cmd"
* You will get IRegexp.tlb file

Warnings in cmd console window is normal.
If your can't compile tlb, look on "errors" in console.

