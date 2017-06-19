# VBScript.Regexp - PCRE2 portable VB6 proxy wrapper

## Portable version (single class + TLB)

Project intended to use inside another single-EXE projects. That is mean completely portable.

This project consist of:
 * cRegExp.cls (Dragokas' proxy wrapper + Fork of Jason's wrapper: https://github.com/jpbro/VbPcre2)
 * IRegexp.tlb (VBScript.Regexp + PCRE2 object models and enums)

## Using
 * Add cRegExp.cls to your project
 * Place pcre2-16.dll near (alternatively: add pcre2-16.dll in resources with ID 501).
 * Add reference to IRegexp.tlb
 * Use as usual "VBScript.Regexp" object model, just insert another declaration:
instead of
[code]
Dim oRegexp as Object
set oRegexp = CreateObject("VBScript.Regexp")
[/code]
use this:
[code]
Dim oRegexp as IRegExp
set oRegexp = New cRegExp
[/code]

To switch beetween VBScript.Regexp and PCRE2 versions, use property .UsePcre

## Where can I get pcre2-16.dll ?
 * See pre-compiled one in Tanner's repository: https://github.com/tannerhelland/PCRE2-VB6-DLL/releases
 * Or compile it yourself from Tanner's ".\vstudio" directory.
