# VBScript.Regexp - PCRE2 portable VB6 proxy wrapper

## Portable version (single EXE: class + TLB)

Project intended to use inside another single-EXE projects. That is mean completely portable.

This project consist of:
 * cRegExp.cls (Dragokas' proxy wrapper + Fork of Jason's wrapper: https://github.com/jpbro/VbPcre2)
 * IRegexp.tlb (VBScript.Regexp + PCRE2 object models and enums)
 * pcre2-16.dll (PCRE2 engine, precompiled by Tanner from official sources)

## Using
 * Add cRegExp.cls to your project
 * Place pcre2-16.dll near (alternatively: add pcre2-16.dll in resources with ID 501 - already included in demo above).
 * Add reference to IRegexp.tlb (you need to run VBP project as admin. first time)
 * Use as usual "VBScript.Regexp" object model, just insert another declaration:
instead of:
```
Dim oRegexp as Object
set oRegexp = CreateObject("VBScript.Regexp")
```
use this:
```
Dim oRegexp as IRegExp
set oRegexp = New cRegExp
```
or this:
```
Dim oRegexp as Object
Dim oRegexpProxy as IRegExp
Dim oRegexpProxy = New cRegExp
set oRegexp = oRegexpProxy
```

To switch beetween VBScript.Regexp and PCRE2 engines manually, use property .UsePcre
To access PCRE2 extended object model directly, use property .PCRE2

## Where can I get pcre2-16.dll ?
 * See pre-compiled one in Tanner's repository: https://github.com/tannerhelland/PCRE2-VB6-DLL/releases
 * Or compile it yourself from Tanner's ".\vstudio" directory.
