# VBScript - PCRE2 portable VB6 proxy wrapper

## WARNING: This is still Alpha-version:
* It's not tested enough.
* It's not support .FirstIndex property yet (waiting for realization in PCRE2 wrapper).

(c) Copyright: 2017 Polshyn Stanislav <dragokas <at> safezone.cc>
(c) Based on PCRE2 wrapper by Jason Peter Brown (jpbro): https://github.com/jpbro/VbPcre2

This fork is basically intended for insurance that regular expressions engine, based on "VBScript.Regexp" in your program will not fail. If it is failed (e.g. damaged VBScript.dll, or registry data), PCRE2 library will be used instead automatically.

The main reason why you may need this proxy wrapper, instead of original PCRE2 wrapper by Jason:
you can easily integrate it in your ready big project with minimal steps, like:

replacing your code:

instead of
[code]
Dim oRegexp as Object
set oRegexp = CreateObject("VBScript.Regexp")
[/code]

with:

[code]
Dim oRegexp as IRegExp
set oRegexp = New cRegExp
[/code]

Add 1 tlb + 1 cls + pcre2-16.dll file to your project, and that's all.
Object model of wrapper fully imitate "VBScript.Regexp" object model for you.
_____________________________________________________________________________________

For details, look in Readme.md file of .\Using directory.

Have a nice day :)
