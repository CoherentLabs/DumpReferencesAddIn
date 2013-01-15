DumpReferences Add In
=====================

DumpReferences is an add-in for Visual Studio that prints the references of all projects loaded in the solution, checks
their GUIDs and prints eventual issues. It is intended for C/C++ solutions.

Purpose
=======

The add-in was created in order to alleviate problems related to two quirks in Visual Studio and msbuild:

1. If you build a Solution through msbuild and project A in it depends on another project B that is NOT in that solution, project B will be 
compiled and linked in it's default configuration. Usually this is the Debug configuration. This might lead to 
some Debug libraries being linked in a Release project or some other config inconsistency.

2. If project A in a solution references project B, also in the solution, then Visual Studio fails any compilation 
attempt with the cryptic "The project file ' ' has been renamed or is no longer in the solution." giving no actual info 
on what's wrong. There is a bug on msdn about this: http://connect.microsoft.com/VisualStudio/feedback/details/635201/project-file-has-been-renamed-or-is-no-longer-in-the-solution.
It's unclear if and when it'll be fixed and most solutions suggested by users involve removing all projects, adding them one-by-one 
and testing which one fails. This is unfeasible in large solutions.

How it works
=============

The add-in is fairly simple - it walks all projects in the solution and their references and prints:
 - all projects currently in the solution
 - all references that each project has
 - all unique references in the solution
 - possible issues: referenced projects missing from the solution; projects available in the solution but referenced by other projects with another GUID

To dump the diagnostic just right-click on your solution in the Solution Explorer and click "Dump References".
The results will be available in the Output window in the Dump References pane.
 
Build
======

Just build the solution and place the output DLL and DumpReferencesAddIn.AddIn in your VS add-in folder.

Testing
=======

If you want to test the add-in you could follow the instructions on how to do it available on MSDN, or just:
 - Rename DumpReferencesAddIn.AddIn to DumpReferencesAddInTesting.AddIn or something and make it point to the output Debug DLL of the add-in
 - Run the debugger on VS - that is in the project config under "Debug"->"Start Action" select "Start external program". 
 Point to "(VSInstallFolder)devenv.exe /resetaddin DumpReferencesAddIn.Connect" where "(VSInstallFolder)" is obviously the folder 
 where your VS IDE is installed.

Testing Done
============

The add-in has been successfully tested on MSVS 2012. It should also work on MSVS 2010 and 2008.

License
========
Copyright (C) 2013 Coherent Labs AD

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
documentation files (the "Software"), to deal in the Software without restriction, including without limitation the 
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit 
persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions 
of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, 
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

