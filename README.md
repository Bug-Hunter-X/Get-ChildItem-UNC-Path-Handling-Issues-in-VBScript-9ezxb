This repository demonstrates a common yet subtle bug in VBScript's `Get-ChildItem` function when dealing with UNC paths.  The issue involves inconsistent or erroneous results when retrieving file and directory listings from network shares, especially with long or complex paths.  The `bug.vbs` script showcases the problem; the `bugSolution.vbs` offers a more robust workaround using alternative methods or libraries.