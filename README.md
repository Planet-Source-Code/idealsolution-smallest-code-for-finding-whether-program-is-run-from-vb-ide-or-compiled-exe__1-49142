<div align="center">

## SMALLEST CODE FOR FINDING WHETHER PROGRAM IS RUN FROM VB IDE OR COMPILED EXE


</div>

### Description

This function will return whether your program is running in Visual Basic Or it is running from the compiled EXE.

This Function tries to print in the immediate window using the Debug.print method, which is available only in VB IDE and will be removed while compiling the code to EXE (or dll or ocx). The value being print using Debug.print method the raises a division by zero error and the error handler set the InIDE function to TRUE.

I saw another code in Planet source code doing the same thing using a static variable and also calling the same function recursively. but this code is smaller than that.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-10-14 12:20:02
**By**             |[IdealSolution](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/idealsolution.md)
**Level**          |Beginner
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[SMALLEST\_C16584710142003\.zip](https://github.com/Planet-Source-Code/idealsolution-smallest-code-for-finding-whether-program-is-run-from-vb-ide-or-compiled-exe__1-49142/archive/master.zip)





### Source Code

<P>'// This function will return whether your program is running in Visual Basic Or it is running from the compiled EXE. This Function tries to print in the immediate window using the Debug.print method, which is available only in VB IDE and will be removed while compiling the code to EXE (or dll or ocx). The value being print using Debug.print method the raises a division by zero error and the error handler set the InIDE function to TRUE. </P>
<P>'// I saw another code in Planet source code doing the same thing using a static variable and also calling the same function recursively. but this code is smaller than that. </P><BR>
Private Function InIDE() As Boolean <BR>
  On Error GoTo Xit<BR>
  '//this line will execute only in VB IDE <BR>
  '//generate a division by zero error<BR>
  Debug.Print 1 / 0 <BR>
  Exit Function<BR>
Xit:<BR>
  InIDE = True<BR>
End Function<BR>
<BR>
<u>Follow-up for the comments posted regarding this posting </u>
<p>There are 2 suggestions that are posted as smaller codes than this one. <br>
True, they are lesser lines than this code, but they are less efficient (takes more time) when compiled to exe. This is because code in this article becomes virtually zero since debug.print line is eliminated and the only remaining line is 'Exit Function'. Both the other codes has some line to process. For comparing these 3 methods, I have created a test project. Download it from the files of this article.</p>
<p><u>Conclusions <u/><BR>
<B>1. App.Log method is fastest in the VB IDE <BR>
2. Code used in this article is fastest in compiled EXE </B><BR>
It is better to have the Exe run faster.
</p>

