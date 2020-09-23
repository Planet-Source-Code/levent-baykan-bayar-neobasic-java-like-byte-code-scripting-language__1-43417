===========================================================================
NeoBasic High Power Semi-Object Oriented Scripting Language
Author: Levent Baykan Bayar (leventbbayar@operamail.com)
Release Date:02.21.2003
===========================================================================
HOW TO USE IT IN YOUR PROJECTS:
===========================================================================
1)Add compile class in your project
2)See code of Command6_Click & Command7_Click in CForm
to learn how to call class
===========================================================================
CREDITS:
=> I have learned a lot from Justin Tunney's JELL,
   I have used while and if sub routines directly from his engine.
   Thanks him a lot...
   You can find his works from PSC by searching "JELL"


============================================================================
You can use this code in your applications as long as you
give me a proper credit and inform me for the changes you've made on it
This code is freeware as long as you use it in freeware, If you want to use it
in commercial applications you must pay me something for it.

I am releasing this piece of code for the public for
people who wish to learn from it.
============================================================================
Features
============================================================================
1)Faster than text interpreting systems. interpretes using byte-code system.
2)Local,global variable and array support.
3)Optimization feature for math expressions.
4)Semi-Object orientence,see GetIdent,Ch_Ident subs
5)Multiple Commands in one line
  To send a command like this msgbox("Your name is " + inputbox("What is your name?"))
6)System Ident variables
    To show time for example : msgbox(time)
7)Include directive for including ActiveX DLLs to script
to use functions in them.see file_io.txt example,and see io_dll folder to see a sample of 
dll.

============================================================================
Bugs/To Do
============================================================================
1)Unfortunately many bugs possible,because of code's complexity.Try yourself.
2)Byte-Code system may be improved.(I advice P-code compiling.)
3)Strange speed problem when compiled to native code.
4)Some bugs in optimization system.
============================================================================
CONTACT:
mail:leventbbayar@operamail.com
web:http://www.geocities.com/leventbbayar