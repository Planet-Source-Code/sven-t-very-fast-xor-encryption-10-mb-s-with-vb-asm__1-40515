
Descrition:

This project provides a class for xor-encrypting strings and files.
The class is over 100 times faster than normal vb-code, it encrypts with up to 13 MB/s!
VB does only about 0.1 MB/s, that means it encrypts a 10MB string in less than 1 second, where other routines posted on psc would take about 90 seconds!
You can also pass files to the class which will be encrypted with up to 6 MB/s!
(tested on a P-II 350)
How it works:
The class doesn't use any dlls, it all works with vb-code and a little bit assembler!
another advantage of this technique is that it has the same speed when it is used in the ide.
The sourcecode is commented, so i think you can also learn something about using asm in vb.
I hope you find it useful, thanks for your vote.

Requirements:
you can use the class as it is.
if you want to compile the asm sourcecode, you need the nasm-compiler.
get it for free at http://sourceforge.net/projects/nasm
copy the nasmw.exe to this directory and call the file GenBin.bat to compile it.
the compiler will generate binary files like the file fastxor.bin in this directory.
