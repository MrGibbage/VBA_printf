# VBA printf
A VBA function that mimics java's printf function.

I have tried to make a function that should operate just like java's printf function, as documented here:
https://docs.oracle.com/javase/8/docs/api/java/util/Formatter.html

For example:
Debug.Print printf("The quick brown %-10S jumps over the lazy %-8S!", "foxxy fox", "dog")
Debug.Print printf("%8S The quick brown %S jumps over the lazy!", "foxxy fox", "dog")
Debug.Print printf("floats: %+4.2f %+.0e %+E %+0000.0f \n", 3.1416, 3.1416, 3.1416, 3.1416);
Debug.Print printf("The quick 10%% brown %S jumps over the\nlazy %s", "fox", "dog")
Debug.Print printf("Boolean tests (vbFalse) %b\n(nothing) %b\n(Null) %b\n(vbNull) %b\n(1=1) %b\n(1=0) %b\n(1) %b\n(0) %b XXX", vbFalse, , Null, vbNull, 1 = 1, 1 = 0, 1, 0)
Debug.Print printf("Hex test: %h XXX", 255)
Debug.Print printf("Param test, 3: %3$s, 2: %2$S, 1: %1$s, 4: %s XXX", "one", "two", "three", "four")
Debug.Print printf("Char test:\n&H63: %c\nf: %c XXX", &H63, "f")
Debug.Print printf("D %2$+5d XXX", 32, 16)
Debug.Print printf("D %2$00000.d XXX", 32, 16.5)
Debug.Print printf("Oct test: %o XXX", 64)
Debug.Print printf("Hex test: %h XXX", 255)
Debug.Print printf("E: %0.000E XXX", 256789125)
Debug.Print printf("H: %tF XXX", Now)

There are a couple of functions that will not be supported:
* %Tc, %Tz, %TZ Because VBA does not support Time Zones. I could try and write something to fake it, I suppose...
* %TL, %TN because VBA does not support milliseconds and nanoseconds
* %G (alternate scientific notation) because I don't understand it and I can't find any decent examples
* %A (alternate hexadecimal notation) because I don't understand it and I can't find any decent examples

I have offered some example usage code at the bottom, which should look just like any java use of the printf function.

That's my goal, anyway.

I'm not very good programmer, so feel free to offer suggestions. I'm trying to learn!
