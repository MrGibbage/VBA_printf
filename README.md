# VBA printf
A VBA function that mimics java's printf function.

I have tried to make a function that should operate just like java's printf function, as documented here:
https://docs.oracle.com/javase/8/docs/api/java/util/Formatter.html

There are a couple of functions that will not be supported:
* %Tc, %Tz, %TZ Because VBA does not support Time Zones. I could try and write something to fake it, I suppose...
* %TL, %TN because VBA does not support milliseconds and nanoseconds
* %G (alternate scientific notation) because I don't understand it and I can't find any decent examples
* %A (alternate hexadecimal notation) because I don't understand it and I can't find any decent examples

I have offered some example usage code at the bottom, which should look just like any java use of the printf function.

That's my goal, anyway.

I'm not very good programmer, so feel free to offer suggestions. I'm trying to learn!
