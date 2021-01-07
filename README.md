# Hacky-Holidays-2020-Writeups

[TOC]

## Happy New Maldoc

>   #### CHALLENGE INFORMATION
>
>   Santa received a document containing some wishes for him… We're not entirely sure though whether it's virus-free.
>
>   *Author information: This challenge is developed by Deloitte.*

We are provided a file named HappyNewYear.pptm. This is indicates that it is a power point presentation with macros enable (as indicated by the .pptm extension). 

After opening the file we see:

![ppt.png](./images/ppt.png)

Before doing anything I like seeing what will happen so I run the presentation (macros are enabled) and I get a pop saying:

>   These wishes are not intended for you!

Dang it! One can always hope that something happens without doing anything XD. Anyway, considering that macros are enabled it would be a good idea to check out the macro scripts. You can do this by navigating to `View > Macros`. Select any of the scripts that pop up and click on `edit` and a screen with all the functions in the script will show up. It looks like this:

<script src="https://gist.github.com/goelp14/01f1fad0c4abc5f57f1ed3aab5e751a8.js"></script>