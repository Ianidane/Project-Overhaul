import sys
import os
global input1           #This variable is defined in C# Form1.cs
global input2
input2 = "friendship"   #this variable can be seen in C# - there we used .GetVariable to see it


def makefile(txt):
    os.chdir("C:/Users/Alex/Desktop/Work/Project-Overhaul/C#/wfavbVBA2PHPone.ConvertedToC#") 
    with open("success.txt", "a") as myfile:
        myfile.write(txt)

makefile(input1)        #Using input1 variable that we defined in C# - the argument is "yaaay"
