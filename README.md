# Unexpected Error Handling in VBScript: Err.Raise

This repository demonstrates a common yet subtle error in VBScript's error handling using `Err.Raise`.  The issue arises from the way VBScript handles errors and the context in which `Err.Raise` is used.

The provided `bug.vbs` script shows the problematic code, while `bugSolution.vbs` offers a corrected and more robust approach.

## Problem
The original code attempts to check for an empty parameter and raise an error if found.  However, the error handling might not always function as intended, depending on how the function is called and handled by calling code. 

## Solution
The solution focuses on proper error handling and context management. Using `On Error Resume Next` before the potential error and `On Error GoTo 0` afterward ensures that if an error does occur, it can be caught and handled gracefully. 

## How to run
Save the `.vbs` files and run them through a VBScript interpreter (usually by double-clicking them on a Windows system).