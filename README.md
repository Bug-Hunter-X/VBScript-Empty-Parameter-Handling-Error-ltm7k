# VBScript Empty Parameter Handling Error

This repository demonstrates a common error in VBScript: improper handling of empty parameters.  The `bug.vbs` file shows the problematic code that fails if a function receives an empty parameter. The `bugSolution.vbs` file provides a corrected version with robust error handling.

**Problem:** VBScript doesn't automatically prevent the execution of functions with empty parameters. This can lead to unexpected runtime errors if the function isn't designed to handle these cases gracefully.

**Solution:** Always explicitly check for empty or null parameters using `IsEmpty()` before accessing their values and handle such cases appropriately.