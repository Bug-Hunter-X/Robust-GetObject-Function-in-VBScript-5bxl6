# Robust GetObject Function in VBScript

This repository demonstrates a common issue with VBScript's `GetObject` function and provides a more robust solution.  The original `GetObject` function can fail if the specified COM object is not already running, leading to script errors.  The improved version handles this gracefully.

## Bug Description

The original `GetObject` function attempts to get a reference to an existing COM object. If the object isn't running, it fails, causing a script error.  This is problematic when you need to ensure the object exists before working with it.

## Solution

The improved function first attempts to get the object using `GetObject`. If that fails, it uses `CreateObject` to create a new instance of the object. This approach ensures the script continues even if the object needs to be created.

## Usage

Simply replace your existing calls to `GetObject` with calls to the improved `GetObject` function shown below.  Remember to handle potential errors appropriately in your main script, as `CreateObject` might still fail due to other reasons (like missing dependencies).