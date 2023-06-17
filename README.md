# 🖥️ Enable all macros within Office 365
This simple script works as of 6/16/2023. In the future it may not work or the paths to Excel/Word may change to versions later than 16.0

## What is this?
This script is used to easily configure the enabled state of macros within the Office 365 environment.

It will only configure the currently logged in user. It does *not* lock the variable in this state, meaning the user can just turn it back off if they want.

## 🚀 How do I use?
Run the script, enablemacros.ps1 as System to reconfigure the current user's registry keys to enable Excel/Word macros automatically.

It is best used with an RMM that supports Powershell.

## 🔓 What are the security implications?
This script is not for regular users; you should only ever implement this in a specific environment where you are using an outdated application that requires macros off.

📖 Read more about [how macros are dangerous](https://support.microsoft.com/en-us/office/protect-yourself-from-macro-viruses-a3f3576a-bfef-4d25-84dc-70d18bde5903)

Of course, you can configure this script to hard-disable all macros on the current user, as well.
