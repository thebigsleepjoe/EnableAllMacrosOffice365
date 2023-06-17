# 🖥️ Enable all macros within Office 365
This simple script works as of 6/16/2023; in the futur the paths to the apps (Word, Excel, Outlook) may change to versions later than 16.0.

The script can easily be changed to support version changes, however.

## What is this?
This script is used to easily configure the enabled state of macros within the Office 365 environment. By default, it reconfigures Outlook, Word, and Excel. It can be easily modified to support other Office 365 applications, or to disable macros instead of enabling them.

It will only configure the currently logged in user. It does *not* lock the variable in this state, meaning the user can just turn it back off if they want.

## 🚀 How do I use?
Run the script, enablemacros.ps1 as System to reconfigure the current user's registry keys to enable Excel/Word macros automatically.

It is best used with an RMM that supports Powershell.

## 🔓 What are the security implications?
This script is not for regular users; you should only ever implement this in a specific environment where you are using an outdated application that requires macros always on.

📖 Read more about [how macros are dangerous](https://support.microsoft.com/en-us/office/protect-yourself-from-macro-viruses-a3f3576a-bfef-4d25-84dc-70d18bde5903)

Of course, you can configure this script to hard-disable all macros on the current user, as well.
