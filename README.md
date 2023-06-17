# 🖥️ Enable all macros within Office 365

## What is this?
This script is used to easily configure the enabled state of macros within the Office 365 environment. By default, it reconfigures Outlook, Word, and Excel. It can be easily modified to support other Office 365 applications, or to disable macros instead of enabling them.

It will only configure the currently logged in user. It does *not* lock the variable in this state, meaning the user can just turn it back off if they want.

## 🚀 How do I use?
Run the script, enablemacros.ps1 as System to reconfigure the current user's registry keys to enable Excel/Word macros automatically.

If you want to disable macros, run disablemacros.ps1 as System. Works same as above, but just disables macros instead, w/o notifying.

It is best used with an RMM (Remote monitoring and management) tool that supports Powershell scripting.

## 🔓 What are the security implications?
This script is not for regular users; you should only ever implement this in a specific environment where you are using an outdated application that requires macros always on.

📖 Read more about [how macros are dangerous](https://support.microsoft.com/en-us/office/protect-yourself-from-macro-viruses-a3f3576a-bfef-4d25-84dc-70d18bde5903).

Of course, you can use the alternate script to disable all macros regardless.

## Will this work for ___?
1. Navigate to "Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\" within regedit.
2. Check Word, Excel, and Outlook for what versions are present on your machine.
3. If you see "Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Security" then yes.

You can modify the script to support other versions of Office by changing the "16.0" to whatever version you need. It's super easy.

As of 6/16/2023, the scripts do work for the Office 365 suite.