Prerequisites
-------------
Addin code is hosted on: https://danilkorotenko.github.io/OutlookAddin_SimpleBlock/index.html

[Manifest](/manifest.xml)

Issue Description
-----------------
Outlook on-send addin is invoked for appointments, despite only "ItemIs"..."Message" declared in manifest.

Steps To Reproduce
------------------
1. Install manifest: [Manifest](/manifest.xml)
2. Create a new appointment. Add at least one person to it.
3. Press the "Send" button.

Actual Result
-------------
The appointment is blocked to send.

Expected result
---------------
The message is sent.

Hardware
--------
MacBook Air M1, 2020

Software
--------
macOS 15.4.1

Outlook for Mac. Version 16.96.1 (25042021)

Screenshots
-----------
![](/Screenshot1.png)
