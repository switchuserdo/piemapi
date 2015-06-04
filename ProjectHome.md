Developed as part of a project for work, PieMAPI provides a greatly simplified way to access Extended MAPI, one of the ways to access Microsoft Outlook data.

While it is not a complete implementation, and doesn't aim to be, it should be a good starting point for most MAPI projects.

Features it provides include:
  * Enumerate inbox/sentbox items with three lines and a for-loop
  * Easily extend Outlook COM objects -- useful for getting at protected values without needing permission, such as SenderEmailAddress and Recipients
  * Save messages (including attachments) to the computer with a single call
  * Easy access to the Recipients of a message