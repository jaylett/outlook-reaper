# Outlook Repeat

Outlook Reaper is a smallish Python script that will pull out a folder
of email messages from Outlook as a Unix-style mbox. It has a number
of problems, listed below, but if they don't bother you too much it
might be helpful.

The way it works is based on [a previous script I
found](http://www.boddie.org.uk/python/COM.html), which has useful
setup information in terms of getting Python and COM talking
together.

Using it is pretty simple: run the script, navigate through your
Outlook folders, and then extract into the filename you gave on the
command line. So something like the following:

```
C:\> outlook_reaper.py out.mbox
------------------------------------------------------------
[Root]
0 Personal Folders
Select number of [E]xtract [Q]uit, followed by <Return>
0
------------------------------------------------------------
Personal Folders
0 Deleted Items
1 Inbox
2 Outbox
3 Sent Items
4 Calendar
5 Contacts
6 Journal
7 Notes
8 Tasks
9 Drafts
10 RSS Feeds
11 Junk E-mail
Select number of [E]xtract [Q]uit, followed by <Return>
1
------------------------------------------------------------
Inbox
Select number of [E]xtract [Q]uit, followed by <Return>
e
Extracting to out.mbox
```

Patches welcome; while I'm unlikely to need this in future, building
on this and keeping it in the same place will probably seed better in
search engines and avoid the meandering paths I took while searching
for a similar utility.

 * Nested messages generally don't work; I tried to use
   [Redemption](http://www.dimastr.com/redemption/) to access the
   actual Outlook message object for a message attachment, but it's
   not playing ball here. You shouldn't need Redemption to run the
   script; hopefully it will try and fail to use the extended features
   if it's not present on your system. (And try and fail for different
   reasons if it's there.)
 
 * No automatic translation from LDAP-style addresses to RFC 282[12]
   (I don't even know if it's possible while connected to an Exchange
   server, but I'm not so it didn't seem worth trying); you have to
   provide a map. The script will report any untranslated
   addresses. Expect this to take some time if you're doing lots of
   emails; I have about 850 mappings for emails spanning seven years,
   which is about three times longer than the code.
 
 * Sometimes emails addresses don't have domain names. I don't really
   care, but I guess you could extend `fixup()` and give them a
   default domain if it bothers you.
 
 * There are encoding problems. I don't really understand what's going
   on, but characters I'm not expecting are sneaking in all over the
   place. Also, `CR` characters keep on cropping up at the end of
   some, but not all, lines. This doesn't stop me getting at the
   actual data, so I'm not overly concerned myself.
