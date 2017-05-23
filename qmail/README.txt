		
Welcome to Qmail !	

--> To Get started make a copy of demo qmail.ini and rename it qmail.ini
      then compile up a copy of qmail and start it!
    
    Now right click on the main window and choose options...

    (
     If you get an error on load it is probably because you dont have a D:
     drive on your system. edit the file paths in the ini and change default
     path to c: for default folders...next revision will be database driven
     app
    )
    
    It will Shell notepad with the ini file...fill in your account info
      and scroll down through the ini and see if there is anything else
      you want to change. Most options are preset for win98 systems..if
      you are on NT then you will have to change the path to wordpad..this
      is only used for mails with > 50000 characters which a textbox cant handle
    
    when you save changes and exit notepad qmail will know and will reload all 
      of your new preferences :)


Author: dzzie@yahoo.com
Sight : http://www.geocities.com/dzzie
Date  : Jan 2001
Time  : > 200 hrs lost track with all revisions + additions


-- Things To Know --
   Power features -
        frmRead - highlight an email address and hit reply and new message pops up with
                  that rcpt.
                 
                  Highlight a text selection and hit reply and it will be transfered to
                  msg body as reply text

                  highlight a base64 of quoted printable text and select decode text and
                  have the message decoded inline and put back in the text area. Also 
                  removed all html formatting.

        frmMessages - if you have an email address in the clipboard and hit new message
                      that email will be automatically used for the To: value

   frmRead -> rtclick -> decode text 
	decodes selected block of text inline and replaces it with the decoded text.
	This uses some smart logic. If it is base64 encoded it will unmime it, it is its 
	htmltext or quoted printable it will reformat it.

   frmRead -> rtclick -> decode file
	copies selected text of MIME encoded data to a temp file in attachments folder
	

   Inifile -
	All the ini file entries should be pretty straightforward by name. If you see
	an entry named number= that specifies how many configurations to look for in
	that section. like [bitfiles], number=2, bitfile1=xxx, bitfile2=yyy..Choose
        frmMessages ->rt click ->options to shell file, will wait until notepad quits
        then will reload preferences for you.

   Bitfiles -
	These are meant to save tidbits from emails...you can have as many as you want
	based on common topics you recive mails about. Highlight the text you want to 
	save, then rtclick in frmRead ->To BitFile -> whichever bit file you choose.
	the file will be opened for append and the new data added with ------- delimiters
	
   Multi Rcpts -
	You can have as many recipiants as you wish per email..just comma delimit thier
	email addresses in the TO or CC fields..To insert new rcpts from your addressbook
	just right click in the To: or CC: fields to select thier names.

   Attachments - 
	Are supported...up to 1.14Mb...most mailboxes wont let you go above a 2mb message
	anyway..and i limit it more cause i hold the whole msg blob in memory just before
	send. You can add a SINGLE attachment to any mail by dragging the file icon and
        dropping it in the Atch: text box. (they will all be declared right now as zips
        lets see what kinda funnies that plays with ole outlook :)\

   Sending Emails -
	Right click on messgaebody textbox for popup menu...spell check needs you to
        have M$ word installed (no way am i going to program that in from scratch)

   Incorperating mailto: support with IE.
        I have included the reg file you can integrate with your own IF YOU ARE ON WIN2k
        you are on your own ifyou are on win98 i dont know if they are the same but it
        could really make a mess if you get it wrong! If you mess up your computer you
        have been warned! You are probably going to have to edit the path to the exe
        also..rember to always escape your \ 's ..if you dont understand how to read
        .reg files dont mess with it until you read more about them.


-- TO DO --

       add in progress bar to frmSend and have message arg be a textfile it can read
       in sections for large attachment files...(not to necessary may never implement)

       possibly addin ability to use external command line pgp program

       i do NOT intend to ever make this parse attachments automatically...i want to
       have to manually decode whatever i choose so i have a choice in the matter...
	

-- Final Notes --

all of the "To Do" list are really quite minor to implement and are of no priority
to me...so dont wait on me to implement them.

this program is fully debugged and 100% functional with what is here...i get anywhere
between 30-130 mails a day so i can vouch for it. (This heavy mail load is why i needed
a custom client to begin with)

If you are on alot of mailing lists i think you will find this client has everything
you need to save important bits, archieve as mails as txt documents (use in conjunction
with chm-spider to save them all in a browsable database!)...anyway

everything in this program except the b64.bas was coded from scratch...no help...
nothing..all original..i originally did code my own base 64 bas file but it was
dredfully slow so i opted for a PSC one + 7 hours debugging :-\

the ini.bas and simple-fso.bas are worthy of the archieves *nods* 

anyway enjoy...feedback is always appreciated and I am impressed if you actually
read this file all the way down to here good boy *pats head* :)


-- CHANGE LOG --

august 01: Ported DLL to compile in to save memory..mails now written to disk 
	   as they come off the wire, only file names passed back from function.

august 16: found small bug that fired err msg when downloading some attachments
	   when string 'OK' was found in first 4 chrs of mime encoded block of
	   an incoming packet (happened mabey 1/30 attachments)

august 24: major overhaul..removed compile in fso and replaced with single bas file
	   simple-fso, removed frmProgress, File decodes shell to cmdline unmime.exe
	   fixed minor bug in unmime routine...added rt click menu to frmRead txtbox
	   streamlined frmRead buttons/layout..added smart reply to/save logic to
           detect and react to selected enties...now has option to include msg body
	   in reply..it is coming very close to a final release!

sept   16: removed frmOptions because i have no intentions of finishing it. I will
	    leave it in the zip. It is 90% done but wont edit accounts, xheaders,
	    fonts, or bit files. Once i am done with additions/ changes i might
	    finally finish it but it is pretty striaght forward so editing the ini 
	    file by hand is totally acceptable especially since it is an infequent act.

oct     1: changed format of xheaders from object for strings array..frmoptions will
	    now need to be modified before you can re-add it to project.

oct    15: Added attachment sending functionality , CC , multi-recpt, spell check,
            add rcpt, send again, Save (NoSend), insert rcpt, queryunload to keep
	    from loosing work accidently (only after loosing 3 page email of course):

oct    17: huge update to table of contents routines...added 2 more listviews now lvs
           are updated in memory exclusivly and then new tocs written on exit..saves
           many many multiple ogh so lots (did i mention lots?) of file writes..believe
           it or not it actually simplified the code by about 200 lines : D

Oct    24: changed implementation of multiple account checking procedure to how outlook
           handles it. You can now check each account seperatly, or can check all 
           accounts at once :D shift-Z in frmMessages checks default account

Oct    25: when you choose options now shell and wait then reloads ini file and prefs
           when notepad quits.

Oct    31: made longtimer.ctl and added in functionality for automatic mail checking

Nov     2: removed the external C mime and unmime exe's because shell() would puke
           if the command string got to long !...so had to use a PSC one...it is
           pretty modified from teh original 7+ hrs debugging yuk but now it works
           perfect...md5 hash and all.

Nov     8: added in progressbar on frmSend and timeout timers on frmSend and frmCheck



-dzzie


       	
			
