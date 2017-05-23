This source should be pretty clean and readable.

The basic winsock code was taken from an open source dll
that I have built up and added in object parsing routines
to make it cleaner to use as well as recoded portions of it.
See the credits and changelog at the top of the connection.bas file.

Connection Strings:

	This program will be able to connect with 2 different style
   connection strings.

	ftp://user:password@server.com:port
	ftp://server.com:port/some folder/

   if user:password is not supplied it will assume anonymous login

   if port is not supplied it will assume default port 21

   I also implemented a way so that you can integrate a custom protocol
   with IE so that when you want to ftp to a sight from teh browser you can just
   make a url with the name like fudd://
   
   I tried to take over the ftp:// protocol but it screwed up so bad IE lost some 
   functionalities...the ftp:// is more than just set in teh registry it is ingrained
   into IE and inet, urlmon etc : 0-_

   See below on how to set up the registry to add the fudd: protocol it isnt going to
   be to useful but it might be for some stuff :(

Usage:

	We already covered the connection strings so now all that is left is
	a couple user features. 

	Uploading:  You upload By dragging & dropping files in the treeview 
		from the desktop or an explorer window. If you have to 
		navigate and find a file on your system it is easier to do it
		in a full window than a common dialogue i say. This has a certain
		bug with it though right now..for some reason this freezes the
		explorer window until the upload is finished :( anyone know why?
		
		    If you try to upload a file with teh same name as one already
		on the server the remote file will be automatically deleted first.

		    If you want to append a file on the server to resume
           	a broken upload...right click -> set mode -> APPE
           	now this option will be checked and if you try to
           	upload a file with a same filename already on the
           	server and yours is larger..the remainder will be
           	be uploaded...it isnt used much so it isnt default

		Note you cannot upload or download files from the Quote command.


	Downloading: To download first set your download directory. This is
		a folder path saved in teh registry where all downloads will be
		saved to. I did it this way for convience sake. Also all downloads
		will be saved under there original name. This will not overwrite
		an existing file of the same name. If the original is smaller
		it will prompt you if you wish to resume the download.

		You set the dl dir from teh right click menu.

		If you have set a default bookmarked sight then you can drop a
		file on the programs icon to have it automatically uploaded for
		you. See Bookmarks section on how to do this.


	Raw command: The bottom most text box is for raw command interface. Just added
		this 01/18/02. Cant GET from it yet..but comeon how hard is it to
		double click on the list box? this is because need to know file size
		to call downlaod routine...we can deal with it eghh..

	Log window: This is the lower text box on the screen. It contains all the
		Server replies sent over the control connection. You will not see
		the data transfers like directory permissions and such here. If
		you want to stop it from scrolling and read it just double click
		inside it and its text will pop up in a new window.

	Bookmarks: You can save common ftp sights for easy use latter. Just enter
		the connection string right click and choose save. It will then prompt
		you for a name to call it by and it will be added to the combo box.
		These bookmarks are saved in the file ftp.txt that will be created
		in app.path when you add your first bookmark. This is a plain text
		file. This program also supports setting a default bookmark so that
		y9ou can upload files by dropping them on the exe icon. To do this
		you must manually edit the ftp.txt file and add a '*' character
		to the beginning of the bookmarks name you assigned (or when you
		name the bookmark just start the name with an astrics.)
	
	QuickView: If you want to read a text file on the server just highlight the
		file then right click and choose quickview. The File will be 
		downloaded and saved in memory only (not written to disk) and will
		be displayed in a built in text box form.
			
		   New to quick View in this latest build -> Just added a new function
		in ftp code so you can now upload changes you make to files in quickview
  		without ever having to save teh file to disk and edit it :D

		   After you have made your changes just right click and choose upload
		changes :)

	Search:	  This will let you search the current directory listing for file
		names that match the string you provide. I have included support
		for conventional * wild cards so *.gif would match any file that
		contained teh string .gif anywhere in the string as would *.g*f.
		   After you preform the search the listbox will be filled with
		the results. Now you can select the ones you want to keep or delete
		by right clicking on the list box and choose to delete the selected
		or unselected. Then you can hit the download or delete buttons and
		after confirming your selection the commands will be processed. There
		is no abort once you start so be sure of what you are doing.

	Setting up Custom Protocol for IE to shell
		I named the protocol to use here as fudd:
                Set up your reg key to look like this..if you dont understand this
                layout dont attempt the modificatio you have been warned.

		I have included a reg file from my win2k machine with IE5+
                make sure the paths are right if you use it...make sure you 
                know what you are doing if you use it...if you mess up your
                registry you are completly on your own and your system may not 
                work right without a complete reinstall of windows...so know what
                you are doing ! (even i messed up my registry trying to get it to
                take over the handeling of the ftp: protocol so you have been warned!)
		
                If anyone successfully gets it to take over the ftp protocol let me know
                how !

		HKCR create a new key 'fudd'
		   default value = 'URL:Fudd protocol'
		   stringvalue named 'Url Protocol'
			new key 'shell'
				new key 'open'
					new key 'command'
						default value= "<pathto prog>" %1

	

		
Bugs:

	There is a subscript out of range error that fires when you connect
to some servers. I tracked this down to some servers sending quota data after
a brief delay from normal login response
	
	The error is caused because we already sent a Ftp data command and expecting
a certain response which the quota data most certainly is not :-\

	I have added a timer to frmMain to delay showfileList after login 700ms
which for me on DSL is enough. I will create a more robust fix in time. This isnt
to big a deal because worst case scenario is you just refresh file list.
	
	I have transfered about a gig with this program and have had not one 
corrupt transfer.

	Last thing to note is that it dosent handle links right (basically like
windows shortcuts or alias)...if a file appears with permissions starting with "l"
then this is what it is...you can use them but you will have to do it manually 
by changing directory with the quote command and knowing how to read the name of
the link from the name shown in the treeview. You dont run into these a whole lot
so not a big deal.

	Anyway...have fun with it...I know we could have used the inet.dll api to
code this alot simplier without all the winsock stuff...but there is no fun in 
that.  If you want to understand the protocall and how it works...you have to see
the code itself not just teh wrapped function calls. I learned alot from this and
i really thank Neotext for sharing up his dll. I dont know if he borrowed any of
the code from anywhere else but at anyrate it was very clean and to the point and
helped me really understand the protocall.
