To bring Smallserver online you have to drag and drop a file 
into the ServeFile textbox.

If you any other web server running on port 80, it will 
probably crash, so make sure no other web server is currently
running.

Text, HTML and Image files (jpg,gif) files will be served
as normal web server would logging requests.

Smallserver also supports some other custom file types which
allow it to preform some custom actions. The /Scripts directory
contains samples of these files types briefly outline below.

- .auth file - when a request is made to the server it will pop up
a http auth dialog with the custom realm information you specify thhen
redirect them to the URL provided. This is used to try to fake users
into providing you with their login information.

- .loc file - when the server receives a request, it will automatically
use a http redirect them to the URL you priovide in the file, this is 
used to transparently redirect users to another resource with out you
having to chew up your bandwidth serving up actyual files

- .raw file - sends the actual http headers and data you supply in 
the file directly back to teh browser

using the raw checkbox also allows you modify raw headers returned to
browser

Other features of small server...if Auth checkbox is checked, then the
server will only serve the speciofoed fiole for the Auth URL textbox
URL..this is to prevent people from tryign to probe the server . If
someone is trying to probe server...all of thier requests will be redirected automatically to probes url. 

If this option is not enabled then any request to the server for any
URL will result in teh same serve file being sent back  

Last thing to mention is design feature...if the server detects
that it is telenetter trying to manually send request to server
then it will spam their terminal, disconnect them and alert you.

It is a handy tool with alot of possibilities...used creativly it is 
great tool.

Have fun :)