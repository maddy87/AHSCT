# AHSCT
Automated Word Notification Tool
Windows Application in C# with DB connections to Interface

Clients were supposed to recieve notifications detailing outages of an specific outage with the detail history of the problem
periodically. The manual process was tedious and erroneous and much of the time was wasted formatting the whole data.
The tool enables the analyst to work on the problem at hand an spend minimum time creating and sending notifications.

Design : 

 Splash screen to indicate application connectivity.
 Sleek Modular design with importance to process than convinience.
 Limited and only desired interaction
 Tabbed interface for easy navigation
 Hide on minimized to avoid cluttering on taskbar
 Seamless flow through the application hence least guidance expected

Functionality : 

  Threaded environment allowing multiple threads load datasets when the application is launched
  Dynamically updates the templates with the information in the desired bookmarks location and delete the update the    information without changing the format.
  Exporting Data into Excel Format
  Sending Feedback to the developer
  ErrorReportingModule just in case there is an error occured
  Sending Email by trigerring macros inside the word document.
  Using Microsoft InterOp Assemblies perform spellcheck on the input provided
  Get the bookmarks intialized in the document and dynamically updating their   content with the one's provided by the programmer.
  
 The empty Database File DBAHSCT.mdb is uploaded with the DB design and table structure intact however the details of the 
 applications have been removed as it was confidential to the organization.

