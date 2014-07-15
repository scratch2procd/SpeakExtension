Scratch Speak Extension by Procd

This extension provides text to speech blocks for the Scratch offline editor.

This extension is for use with the Scratch offline editor. It has been written in C# and runs on Windows. The extension is set to run on port 8080 by default. To change, then you will have to edit the s2e file and run the SpeakExtension.exe with a command line paramter
SpeakExtension -p=<port number>
where <port number> is the new port number.

*** Important : ***
The SpeakExtension must be run as Administrator otherwise you will get an access denied exception. If you don't want to do this then you will have to use netsh to provide the necessary rights to the SpeakExtension URL. Much easier to just right click the exe and select Run As Administrator!

The extension is a console application and will only end once return is pressed.

You will need to tell Scratch about the extension so shift + click the File menu in the Scratch offline editor and select "Import Experimental Extension". Find the Speak.s2e file.
You should now have new blocks in the more blocks category. You should also have a green dot to show that the SpeakExtension has been connected. If it's red ensure the extension is running and ports are correct.

You can test the extension without scratch by using a browser. The url for port 8080 would be

http://localhost:8080/speak/hello/

You should hear the computer say "hello"

The new blocks the extension provides are;

     Speak %s
     Speak %s and wait  ***NOTE: Scratch has a bug so does not send the command to the extension, so this block does not work ***
     Set Volume %n
     Change Volume by %n
     Set Rate %n
     Change Rate by %n

    and reporters blocks;

    Volume  [0.100]
    Rate    [-10..10]
    Gender  ["male, female", "other"]

The voice the speech synthesizer uses is the voice used for Windows "Ease of Access" which can be changed in the control panel. Other voices can be downloaded so you can change the voice age and gender in  the control panel if required and the extension will use that.

Happy Scratching.

(c) 2014 Procd

