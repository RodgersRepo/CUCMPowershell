# CUCM Powershell script examples

A collection of PowerShell scripts that demonstrate using REST and SOAP to interface with CUCM.

### Scripts

**CucmActiveCalls.ps1** - Script that creates a real time season to each CUCM subscriber, then displays the active calls. For more information about the Performance Monitor API (which this script accesses), Google the term **cucm perfmon api devnet**.
Change lines 72 to 75 in the script to the IP addresses for up to four CUCM servers executing the Call Manager service.

![Figure 1 - CUCM Active calls script screen shot](/./CUCMCallMonScreeShot.png "CUCM Active calls script screenshot")

## Installation

Click on the link for the script above. When the PowerShell code page appears click the **Download Raw file** button top right. Once downloaded to your computer have a read of the script in your prefered editor. All the information for executing the script will be in the script synopsis.
The developement system this script was tested on used self signed CUCM certificates. To stop certificate errors, a Dot NET core class is used within the PowserShell script to overcome this error. This means you will need Dot NET installed on your target PC, or use signed certificates.
If you decide to use signed certificates comment out lines 78 to 94 of the PowerShell script, then change the CUCM IP addresses for fully quailified domain names.

## Usage

To execute the PowerShell scripts in this repository. Save the ps1 file to a folder on your computer, then from a powershell prompt in the same folder.
```sh
Run .\scriptname.ps1 
```
Change `scriptname.ps1` for the name of the script above.

If your Windows enviroment permits, you could create a shortcut to the script. Paste the following line into the shortcut.
```sh
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\<PathToYourScripts>\scriptname.ps1"
```
Then just double click the shortcut like you are starting an application. Check the correct path to the  powershell executable on your system.

## Credits and references

#### [AXL API with SOAPUI and Powershell](https://www.youtube.com/watch?v=tb9hINfg2nY&list=LL&index=10&t=421s)
A really great guide about how to get started sending AXL requests. Totaly recomend this video.

----
