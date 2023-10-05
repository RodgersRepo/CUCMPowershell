# CUCM Powershell script examples

A collection of PowerShell scripts that demonstrate using REST and SOAP to interface with CUCM.

### Scripts

**CucmActiveCalls.ps1** - Script that creates a real time season to each CUCM subscriber, then displays the active calls. For more information about the Performance Monitor API (which this script accesses), Google the term **cucm perfmon api devnet**.
Change lines 72 to 75 in the script to the IP addresses for up to four CUCM servers executing the Call Manager service.
[![CUCMCall-Mon-Scree-Shot.png](https://i.postimg.cc/gJ1sTCwb/CUCMCall-Mon-Scree-Shot.png)](https://postimg.cc/N97RyP7d)
![Figure 1 - CUCM Active calls script screen shot](/./CUCMCallMonScreeShot.png "CUCM Active calls script screenshot")

## Installation

Click on the link for the script above. When the PowerShell code page appears click the **Download Raw file** button top right. Once downloaded to your computer have a read of the script in your prefered editor. All the information for executing the script will be in the script synopsis.

# Usage

To execute the PowerShell scripts in this repository. Save the ps1 file to a folder on your computer, then from a powershell prompt in the same folder.
```sh
Run .\scriptname.ps1 
```
Change `scriptname.ps1` for the name of the script above.

If your Windows enviroment permits, you could create a shortcut to the script. Paste the following line into the shortcut.
```sh
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\<PathToYourScripts>\CucmActiveCalls.ps1"
```
Then just double click the shortcut like you are starting an application. Check the correct path to the  powershell executable on your system.

## Credits and references

#### [AXL API with SOAPUI and Powershell](https://www.youtube.com/watch?v=tb9hINfg2nY&list=LL&index=10&t=421s)
A really great guide about how to get started sending AXL requests. Totaly recomend this video.

----
