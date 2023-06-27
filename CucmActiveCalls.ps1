<#
.SYNOPSIS
  Name: CucmActiveCalls.ps1
  This script polls each of four Subscribers for the amount of active calls. Uses
  the CUCM Performance Monitoring API.
  GUI elements created using
  XAML and windows presentation framework (wpf).
  Uses threading to seperate the GUI from long running HTTPS requests code.
  Stops GUI freezing. See below URL for good explanation
  https://smsagent.blog/2015/09/07/powershell-tip-utilizing-runspaces-for-responsive-wpf-gui-applications/
  PoSh version 5.1.14393.206 and above.

 
.DESCRIPTION
  The script obtains a token from CUCM to create a real time session. The
  token is realeased when the script is 'Canceled'. Google 'cucm perfmon api devnet'
  to read the Cisco document. From the cisco document this script undertakes the 
  following sequence of operations:-
  
  1.perfmonOpenSession
  2.perfmonAddCounter (in this case Active Calls)
  3.perfmonCollectSessionData (repeated)
  4.perfmonCloseSession
 
.NOTES
  Copyright (C) 2023  RodgeIndustries2000
 
 
     This program is free software: you can redistribute it and/or modify
     it under the terms of the GNU General Public License as published by
     the Free Software Foundation, version 4.
 
     This program is distributed in the hope that it will be useful,
     but WITHOUT ANY WARRANTY; without even the implied warranty of
     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
     GNU General Public License for more details.
 
     To view the GNU General Public License, see <http://www.gnu.org/licenses/>.

    Release Date: 19/04/23
    Last Updated:        
   
    Change comments:
    Initial realease V1 - RITT
   
   
  Author: RodgeIndustries2000
       
.EXAMPLE
  Run .\CucmActiveCalls.ps1 <no arguments needed>
  Or create shortcut to:
  "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\<PathToYourScripts>\CucmActiveCalls.ps1"
  Then just double click the shortcut like you are starting an application. Check the correct path to the
  powershell executable on your system.

#>


#-----------------[ Declarations ]----------------------------------------------------#

Set-StrictMode -Version Latest
$sync = [Hashtable]::Synchronized(@{ })                             # Syncronised hash table for talking accross runspaces.
                                                                    # Multiple threads can safely and efficiently add or remove
                                                                    # items from this type of hash table
                                                                    # This hash table stores node names from the XAML below
                                                                    # hash tables are key/value stored arrays, each      
                                                                    # value in the array has a key. Not Case Senstive
$ErrorActionPreference = "Stop"                                     # What to do if an unrecovable error occures
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Path         # This scripts path
$scriptName = $MyInvocation.MyCommand.Name                          # This scripts name
$global:cucmIpAddr = @(
[ipaddress] "10.10.1.4"
[ipaddress] "10.10.1.5"
[ipaddress] "10.10.2.4"
[ipaddress] "10.10.2.5"
)                                                                   # An array of CUCM addresses, adjust for your enviroment

$addDotNetCoreClassCertCheck = @"
 using System.Net;
 using System.Security.Cryptography.X509Certificates;
 public class TrustAllCertsPolicy:ICertificatePolicy {
    public bool CheckValidationResult (
        ServicePoint srvPoint,
        X509Certificate certificate,
        WebRequest request,
        int certificateProblem
    ){ return true; }
 }
"@

# Add the dotNet cert check above. Accepts unverifiable certs, like the potential
# security risk ahead browser warning you get. Instead of asking the user just accepts
Add-Type -TypeDefinition $addDotNetCoreClassCertCheck
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy​

# Add these assemblys
Add-Type -AssemblyName presentationframework, presentationcore      
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# Negotiate TLS version start at 1.1 up to 1.2
[System.Net.ServicePointManager]::SecurityProtocol = 'Tls11, Tls12'

######################################################################################
#       Here-String with the eXAppMarkupLang (XAML) needed to display the GUI        #
######################################################################################

# A here-string of type xml
[xml]$xaml=@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="CanResizeWithGrip" Name="activeCallsgui"
        Title="CUCM Active Calls Script - Version 1.0" Height="650" Width="1200"
        FontSize="17" FontFamily="Segoe UI">

   <Window.Resources> <!--Match name with the root element in this case Window-->
        <!--Setting default styling for all buttons-->  
        <Style TargetType="Button">
         <Setter Property="Width" Value="143" />
         <Setter Property="Height" Value="32" />
         <Setter Property="Margin" Value="10" />
         <Setter Property="FontSize" Value="18" />
         <Setter Property="Background" Value="#FFB8B8B8" />
        </Style>
        <Style TargetType="TextBox">
         <Setter Property="Background" Value="#FFB8B8B8" />
         <Setter Property="Height" Value="32" />
        </Style>
        <Style TargetType="ComboBox">
         <Setter Property="Background" Value="#FFB8B8B8" />
         <Setter Property="Height" Value="32" />
        </Style>
        <Style TargetType="PasswordBox">
         <Setter Property="Background" Value="#FFB8B8B8" />
         <Setter Property="Height" Value="32" />
        </Style>
     </Window.Resources>

    <Grid>
     
      <Grid.RowDefinitions>
        <RowDefinition Name ="Row0" Height="41*"/><!--Row 0 Row Heights as percentage of entire window-->
        <RowDefinition Name="Row1" Height="50*"/> <!--Row 1-->
        <RowDefinition Name="Row2" Height="9*"/> <!--Row 2-->
      </Grid.RowDefinitions>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="25*"/>           <!--Column 0-->
        <ColumnDefinition Width="25*"/>
        <ColumnDefinition Width="25*"/>
        <ColumnDefinition Width="25*"/>
      </Grid.ColumnDefinitions>

      <DockPanel>
        <Menu DockPanel.Dock="Top" Background="#FFFFFFFF">
            <MenuItem Header="_File">
                <MenuItem Header="_About" Name="menuItemAbout"/>
                <MenuItem Header="_Save Console Messages" Name="menuItemSaveAS" IsEnabled="False"/>
                <Separator />
                <MenuItem Header="_Exit" Name="menuItemExit"/>
            </MenuItem>
        </Menu>
      </DockPanel>
       
      <GroupBox Name="instructionsGrpBox"  Header="Instructions" Margin="10" Padding="10" Grid.Row="0" Grid.Column="0" Grid.ColumnSpan="4" Grid.RowSpan="2" Visibility="Visible" >
             <StackPanel>
                <TextBlock Foreground="teal" FontSize="30" TextWrapping="Wrap" VerticalAlignment="Center" HorizontalAlignment="Center">
                  CUCM Active Calls Monitor <LineBreak />
                </TextBlock>

                <TextBlock Name="instructionsTxtBlk" TextWrapping="Wrap" VerticalAlignment="Center" HorizontalAlignment="Center">
                   This script creates a real time season to each CUCM subscriber, then displays the<LineBreak />
                  active calls.<LineBreak />
                  Toggle the <Bold>GO</Bold> button to start/stop the active calls counter.<LineBreak />  
                  For more information about the Performance Monitor API (which this script accesses) see.<LineBreak /><LineBreak />
                  https://develeoper.cisco.com<LineBreak />
                  Or Google the term <Bold>cucm perfmon api devnet</Bold><LineBreak /><LineBreak />
                  Press the <Bold>Next</Bold> button to continue. Press the <Bold>Cancel</Bold> button at any time to quit this script<LineBreak />
                  Please be patient when the script closes. The script attempts to close any open CUCM sessions prior to closing the GUI.<LineBreak />
                </TextBlock>
             </StackPanel>
      </GroupBox> <!---->

   
    <GroupBox Header="Enter your CUCM Credentials here" Name="credsGrpBox" Margin="10" Padding="10"  Grid.Row="0" Grid.Column="0" Visibility="Hidden" >
        <StackPanel>
        <TextBlock>Username:</TextBlock>
        <TextBox Name="credsTxtBox1" />
        <TextBlock>Password:</TextBlock>
        <PasswordBox Name="credsTxtBox2" />        
        </StackPanel>
    </GroupBox>

    <GroupBox Header="Console Messages will appear here" Name="cmdGrpBox" Margin="10" Padding="10" Grid.Row="0" Grid.Column="1" Grid.ColumnSpan="3" Visibility="Hidden">
        <Grid Name="resultsGrid">
         <ScrollViewer>
           <TextBlock Name="Output_TxtBlk" TextWrapping="Wrap" TextAlignment="Left" VerticalAlignment="Stretch" />
         </ScrollViewer>
        </Grid>
    </GroupBox>

    <GroupBox Header="Sub 1 Active calls" Name="resultsGrpBoxSub1" Margin="10" Padding="10" Grid.Row="1" Grid.Column="0" Visibility="Hidden" >
        <StackPanel>
            <Label Name="sub1Label" Content="0000" HorizontalAlignment="Center" VerticalAlignment="Top"  FontSize="95" />
            <Button Name="sub1Go" HorizontalAlignment="Center" VerticalAlignment="Bottom">GO</Button>
            <Button Name="sub1Stop" HorizontalAlignment="Center" VerticalAlignment="Bottom" Visibility="Hidden">STOP</Button>
        </StackPanel>    
    </GroupBox>

    <GroupBox Header="Sub 2 Active calls" Name="resultsGrpBoxSub2" Margin="10" Padding="10" Grid.Row="1" Grid.Column="1" Visibility="Hidden" >
        <StackPanel>
            <Label Name="sub2Label" Content="0000" HorizontalAlignment="Center" VerticalAlignment="Top"  FontSize="95" />
            <Button Name="sub2Go" HorizontalAlignment="Center" VerticalAlignment="Bottom">GO</Button>
            <Button Name="sub2Stop" HorizontalAlignment="Center" VerticalAlignment="Bottom" Visibility="Hidden">STOP</Button>
        </StackPanel>
    </GroupBox>

    <GroupBox Header="Sub 3 Active calls" Name="resultsGrpBoxSub3" Margin="10" Padding="10" Grid.Row="1" Grid.Column="2" Visibility="Hidden" >
        <StackPanel>
            <Label Name="sub3Label" Content="0000" HorizontalAlignment="Center" VerticalAlignment="Top"  FontSize="95" />
            <Button Name="sub3Go" HorizontalAlignment="Center" VerticalAlignment="Bottom">GO</Button>
            <Button Name="sub3Stop" HorizontalAlignment="Center" VerticalAlignment="Bottom" Visibility="Hidden">STOP</Button>
        </StackPanel>    
    </GroupBox>

    <GroupBox Header="Sub 4 Active calls" Name="resultsGrpBoxSub4" Margin="10" Padding="10" Grid.Row="1" Grid.Column="3" Visibility="Hidden" >
        <StackPanel>
            <Label Name="sub4Label" Content="0000" HorizontalAlignment="Center" VerticalAlignment="Top"  FontSize="95" />
            <Button Name="sub4Go" HorizontalAlignment="Center" VerticalAlignment="Bottom">GO</Button>
            <Button Name="sub4Stop" HorizontalAlignment="Center" VerticalAlignment="Bottom" Visibility="Hidden">STOP</Button>
        </StackPanel>    
    </GroupBox>

    <StackPanel Name="buttonCancelStackPanel" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="3" HorizontalAlignment="Right" Orientation="Horizontal" >            
            <Button Name="myButton1" Content="Next"  />
            <Button Name="myCancelButton" Content="Cancel" />
    </StackPanel>
    </Grid>
</Window>
"@

#------------------[ Functions ]------------------------------------------------------#
#######################################################################################
#        Function that open/manages a session with CUCM perfmon web service           #
#       never actually gets called, it is read into $codeToExec of function           #
# Start-NewRunspace as a scriptblock. The new runspace executes the code as a code    #
# block, this stops the GUI freezing. New runspace managed via the sync hash table    #    
# $sync                                                                               #
####################################################################################### 

function Start-CucmSession 
{
    param ($goButton, $countLabel, $cucmSubIpAddr, $creds)

    # thinking of how to kill the PS instance if it all goes wrong
    # look more into this
    # $y = Get-PSHostProcessInfo
    #$y = $x | Where-Object -FilterScript {$_.MainWindowTitle -Like "CUCM Realtime CallsActive - Version 1"}
    #$y.ProcessId

    # Performance monitor name for this instance to save to the sync hash table
    # follows the format $perfmonsubXGo, $tokensubXGo where X is the subscriber number
    $perfMon = "perfMon" + $goButton
    $cucmSessionToken = "token" + $goButton
    $stopButton = $goButton -replace "Go$", "Stop"
    
    try
    {
        # the url for the cucm perfmon web service description language (wsdl) file
        # from this file powershell can find the correct methods and properties
        # to interface with cucm perfmon. Takes a bit of time to download, thats why
        # the first transaction takes a while
        $urlPerfMon = "https://" + $cucmSubIpAddr + ":8443/perfmonservice2/services/PerfmonService?wsdl"
        
        # New-WebServiceProxy contacts CUCM subscriber and returns a series of objects,
        # these can be stored in the hash table $sync and accessed by other powershell instances
        # infact $sync is the perfect place to store any object that you want to access later
        $sync.$perfMon = New-WebServiceProxy -Uri $urlPerfMon -Credential $creds

        # Refresh the GUI Output_TxtBlk via dispatcher to update console mesages.
        # Open a session with the CUCM subscriber, be careful, everytime you access this method
        # $sync.$perfMon.perfmonOpenSession() you start another session, even if you put Value
        # on the end rather than show existing value just creates a new instance
        $sync.Output_TxtBlk.Dispatcher.Invoke([action]{$sync.Output_TxtBlk.Text += "Creating session to $cucmSubIpAddr`n"})
        $sync.$cucmSessionToken = $sync.$perfMon.perfmonOpenSession()
        $sync.Output_TxtBlk.Dispatcher.Invoke([action]{$sync.Output_TxtBlk.Text += "Connection success, token is " + $sync.$cucmSessionToken.Value + "`n"})
    
        # Get the namespace from the wsdl. Namespaces ensure each method in the wsdl has a
        # unique Uniform Resource Identifier (URI). Imagine if two wsdl's used the same method name
        # the only way an app can tell them apart is by using the full URI to each method
        $ns = $sync.$perfMon.GetType().namespace

        # This is where it gets annoying. Some of the wsdl methods will
        # only accept objects as properties. Why not just a plain old string or integer??
        # You have to create your own custom objects. This is the object SessionHandleType
        # any property defining the session token must be of this type
        #UPDATE-- YOU DONT NEED TO NAME THESE!!!
        # The session token is now in a format cucm will accept. Now we
        # can set up the counter/s we want to monitor as part of our session.
        # As before first thing to do is creat an object for the counter types
        $cucmCounterToAdd = New-Object -TypeName ($ns + ".CounterType")

        # As per the wsdl rules the CounterType object only accepts counter properties
        # if they are in an object with the type name CounterNameType. So you
        # need to create that too. Who designs this nonsense!!
        $cucmActiveCallsCounter = New-Object -TypeName ($ns + ".CounterNameType")

        # The counterNameType will only accept a list of counters as an array.
        # You can add as many counters as you like to the end of the array comma seperated
        # But I am only adding one counter, active calls
        $cucmActiveCallsCounter.Value = @("\\$cucmSubIpAddr\Cisco CallManager\CallsActive") 
        
        # Add the counters of interest to the counters you want to add to this session
        $cucmCounterToAdd.Name = $cucmActiveCallsCounter

        # Finally you are ready to send the session properties to cucm. That being
        # your session token (handle) and the counters you are interested in
        $sync.$perfMon.perfmonAddCounter($sync.$cucmSessionToken, $cucmCounterToAdd)
        
        # Collect and display the Active Calls value until the STOP button pressed
        # cucm will throw an exception if requests for values exceds 50 requests
        # a minute.
        Do
        {
            $activeCallsValue = "{0:d4}" -f $sync.$perfMon.perfmonCollectSessionData($sync.$cucmSessionToken).Get(0).Value            
            $sync.$countLabel.Dispatcher.Invoke([action]{$sync.$countLabel.Content = $activeCallsValue}) # Refresh GUI subscriber labels element
            start-sleep -seconds 5
        } 
        Until ($sync.$stopButton.Visibility -eq "Collapsed")       
    
    }
    catch 
    {
        $sync.Output_TxtBlk.Dispatcher.Invoke([action]{$sync.Output_TxtBlk.Text += "UNABLE TO CONTACT $cucmSubIpAddr`n $_`n"})
    }     
}

#######################################################################################
#        Function to open file save dialog, returns filename and path                 #
#######################################################################################

function Get-FileSaveDialog ($fileType="txt")
{
    $OpenSaveDialog = New-Object System.Windows.Forms.savefiledialog
    $OpenSaveDialog.initialDirectory = $global:scriptPath
    $OpenSaveDialog.filter = "$fileType files (*.$fileType)| *.$fileType"
    if ($OpenSaveDialog.ShowDialog() -eq 'Ok') {return $OpenSaveDialog.filename}  
}

#######################################################################################
#        Function to exit the script, called by the cancel buttons                    #
#######################################################################################

function Close-AndExit
{
    # Trying to adopt good housekeeping. By collapsing the stop buttons
    # will trigger a change in InvocationStateInfo, see Start-NewRunspace 
    # function below. That in turn runs $jobDoneScriptBlock which should
    # close any open CUCM http sessions 
    $sync.sub1Stop.Visibility = "Collapsed" 
    $sync.sub2Stop.Visibility = "Collapsed" 
    $sync.sub3Stop.Visibility = "Collapsed" 
    $sync.sub4Stop.Visibility = "Collapsed" 
    
    try {
        $sync.activeCallsgui.Close()
    }
    catch {
        # If an exception occurs, ignore it
        $null = $_
    }        
}

#######################################################################################
#      Function to start a new run space, creates new poSH instance containing code   #
#   from a function . Runs independantly of the GUI instance. Stops GUI freezing      #
#######################################################################################

function Start-NewRunspace ([string]$codeToExec, $goButton, $countLabel, $cucmSubIpAddr, $creds)
{
    # This scriptblock is invoked by the Register-ObjectEvent cmdlet at the bottom of this function. 
    # When the InvocationStateInfo of the new runspace triggers the completed event this code
    # closes any http sessions to CUCM. Just trying to maintain good housekeeping
    [scriptblock] $jobDoneScriptBlock = {
        if($Sender.InvocationStateInfo.State -eq 'Completed')
        {
            #$global:runspaceForHttp.Close()
            #$global:runspaceForHttp.Dispose()
            
            try
            {
                # Collect the name of the subscriber to stop
                # from the MessageData passed from Register-ObjectEvent
                # it will be 'subXGo' where the X is the subcriber number
                $runSpaceToStop = $Event.MessageData.goButton
                $perfMon = "perfMon" + $runSpaceToStop
                $cucmSessionToken = "token" + $runSpaceToStop

                # Pop up box not used, uncomment for troubleshooting
                #[System.Windows.MessageBox]::Show(
                #    $perfMon,
                #    'End of scriptblock',
                #    'OkCancel'
                #)

                $sync.Output_TxtBlk.Dispatcher.Invoke([action]{$sync.Output_TxtBlk.Text += "closing " + $sync.$cucmSessionToken.Value + "`n"})
                $sync.$perfMon.perfmonCloseSession($sync.$cucmSessionToken) # No need to put Value here as it wants an object not a string
                Get-Runspace | Where-Object {$_.Name -eq $runSpaceToStop} | ForEach-Object {$_.Close();$_.Dispose()}
                               
            }
            catch 
            {
                $sync.Output_TxtBlk.Dispatcher.Invoke([action]{$sync.Output_TxtBlk.Text += "Closing runspaces errors`n $_`n"})
            }  
        }
    }
    
    # Find which code is to be exec in new runspace
    # [powershell]::Create starts another poSH instance 
    # A job would start another poSH process with all
    # the baggage

    # Get the code to execute in the runspace from the function
    # name of the function is in variable $codeToExec
    # but for this script is always Start-MyCounter

    $code = Get-Content Function:\$codeToExec -ErrorAction Stop;

    # Create a new powershell instance outside
    # of the instance that controls the GUI
    # add any auguments to the code this instance executes

    $newPsInstance = [PowerShell]::Create().AddScript($code)
    $newPsInstance.AddArgument($goButton).
                AddArgument($countLabel).
                AddArgument($cucmSubIpAddr).
                AddArgument($creds)
    
    # Create a new runspace
    $runspace = [RunspaceFactory]::CreateRunspace()

    # Add the runspace object to the powershell instance
    $newPsInstance.Runspace = $runspace

    # Set the new runspace parameters then open
    # the runspace for use
    $runspace.Name = $goButton
    #$runspace.ApartmentState = "STA"
    #$runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    # Add the sync hash GUI table to the runspace
    # all runspaces can then manipulate this syncronised hash
    # table
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    
    # Invoking the new powershell instance executes the code
    # in the new runspace 
    $newPsInstance.BeginInvoke()

    # Monitor the new runspace for event changes, call the $jobDoneScriptBlock
    # should this run space trigger an event
    Register-ObjectEvent –InputObject $newPsInstance –EventName InvocationStateChanged –Action $jobDoneScriptBlock -MessageData @{goButton = $goButton} 
 }    

#######################################################################################
#      Function to check that the user has either entered some text in the txt        #
#######################################################################################

function Check-UserInput ()
{
    $resultBool = $false
    $messageTxt = "Please supply your CUCM username and Password!!"

    # check the credential text boxes have been filled out
    If (-not ([string]::IsNullOrEmpty($sync.credsTxtBox1.Text)) -and -not
             ([string]::IsNullOrEmpty($sync.credsTxtBox2.Password )) )
    {
        $resultBool = $true
        $messageTxt = ""
    }

    return $resultBool, $messageTxt

}

#######################################################################################
#         This function is called by the sub go event handler. Checks user input      #
#       by calling function Check-UserInput then launches the CUCM perfmon session    #
#######################################################################################

function Launch-CucmSession ($subscriberNumber)
{
    $formCorrect, $msgText = Check-UserInput
    if ( $formCorrect )
    {
        # Encode the users cucm credentials
        # $sync.credsTxtBox2.SecurePassword is the secure (unreadable) version of the Password string 
        $yourCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sync.credsTxtBox1.Text,$sync.credsTxtBox2.SecurePassword

        # Hide the GO button, show the STOP button
        $goButton = "sub" + $subscriberNumber + "Go"
        $stopButton = "sub" + $subscriberNumber + "Stop"
        $subLabel = "sub" + $subscriberNumber + "Label"
        $cucmSub = $global:cucmIpAddr[[int]$subscriberNumber - 1].IPAddressToString
        $subscriberNumber = "Subscriber" + $subscriberNumber        
        $someLines = $('-' * 26)
        $sync.Output_TxtBlk.Foreground = "Black"
        $sync.Output_TxtBlk.Text += "$somelines     BEGIN POLLING ACTIVE CALLS FOR $subscriberNumber    $somelines`n"
        $sync.Output_TxtBlk.Text += "$somelines NOTE: Initial connection can take 20 seconds to negotiate!! $somelines`n"
        $sync.$goButton.Visibility = "Collapsed" # Not visable, not taking up any space on the GUI
        $sync.$stopButton.Visibility = "Visible"
        Start-NewRunspace -codeToExec "Start-CucmSession" -goButton $goButton -countLabel $subLabel -cucmSubIpAddr $cucmSub -creds $yourCreds
    }
    else
    {
        $sync.Output_TxtBlk.Foreground = "Red"
        $sync.Output_TxtBlk.Text = $msgText + "`n"
    }
}

#######################################################################################
#    This function is called by the subXstop event handler. Where X is 1 o 4          #
#######################################################################################

function Halt-CucmSession ($subscriberNumber)
{
    # Hide the Stop button, show the Go button
    $goButton = "sub" + $subscriberNumber + "Go"
    $stopButton = "sub" + $subscriberNumber + "Stop"
    $countLabel = "sub" + $subscriberNumber + "Label"
    $subscriberNumber = "Subscriber" + $subscriberNumber    
    $someLines = $('-' * 26)
    $sync.Output_TxtBlk.Foreground = "Black"
    $sync.Output_TxtBlk.Text += "$somelines STOP POLLING ACTIVE CALLS FOR $subscriberNumber $somelines`n"
    $sync.$goButton.Visibility = "Visible"
    $sync.$stopButton.Visibility = "Collapsed" # Not visable, not taking up any space on the GUI
    $sync.$countLabel.Content = "0000" # Reset the counter
}

#----------------[ Main Execution ]---------------------------------------------------#
#######################################################################################
#               Read the XAML needed for the GUI                                      #
#######################################################################################

$reader = New-Object System.Xml.XmlNodeReader $xaml
#$myGuiForm=[Windows.Markup.XamlReader]::Load($reader)
$sync.Window=[Windows.Markup.XamlReader]::Load( $reader )

# Collect the Node names of buttons, txt boxes etc.

$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$sync.Add($_.Name, $sync.Window.FindName($_.Name))}

#######################################################################################
#               This code runs when the Menu Item about button is clicked             #
#######################################################################################

$sync.menuItemAbout.Add_Click({
        #Show the help synopsis in a GUI
        Get-Help "$global:scriptPath\$global:scriptName" -ShowWindow
})

#######################################################################################
#               This code runs when the Menu Item save console button is clicked      #
#                     This item is grayed out until GO button is pressed              #
#######################################################################################

$sync.menuItemSaveAs.Add_Click({
        #Save console data as a txt file
        $saveAsFilePathName = Get-FileSaveDialog
        If ( $saveAsFilePathName -ne $null )
        {
           $sync.Output_TxtBlk.Text | Out-File -FilePath $saveAsFilePathName
        }                   
})

#######################################################################################
#               This code runs when the Menu Item exit button is clicked              #
#######################################################################################

$sync.menuItemExit.Add_Click({
        #Call the close and exit function
        Close-AndExit
})

#######################################################################################
#               This code runs when the Cancel buttons are clicked                    #
#######################################################################################

$sync.myCancelButton.Add_Click({        
        #Call the close and exit function
        Close-AndExit
})

#######################################################################################
#               This code runs when the Next 1 button is clicked                      #
#   button 1 is the Next button. Takes you from the instructions to the main GUI form #
#######################################################################################

$sync.myButton1.Add_Click({
   
    $sync.instructionsGrpBox.Visibility = "Hidden"
    $sync.credsGrpBox.Visibility = "Visible"
    $sync.cmdGrpBox.Visibility = "Visible"
    $sync.resultsGrpBoxSub1.Visibility = "Visible"
    $sync.resultsGrpBoxSub2.Visibility = "Visible"
    $sync.resultsGrpBoxSub3.Visibility = "Visible"
    $sync.resultsGrpBoxSub4.Visibility = "Visible"
    $sync.myButton1.IsEnabled = $False

    # Clear all old txt blocks and entry fields
    $sync.activeCallsgui.Dispatcher.Invoke([action]{},"Render") # Refresh update the GUI
    $sync.Output_TxtBlk.Text = ""
})

#######################################################################################
#               This code runs when any of the GO Buttons are clicked                 #
#######################################################################################

$sync.sub1Go.Add_Click({Launch-CucmSession "1"; $sync.menuItemSaveAs.IsEnabled = $true})
$sync.sub2Go.Add_Click({Launch-CucmSession "2"; $sync.menuItemSaveAs.IsEnabled = $true})
$sync.sub3Go.Add_Click({Launch-CucmSession "3"; $sync.menuItemSaveAs.IsEnabled = $true})
$sync.sub4Go.Add_Click({Launch-CucmSession "4"; $sync.menuItemSaveAs.IsEnabled = $true})

#######################################################################################
#               This code runs when any of the STOP Buttons are clicked               #
#######################################################################################

$sync.sub1Stop.Add_Click({Halt-CucmSession "1"})
$sync.sub2Stop.Add_Click({Halt-CucmSession "2"})
$sync.sub3Stop.Add_Click({Halt-CucmSession "3"})
$sync.sub4Stop.Add_Click({Halt-CucmSession "4"})

#######################################################################################
#               Show the GUI window by name                                           #
#######################################################################################

# Provide your own closing event (Add_Closing), custom closing event
# copies code from the function Close-AndExit. Will
# close CUCM sessions before exiting.
# Invoke-Expression runs the set of commands in $codeToRunOnExit 
$codeToRunOnExit = Get-Content Function:\"Close-AndExit" -ErrorAction Stop
$sync.activeCallsgui.Add_Closing({
    Invoke-Expression $codeToRunOnExit
})

# Stops all closeing events from running!!
# even your cancel and exit routines stop working!!     
#$sync.activeCallsgui.Add_Closing({$_.Cancel = $true}) 

$sync.activeCallsgui.ShowDialog() | out-null # null dosn't show false on exit​
