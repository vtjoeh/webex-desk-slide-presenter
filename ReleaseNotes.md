# Release Notes

### Release Notes ver 0.1.5.5 

**New Dafault Setting Option:** 
- **Main Video Stream:Mute (default) or Mirror**  
  - Main video is muted by default when an immersive share is started.  Some users may want the main video stream to mirror the content channel.  
- **Support for USB Clicker with Webex Desk Pro**
  - Room OS 11 removed immersive share from the GUI when not in a call.  If you want to use immersive share while not in a call, the USB clicker will give you 4 limited feature presets. 
  - Tested USB Clicker: Amerteer Wirless Presenter USB (RF 2.5 Ghz).  
  - To turn on from web interface: Settings --> Configuration --> Peripherals --> InputDevice Mode: On
  - Immersive share presets: Volume Up,  Volume Down, Page Up, Page Down keys
  - Hyperlink Key (squre icon)
    - Tap once: Show Content Only 
    - Tap 2x: StopImmersiveShare (PC and main video separate)
    - Hold down: Stop sharing any content (main cam only) 
- **Fixed Bug** 
  - Fixed issue if the ‘CameraOnly’ command was the first command sent might cause issues  

### Release Notes ver 0.1.5.4

**Updates to the PowerPoint TemplateDeskMacro_ver0.1.5.4** (2022.07.19)
- **Name Changes**
  - **PowerPoint Template:** Changed name from TemplateDeskProMacro_verXXX.pptm to TemplateDeskMacro_verXXX.pptm since it works on Desk Mini and Desk.  
  - **NoVideo** command changed to **ContentOnly** for clarity.  NoVideo command still works for backwards compatibility.  
- **Ribbon: Webex Desk**
  - For PC only, added Webex Desk tab to ribbon with commands to insert VideoSquares and toggle shape visibility.  Ribbon does not work on Mac.  PowerPoint templates created on the PC by the ribbon commands will work on Mac. 
- **Fixed Bug in PowerPoint** 
  - Password character fix.  
    - Password had problems with the following characters:   $ ” + ’ 
<br/>They can now be used in the PowerPoint password.  
    - Removed ability to use <> in password and notify end user in a dialog box if characters used.  
    This was due to an issue with command line on a PC when sending CURL commands. 
    - Tested the following special characters for use in the password: ` ~ ! @ # $ % ^ & * ( ) _ + - = { } | [ ] \ : ” ; ‘ ? , .


**Updates to WebexDeskSlidePresenter.js** (2022.07.12)
- **Fixed Bug**
  - Fixed a major bug that caused problems with using a PC virtual background when not using the PowerPoint.  Now at the end of the PowerPoint presentation, the Webex Desk will return to the last virtual background setting _except_ for the PC (HDMI or USB-C) source.  After the presentation is done, the PC can then be manually selected as a virtual background if desired. 
- **New Features**
  - **Reset screen by calling** fixscreen@example.com - Call this URI to reset the screen and virtual background settings on the Webex Desk.  The call will not connect but it triggers the reset. (This could have been a button on the touchpanel, but I was trying to avoid that). This was added in case a scenario happens that wasn't tested causes problems on the Webex Desk or if your computer freezes and the PowerPoint can't send the reset command.  Call needs to be made outside of a pre-existing calls. 
  - **Update Notification** - Added code that checks for updated versions of WebexDeskSlidePresenter.js on Github and alerts the user an update is available.  By default this is done on all macro restarts and between 3:00 am and 4:00 am in the morning.  These settings can be customized in the javascript variables: 
checkUpdates.on, checkUpdates.onStartup, checkUpdates.startTime, checkUpdates.window 
See the Javascript comments for more details. 
- **Compatible** 
  - This version of the javascript macro is comptabile with TemplateDeskProMacro_ver0.1.5.3.pptm and TemplateDeskProMacro_ver0.1.5.4.pptm

### Release Notes ver 0.1.5.3 

- **Fixed Bug**
  - Fixed an issue where at the end of a presentation or call the Webex Desk would not show the last PC source. 
- **Default Slide Show End: Nothing**
  - Added a new command **Nothing** that does nothing when the slow show ends.  This setting is experimental and not recommended, but helpful for when I create the tutorial video. 
    - Default Slide Show End Options:  StopContentShare [default], ShowPcInCall, **Nothing** [New]

### Release Notes ver 0.1.5.2 

- **New Commands:**
  - PreviousSlide
  - NoVideo
- Configure Default Settings From PowerPoint
  - Default settings can be changed from the main PowerPoint slide by scrolling to the offscreen shapes under the presentation
  - Default When No Command: 
    - Options: NoVideo [default], StopImmersiveShare, PreviousSlide, SlideNumber #
  - Default Slide Show End:  
    - Options: StopContentShare [default], ShowPcInCall
  - Changed **SideBySide_[0123]** to **SideBySide_[123x]** 
    - Because **0** (number) can be confused with **O** (letter)

### Release Notes ver 0.1.5.1

- Added cover slide from Dirk-Jan Uittenbogaard.
- Slide can be in any order (no longer needs to be the first slide). 
- Username, Password and IP Address are now stored in a slide.  
  - Password in slide 1 automatically becomes transparent at start of macro.  
  - Password transparency can be toggled in presentation by clicking on the eye icon 

- **New Commands:**
  - StopImmersiveShare
  - SpeakerTrackDiagnostic
  - SideBySide
  - SideBySide_[123x]_u
  - Prominent
  - StopContentShare


### Release Notes ver 0.1.4 
Internal only.  Changes listed above in 0.1.5  

### Release Notes ver 0.1.3 
First release on GitHub

