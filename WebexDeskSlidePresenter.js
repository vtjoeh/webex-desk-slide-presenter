/*
WebexDeskSlidePresenter.js ver 0.1.3 for the Webex Desk Pro, Desk and Desk Mini

Purpose: Use PowerPoint and Webex Desk Pro Immersive Share to superimpose your image at a predefined location in a PowerPoint slide.  
This macro requires the PowerPoint run a VB macro that talks to the endpoint.    

Author:  Joe Hughes - joehughe AT Cisco 

See GitHub site for more details and for the PowerPoint VB macro: (Linke to be determeind)
  
*/

import xapi from 'xapi';

// true or false  - Writes debug information to the console. 
const debugOn = false;

let defaultBackground = {};

defaultBackground.State = "Auto";  // On, Off or Auto -  "Auto" - selects last Background image (except USB-C or HDMI) "On" - returns to the defaults listed below, "Off" - returns to "disabled"

// Returns to this Mode if defaultBackground.State = "On".  Options: Disabled, Blur, BlurMonochrome, DepthOfField, Monochrome, Image  (Hdmi and UsbC not recommended)
defaultBackground.Mode = "Image";

// If defaultBackground.Mode = "Image", determine which image is shared.  Options: Image1, Image2, Image3, Image4, Image5, Image6, Image7, User1, User2, User3
defaultBackground.Image = "Image1";

let mainCam = 1;

let state = {};

state.presentationMode;  // "Off", "Sending", "Receiving"  - Keep track if the system is sending, receiving a presentation in a call. Off means no presentation.  

state.activeCalls; // 0, 1 or 1+ -  Keep track of the state of current activeCalls.   

state.videoMute; // "On" or "Off" - On means video mute of the main camera is turned on

state.slideImmersiveShare = "Off"; // "On" or "Off" - value is updated automatically.  

state.contentId = 2 // 2 or 3 - Default input for Content Channel. On Desk Pro  2 = USB-C, 3 = HDMI.  This is updated automatically during a call when presenting. 

state.backgroundModePc; // "UsbC" or "Hdmi"  - Above value is automatically updated when contentId is updated.  

state.hdmi = {};

state.hdmi.signal = "Unknown"; // "OK" or "NotFound".  "Unknown" until updated. 

state.hdmi.sourceId = 3; // Typically 3

state.usbc = {};

state.usbc.signal = "Unknown"; // "OK" / "NotFound". "Unknown" until updated. 

state.usbc.sourceId = 2; // typically 2 

state.lastBackground = { image: defaultBackground.Image, mode: defaultBackground.Mode };

function logFuncName(text) {
  if (debugOn) {
    console.info('function ' + text + '()');
  }
}

function consoleState() {
  if (debugOn) {
    console.info('state:', state);
  }
}

function selectDefaultBackground() {
  logFuncName("selectDefaultBackground");

  if (defaultBackground.State === "Off") {
    xapi.Command.Cameras.Background.Set({ Mode: 'Disabled' });
  }
  else {
    let nextMode, nextImage;

    if (defaultBackground.State === "On") {
      nextMode = defaultBackground.Mode;
      nextImage = defaultBackground.Image;

    } else { // defaultBackground.State === "Auto"
      nextMode = state.lastBackground.mode;
      nextImage = state.lastBackground.image;
    }
    xapi.Command.Cameras.Background.Set({ Mode: nextMode, Image: nextImage }).catch(() => {
      xapi.Command.Cameras.Background.Set({ Mode: defaultBackground.Mode });  // just incase the User Image is not set and there is an error, try again.
      console.error('error on xapi.Command.Cameras.Background.Set({ Mode: defaultBackground.Mode, Image: defaultBackground.Image }), trying again as xapi.Command.Cameras.Background.Set({ Mode: defaultBackground.Mode })');
    });
  }
}

function resetToDefault() {
  logFuncName("resetToDefault")
  // Turn off Immersive Share and go back to default
  selectDefaultBackground();
  xapi.Command.Video.Input.MainVideo.Unmute();
  state.slideImmersiveShare = "Off";
  xapi.Command.Presentation.Stop().then(() => {
    // Do any clean up if active calls is 0 or 1 
    selectDefaultBackground();
  })
}

function virtualBackground(foreground, mode = state.backgroundModePc) {
  logFuncName("virtualBackground")
  xapi.Command.Video.Input.MainVideo.Mute();
  if (state.slideImmersiveShare === 'Off') {
    // Adding some delay bfore making the sitch 
    setTimeout(() => {
      xapi.Command.Cameras.Background.Set({ Mode: mode });
      xapi.Command.Cameras.Background.ForegroundParameters.Set(foreground);
      xapi.Command.Presentation.Start({ ConnectorId: mainCam });
    }, 2050)
  }
  else {
    xapi.Command.Cameras.Background.Set({ Mode: mode });
    xapi.Command.Cameras.Background.ForegroundParameters.Set(foreground);
    xapi.Command.Presentation.Start({ ConnectorId: mainCam });
  }
  state.slideImmersiveShare = 'On';
  consoleState();
}

function virtualLocalCameraBackground(foreground, mode = state.backgroundModePc) {
  logFuncName('virtualLocalCameraBackground')
  xapi.Command.Cameras.Background.Set({ Mode: mode });
  xapi.Command.Cameras.Background.ForegroundParameters.Set(foreground);
  state.slideImmersiveShare = 'On';
}

// Convert the string to an object and send command. For example -  X:8084,Y:0,Scale:20,Opacity:100,Composition:Blend to object {"X":"8084","Y":"0","Scale":"20","Opacity":"100","Composition":"Blend"}
function parseCommand(string) {
  logFuncName("parseCommand");
  let partsArray = string.split(',');
  let keyValueArray = [];
  let locationObj = {};

  for (let i in partsArray) {
    let keyValue = partsArray[i].split(':');
    keyValueArray.push(' "' + keyValue[0] + '" : "' + keyValue[1] + '"')
  }
  locationObj = JSON.parse('{' + keyValueArray.join(',') + ' }');
  return locationObj;
}

function powerPointCommand(pptCmd) {
  logFuncName('powerPointCommand()' + JSON.stringify(pptCmd) + "  ");
  
  // Only accept a command if a presenation is already being sent while in a call
  if (pptCmd.FeedbackId === 'pptVideoSquare' && (state.presentationMode === 'Sending')) {
    if (pptCmd.Text === 'pptImmersiveShareOff' || pptCmd.Text === 'pptImmersiveSlideShowEnd') {
      xapi.command("Presentation Start", { ConnectorId: state.contentId }).then(() => {
        setTimeout(() => {
          xapi.command('Cameras Background Set', { Mode: 'Disabled' });
          xapi.command('Video Input MainVideo Unmute');
          state.slideImmersiveShare = 'Off';
          consoleState();
          selectDefaultBackground();
          if (pptCmd.Text === 'pptImmersiveSlideShowEnd') {
            screenMessage("PPT Immersive Share End", 2);
            xapi.Command.Cameras.Background.ForegroundParameters.Reset();
          }
        }, 100)
      }); // add a little delay for a smoother transition
    } else if (pptCmd.Text === 'pptImmersiveCameraOnly') {
      selectDefaultBackground();
      let location = { X: '5000', Y: '5000', Scale: '100', Opacity: '100', Composition: 'VideoPip' };
      virtualBackground(location, defaultBackground);
    }
    else {
      virtualBackground(parseCommand(pptCmd.Text));
    }
  }
  else if (pptCmd.FeedbackId === 'pptVideoSquare' && state.activeCalls == '1' && (state.presentationMode === 'Off' || state.presentationMode === 'Receiving')) {
    const regex = /X:\d+,Y:\d+,Scale:\d+,Opacity:\d+,Composition:.+/
    if (regex.test(pptCmd.Text)) {
      openPromptDisplay(pptCmd.Text);
    }
  }
  // NO CALL -  how to handle the command
  else if (pptCmd.FeedbackId === 'pptVideoSquare' && state.activeCalls == '0') {
    // This will show the PC Content only   
    if (pptCmd.Text === 'pptImmersiveShareOff') {
      let location = { X: '5000', Y: '5000', Scale: '100', Opacity: '100', Composition: 'VideoPip' }
      virtualLocalCameraBackground(location);
      // This will show the Camera only 
    } else if (pptCmd.Text === 'pptImmersiveSlideShowEnd' || pptCmd.Text === 'pptImmersiveCameraOnly') {
      selectDefaultBackground();
    }
    else {
      let theCommand = parseCommand(pptCmd.Text)
      virtualLocalCameraBackground(theCommand);
    }
  }
  else {
    // Do Nothing 
  }
}

function updateVideoMute(videoMuteOnOff) {
  logFuncName('updateVideoMute');
  if (videoMuteOnOff === "On" && state.videoMute === "Off" && state.slideImmersiveShare === "On") {
    screenMessage("PPT Immersive Share On", 9, 8200, 1300);
  } else if (videoMuteOnOff === "Off" && state.videoMute === "On") {
    xapi.Command.UserInterface.Message.TextLine.Clear();
  }
  state.videoMute = videoMuteOnOff;
  consoleState();
}

function updatePresentationMode(presentationMode) {
  logFuncName('updatePresentationMode');
  state.presentationMode = presentationMode;
  if (presentationMode === 'Receiving' || presentationMode === 'Off') {
    if (state.slideImmersiveShare === "On") {
      resetToDefault();
    }
  }
  consoleState();
}

// Determine active calls.  if active calls change from 0 to 1, reset presentation to default
function determineActiveCalls(newActiveCalls) {
  logFuncName("determineActiveCalls");
  if ((newActiveCalls == 1 && state.activeCalls == 0)) {
    resetToDefault();
  } else if (newActiveCalls == 0 && state.activeCalls == 1) {
    resetToDefault();
    xapi.Command.Presentation.Start();
  }
  state.activeCalls = newActiveCalls;
  consoleState();
}

function setBackgroundModePc(contentId) {
  logFuncName("setBackgroundModePc");
  if (contentId == 2) {
    state.backgroundModePc = 'UsbC';
  }
  else if (contentId == 3) {
    state.backgroundModePc = 'Hdmi';
  }
}

function openPromptDisplay(feedbackId) {
  logFuncName('openPromptDisplay');
  xapi.Command.UserInterface.Message.Prompt.Display({
    Title: 'Attempt to share content',
    Text: 'Would you like to share your PowerPoint now?',
    Duration: 600,
    FeedbackId: feedbackId,
    'Option.1': 'Yes',
    'Option.2': 'Cancel',
  });
}

// If a Prompt Command with FeedbackId X:<100>Y:<1000>Scale:<1000>Opacity:  and Option 1 (OK) is received, then send the last command.  
function promptCommand(promptFeedback) {
  logFuncName('promptCommand')
  const regex = /X:\d+,Y:\d+,Scale:\d+,Opacity:\d+,Composition:.+/
  if (regex.test(promptFeedback.FeedbackId)) {
    if (promptFeedback.OptionId == '1') {
      virtualBackground(parseCommand(promptFeedback.FeedbackId));
    }
  }
}

function updateDefaultPc(connectors) {
  logFuncName('updateDefaultPC');
  // connectors can be an Array of objects or an Object.  If an Object, make it an Array for the next lines of code.  
  if (!Array.isArray(connectors)) {
    connectors = [connectors];
  }
  for (const connector of connectors) {
    if ("id" in connector && connector.id == state.hdmi.sourceId) {
      state.hdmi.signal = connector.SignalState;
    }
    else if ("id" in connector && connector.id == state.usbc.sourceId) {
      state.usbc.signal = connector.SignalState;
    }
  }
  consoleState();
}

function updateEventPresentationLocalSource(localSource) {
  logFuncName('updateEventPresentationLocalSource')
  if (localSource == '2' || localSource == '3') {  // Camera is local soure 1.  
    state.contentId = localSource;
    setBackgroundModePc(state.contentId);
  }
}

function updatePresentationLocalSource(event) {
  logFuncName('updatePresentationLocalSource')
  if (!Array.isArray(event)) {
    event = [event];
  }
  if (Object.keys(event).length > 0 && "Source" in event[0]) {
    if (event[0].Source == '2' || event[0].Source == '3') {
      state.contentId = event[0].Source;
      setBackgroundModePc(state.contentId);
      consoleState();
    }
  }
}

function screenMessage(message, duration = 8, x = 10000, y = 1300) {
  xapi.Command.UserInterface.Message.TextLine.Display({
    Text: message,
    X: x,
    Y: y,
    Duration: duration,
  });
}

function updateVirtualBackgroundState(event) {
  logFuncName('updateVirtualBackgroundState');
  if ("Image" in event) {
    state.lastBackground.image = event.Image;
  }
  if ("Mode" in event) {
    if (!(event.Mode === 'UsbC' || event.Mode === "Hdmi")) {
      state.lastBackground.mode = event.Mode;
    }
  }
  consoleState();
}


setBackgroundModePc(state.contentId);

xapi.Status.Cameras.Background.on(updateVirtualBackgroundState);

xapi.Status.Cameras.Background.get().then(updateVirtualBackgroundState);

// Update the Presentation state.contentId on restart of macro
xapi.Status.Conference.Presentation.LocalInstance.get().then(updatePresentationLocalSource);

xapi.Status.Conference.Presentation.LocalInstance.on(updatePresentationLocalSource); 
// Event to update the state.contentId 
xapi.Event.PresentationStarted.LocalSource.on(updateEventPresentationLocalSource);

xapi.Status.Conference.Presentation.Mode.get().then(updatePresentationMode);

xapi.Status.Conference.Presentation.Mode.on(updatePresentationMode);

xapi.Status.Video.Input.Connector.get().then(updateDefaultPc);

xapi.Status.Video.Input.Connector.on(updateDefaultPc);

xapi.Status.Video.Input.MainVideoMute.get().then(updateVideoMute)

xapi.Status.Video.Input.MainVideoMute.on(updateVideoMute);

xapi.Status.SystemUnit.State.NumberOfActiveCalls.get().then(determineActiveCalls);

xapi.Status.SystemUnit.State.NumberOfActiveCalls.on(determineActiveCalls);

xapi.Event.UserInterface.Message.TextInput.Response.on(powerPointCommand);

xapi.Event.UserInterface.Message.Prompt.Response.on(promptCommand);

/*
Warranty & Licensing:
This is sample code.  There is no warranty for this code and no special licensing to use.  Like any custom deployment, it is the responsibility of the partner and/or customer to ensure that the customization works correctly on the device.
*/