/*
WebexDeskSlidePresenter.js ver 0.1.5.5 for the Webex Desk Pro, Desk and Desk Mini

Purpose: Use PowerPoint and Webex Desk Immersive Share to superimpose your image at a predefined location in a PowerPoint slide.  
This macro requires the PowerPoint run a VB macro that communicates with the endpoint.    

Author:  Joe Hughes
let contact =  'joehughe' + '@' + 'cisco' + '.com'
Signficant contributions by: Dirk-Jan Uittenbogaard

See GitHub site for more details and for the PowerPoint VBA macro: https://github.com/vtjoeh/webex-desk-slide-presenter
No warranty implied or otherwise.  
  
*/

import xapi from 'xapi';

const version = '0.1.5.5';

const debugOn = false;  // true or false.  Default: false - Writes debug information to the console. 

const checkUpdates = {};

checkUpdates.on = true; // true or false.  Default: true.  Message is placed on screen that the WebexDeskSlidePresenter is updated on Github.  

checkUpdates.onStartup = true; // true or false. Default: true.  Checks the status on the startup of the script.  Nothing is shown if the system is in a call.  

checkUpdates.startTime = '03:00'; // hh:mm in military time. Default: '03:00' (e.g. 3:00 am).  Start of time to check for updates. Local time of hte Webex desk. 

checkUpdates.window = 60; // minutes.  Default: 60 min.  The window to check the times. checkUpdates.startTime + checkUpdates.window.  Default updates are checked between 03:00 am and 04:00 am.  

checkUpdates.timerInterval = 30; // minutes. Default: 30 min. setTimeout value to check if updates need to be checked.   This number should be less than the checkUpdates.timerInterval.  

checkUpdates.lastCheckTime = 60 * 25; // last time the update was checked listed as Minute of the Day.   Default is 25 hours.  Value is reset when checked. 

const fixScreenUri = 'fixscreen@example.com'  // URI to call to reset the screen settings of the device.  Only immersive share and camera settings are reset.  

let vidcastSelfview = {};

vidcastSelfview.mode = 'On';  // 'On' or 'Off'. Default: 'On'.   If 'On', selfview is only showed for non-video call uses like vidcast.  

vidcastSelfview.fullScreenMode = 'On'; // 'On' or 'Off'. Default: 'On'.   If 'Off' and selfview.mode = 'On', then selfview.pipPosition is used 

vidcastSelfview.pipPosition = 'LowerRight';  // Only used if selfview.fullScreenMode = 'Off'.  selfview.pipPosition options: CenterLeft, CenterRight, Current, LowerLeft, LowerRight, UpperCenter, UpperLeft, UpperRight

let defaultBackground = {};

defaultBackground.State = 'Auto';  // 'On', 'Off' or 'Auto'. Default: 'Auto'.  'Auto' tries to select last background image (except USB-C or HDMI). 'On' - returns to the defaults listed below, 'Off' - returns to "disabled"

defaultBackground.Mode = 'Image';  // Returns to this Mode if defaultBackground.State = 'On'.  Options: Disabled, Blur, BlurMonochrome, DepthOfField, Monochrome, Image  (Hdmi and UsbC not recommended)

defaultBackground.Image = 'Image1';  // If defaultBackground.Mode = 'Image', determine which image is shared.  Options: Image1, Image2, Image3, Image4, Image5, Image6, Image7, User1, User2, User3

let mainCam = 1;  // typically 1 

let state = {};  // state object keeps track of various states of the Webex Desk

state.presentationMode;  // 'Off', 'Sending', 'Receiving'  - Keep track if the system is sending, receiving a presentation in a call. Off means no presentation.  

state.activeCalls; // 0, 1 or 1+ -  Keep track of the number current activeCalls.   

state.videoMute; // 'On' or 'Off' - Keep track if video mute is on or off. 'On' means video mute of the main camera is turned on (but the main camera might be seen in the content channel)

state.slideImmersiveShare = 'Off'; // 'On' or 'Off' - Keep track if Immersive Share started by this macro is on or off.  Value is updated automatically. Equal, Prominent and SpeakerTrackDiag all count as "immersive share"

state.contentId = 2 // 2 or 3 - Default input for Content Channel. On Desk Pro  2 = USB-C, 3 = HDMI.  Updated to whatever source was shown last. 

state.backgroundModePc; // 'UsbC' or 'Hdmi'  - Automatically updated when state.contentId is updated to whatever source was shown last.  

state.speakerTrackDiag = 'notset'; // 'On', 'Off' or 'notset'  used to store the last command to turn SpearkerTrack diagnostic mode on or off. 

state.cameraOnlyAsContent = 'notset' // 'On', 'Off' or 'notset' used to store the last command to turn the Camera Only when sent as content. 

state.hdmi = {};

state.hdmi.signal = 'Unknown'; // 'OK' or 'NotFound'.  'Unknown' until updated. 

state.hdmi.sourceId = 3; // Typically 3.  Value should not change. 

state.usbc = {};

state.usbc.signal = 'Unknown'; // 'OK' / 'NotFound'. 'Unknown' until updated. 

state.usbc.sourceId = 2; // typically 2.  Value should not change

state.lastCommand // last command received from the PowerPoint

state.lastFeedbackId;

state.allowMuteMainVideo = true; // 'true' or 'false'.  Should start as 'true'.  PowerPoint can override this setting. Corresponds to 'Main Video Stream: Mirror/Mute' in PPT user settings

state.lastBackground = { image: defaultBackground.Image, mode: defaultBackground.Mode };

function logFuncName(functionName, optionalText = "") {
  if (debugOn) {
    console.info('function ' + functionName + '() ' + optionalText);
  }
}

function consoleState() {
  if (debugOn) {
    console.info('state:', state);
  }
}

function selectDefaultBackground() {
  logFuncName("selectDefaultBackground");
  turnSpeakerTrackDiagOff();
  if (defaultBackground.State === "Off") {
    xapi.Command.Cameras.Background.Set({ Mode: 'Disabled' });
  }
  else {
    let nextMode, nextImage;

    if (defaultBackground.State === "On") {
      nextMode = defaultBackground.Mode;
      nextImage = defaultBackground.Image;
    } else { // defaultBackground.State === "Auto"
      nextMode = state.lastBackground.mode || defaultBackground.Mode;
      nextImage = state.lastBackground.image || defaultBackground.Mode;
    }

    xapi.Command.Cameras.Background.Set({ Mode: nextMode, Image: nextImage }).catch(() => {
      xapi.Command.Cameras.Background.Set({ Mode: defaultBackground.Mode });  // just incase the User Image is not set and there is an error, try again.
      console.error('error on xapi.Command.Cameras.Background.Set({ Mode: defaultBackground.Mode, Image: defaultBackground.Image }), trying again as xapi.Command.Cameras.Background.Set({ Mode: defaultBackground.Mode })');
    });
  }
  consoleState();
}

function turnSpeakerTrackDiagOff() {
  logFuncName("turnSpeakerTrackDiagOff", "state.speakerTrackDiag " + state.speakerTrackDiag)
  if (state.speakerTrackDiag !== 'Off') {
    state.speakerTrackDiag = 'Off';
    xapi.Command.Cameras.SpeakerTrack.Diagnostics.Stop();
  }
  consoleState();
}

function resetToDefault() {
  // Turn off Immersive Share and go back to default
  logFuncName('resetToDefault', 'state.slideImmersiveShare: ' + state.slideImmersiveShare)
  if (state.slideImmersiveShare === 'On' || state.speakerTrackDiag === 'On' || state.cameraOnlyAsContent === 'On') {
    state.slideImmersiveShare = 'Off';
    state.cameraOnlyAsContent = 'Off';
    xapi.Status.SystemUnit.State.NumberOfActiveCalls.get().then(activeCalls => {
      logFuncName('resetToDefault()->xapi.Status.SystemUnit.State.NumberOfActiveCalls.get(): ' + activeCalls);
      if (activeCalls == 0) {
        setTimeout(() => {
          xapi.Command.Presentation.Start({ ConnectorId: state.contentId });
          xapi.Command.Video.Selfview.Set({ FullscreenMode: 'Off', Mode: 'Off' });
          xapi.Command.Video.Input.MainVideo.Unmute();
          xapi.Command.Cameras.Background.ForegroundParameters.Reset();
          selectDefaultBackground();
        }, 250)  // add some delay to keep fullscreen selfview from flashing on screen
      } else {
        xapi.Command.Video.Selfview.Set({ FullscreenMode: 'Off' });
        xapi.Command.Video.Input.MainVideo.Unmute();
        xapi.Command.Presentation.Stop().then(() => {
          // Do any clean up 
          xapi.Command.Cameras.Background.ForegroundParameters.Reset();
          xapi.Command.Video.Input.MainVideo.Unmute();
          selectDefaultBackground();
        })
      }
    })
  }
  consoleState();
}

function virtualBackground(foreground, mode = state.backgroundModePc) {
  logFuncName("virtualBackground")
  turnSpeakerTrackDiagOff();
  if (state.allowMuteMainVideo) {xapi.Command.Video.Input.MainVideo.Mute() };
  if (state.slideImmersiveShare === 'Off') {
    // Adding some delay bfore making the switch 
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
  turnSpeakerTrackDiagOff();
  xapi.Command.Cameras.Background.Set({ Mode: mode });
  xapi.Command.Cameras.Background.ForegroundParameters.Set(foreground);
  state.slideImmersiveShare = 'On';
  consoleState();
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

function replaceElementinArray(arr, findItem, replaceItem) {
  arr.forEach((element, index) => {
    if (element === findItem) {
      arr[index] = replaceItem;
    }
  });
  return arr;
}

function presentationEqual(text) {
  logFuncName("presentationEqual() text: " + text)
  let presentationSource = [state.contentId, mainCam]; // set default values
  let mute = true;
  selectDefaultBackground();
  text = text.toLowerCase();
  const regex = /pptimmersive?equal_?([0123]{1,3})?_?([mu])?/;
  const match = text.match(regex);

  state.slideImmersiveShare = 'On';

  if (match[1] !== undefined) {
    presentationSource = match[1].split('');
    presentationSource = replaceElementinArray(presentationSource, '0', state.contentId); // replace 0 with the state.contentId so the last active HDMI/USBC signal is shared
  }

  if (match[2] === 'u') {
    mute = false
  }

  if (mute === true) {
    xapi.Command.Video.Input.MainVideo.Mute();
  } else {
    xapi.Command.Video.Input.MainVideo.Unmute();
  }

  xapi.Command.Presentation.Start({ PresentationSource: presentationSource, Layout: 'Equal' });
  consoleState();

}

function powerPointCommand(pptCmd) {
  logFuncName('powerPointCommand()' + JSON.stringify(pptCmd) + "  ");
 
  state.lastCommand = pptCmd.Text
  state.lastFeedbackId = pptCmd.FeedbackId; 

  if (pptCmd.FeedbackId === 'pptVideoSquareDual'){  // Send as Dual streams for dual stream devices 
    pptCmd.FeedbackId = 'pptVideoSquare'; 
    state.allowMuteMainVideo = false; 
  } 
  else if (pptCmd.FeedbackId === 'pptVideoSquare'){
    state.allowMuteMainVideo = true; 
  }

  // Only accept a command if a presenation is already being sent while in a call
  if ((pptCmd.FeedbackId === 'pptVideoSquare' && state.presentationMode === 'Sending') || pptCmd.FeedbackId === 'pptVideoSquare2') {
    if (pptCmd.Text === 'pptImmersiveShareOff' || pptCmd.Text === 'pptImmersiveSlideShowEnd') {
      xapi.command("Presentation Start", { ConnectorId: state.contentId }).then(() => {
        setTimeout(() => {
          turnSpeakerTrackDiagOff();
          xapi.command('Cameras Background Set', { Mode: 'Disabled' });
          xapi.command('Video Input MainVideo Unmute');
          state.slideImmersiveShare = 'Off';
          selectDefaultBackground();

          if (pptCmd.Text === 'pptImmersiveSlideShowEnd') {
            screenMessage("PPT Immersive Share End", 2);
          }
        }, 100)
      }); // add a little delay for a smoother transition
      consoleState();
    }
    else if (pptCmd.Text === 'pptImmersiveCameraOnly') {
      selectDefaultBackground();
      let location = { X: '5000', Y: '5000', Scale: '100', Opacity: '100', Composition: 'VideoPip' };
      virtualBackground(location, "Image");
      state.cameraOnlyAsContent = 'On';
    }
    else if (pptCmd.Text === 'pptImmersiveStopContentShare') {
      resetToDefault();
    }
    else if (pptCmd.Text.includes('pptImmersiveEqual')) {
      presentationEqual(pptCmd.Text)
    }
    else if (pptCmd.Text === 'pptImmersiveProminent') {
      selectDefaultBackground();
      xapi.Command.Video.Input.MainVideo.Mute();
      xapi.Command.Presentation.Start({ PresentationSource: [state.contentId, mainCam], Layout: 'Prominent' });
      state.slideImmersiveShare = 'On';
      consoleState();
    }
    else if (pptCmd.Text === 'pptImmersiveSpeakerTrackDiag') {
      state.speakerTrackDiag = 'On';
      state.slideImmersiveShare = 'On';
      xapi.Command.Cameras.SpeakerTrack.Diagnostics.Start();
      if (state.allowMuteMainVideo) {xapi.Command.Video.Input.MainVideo.Mute() };
      xapi.Command.Presentation.Start({ ConnectorId: mainCam });
      consoleState();
    }
    else if (pptCmd.Text === 'pptImmersiveSelfviewToggle') {
      noSelfviewToggle();
    }
    else if (pptCmd.Text === 'pptImmersiveNoVideo') {
      noVideo();
    }
    else {
      virtualBackground(parseCommand(pptCmd.Text));
    }
  }
  else if ((pptCmd.FeedbackId === 'pptVideoSquare' || pptCmd.FeedbackId === 'pptVideoSquareDual') && state.activeCalls == '1' && (state.presentationMode === 'Off' || state.presentationMode === 'Receiving')) {
    const regex = /(X:\d+,Y:\d+,Scale:\d+,Opacity:\d+,Composition:.+)|pptImmersive(Speak|Prom|3vid|Equa|Came).*/
    if (regex.test(pptCmd.Text)) {
      openPromptDisplay(pptCmd.Text);
    }
  }

  // NO CALL -  how to handle the commands in situations like vidcast
  else if (pptCmd.FeedbackId === 'pptVideoSquare' && state.activeCalls == '0') {  // This will show the Camera only 
    // This will show the Camera only 
    if (pptCmd.Text.includes('pptImmersiveEqual') || pptCmd.Text === 'pptImmersiveProminent' || pptCmd.Text === 'pptImmersiveNoVideo' || pptCmd.Text === 'pptImmersiveShareOff') {
      let location = { X: '5000', Y: '5000', Scale: '100', Opacity: '100', Composition: 'VideoPip' }
      turnSpeakerTrackDiagOff();
      virtualLocalCameraBackground(location);
      selfviewOff();
    } else if (pptCmd.Text === 'pptImmersiveSlideShowEnd' || pptCmd.Text === 'pptImmersiveStopContentShare') {  // This will show the Camera only 
      selfviewOff();
      resetToDefault();
    }
    else if (pptCmd.Text === 'pptImmersiveCameraOnly') {
      selectDefaultBackground();
      if (vidcastSelfview.mode === 'On') {
        selfviewOn(vidcastSelfview)
      };
    }
    else if (pptCmd.Text === 'pptImmersiveSpeakerTrackDiag') {
      state.speakerTrackDiag = 'On';
      xapi.Command.Cameras.SpeakerTrack.Diagnostics.Start();
      if (vidcastSelfview.mode === 'On') {
        selfviewOn(vidcastSelfview)
      };
      consoleState();
    }
    else if (pptCmd.Text.startsWith('pptImmersiveSelfviewToggle')) {
      toggleSelfView();
    }
    else {
      let theCommand = parseCommand(pptCmd.Text)
      virtualLocalCameraBackground(theCommand);
      if (vidcastSelfview.mode === 'On') {
        selfviewOn(vidcastSelfview)
      };
    }
  }
}

function noVideo() {
  logFuncName('noVideo()');
  presentationEqual('pptimmersiveequal_' + state.contentId)
}

function noSelfviewToggle() {
  logFuncName("noSelfviewToggle()")
  screenMessage("No selfview in call with immersive share on.", 5, 8200, 1300)
}

function toggleSelfView() {
  logFuncName("toggleSelfView()")
  xapi.Status.Video.Selfview.Mode.get().then(selfview => {
    if (selfview === 'On') {
      vidcastSelfview.mode = 'Off'
      selfviewOff();
    } else {
      vidcastSelfview.mode = 'On'
      selfviewOn(vidcastSelfview);
    }
  })
  consoleState();
}

function selfviewOn(sv) {
  logFuncName('selfviewOn' + JSON.stringify(sv));
  xapi.Command.Video.Selfview.Set(
    {
      FullscreenMode: sv.fullScreenMode,
      Mode: sv.mode,
      PIPPosition: sv.pipPosition
    });
  consoleState();
}

function selfviewOff() {
  xapi.Command.Video.Selfview.Set(
    { Mode: 'Off' });
  consoleState();
}

function updateVideoMuteState(videoMuteOnOff) {
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
  logFuncName("determineActiveCalls() newActiveCalls: " + newActiveCalls + " state.activeCalls: " + state.activeCalls);
  if ((newActiveCalls == 1 && state.activeCalls == 0)) {  // Call is connecting
    // do something when call connects
    resetToDefault();
  } else if (newActiveCalls == 0 && state.activeCalls == 1) {  // Call is disconnecting
    // do something when call disconnects
    resetToDefault();
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

function doCheckUpdate() {
  if (checkUpdates.on === true && state.activeCalls === '0' && state.slideImmersiveShare === 'Off') {
    checkVersionUpdate();
  }
}

function turnOnHttpClient(delay = 7000) {
  if (checkUpdates.on) {
    xapi.Config.HttpClient.Mode.set('On');
    if (checkUpdates.onStartup) {
      setTimeout(doCheckUpdate, delay);
    }
  }
}

function checkVersionUpdate() {
  let url = "https://raw.githubusercontent.com/vtjoeh/webex-desk-slide-presenter/main/version.txt";
  let headers = 'Content-Type: application/json'
  xapi.Command.HttpClient.Get({
    "Url": url,
    "Header": headers
  }).then((response) => {
    let responseJSON = response.Body.replace(/\n/g, '').replace(/\'/g, '"');
    responseJSON = JSON.parse(responseJSON);

    if (responseJSON['WebexDeskSlidePresenter.js']) {
      let newVersion = responseJSON['WebexDeskSlidePresenter.js'];
      if (compareReleaseNumber(newVersion, version)) {
        notifyUserOfUpdate(newVersion);
      } else {
        console.info('WebexDeskSlidePresenter.js is up to date. Current version: ' + version + '. Github version: ' + newVersion);
      };
    }
  }).catch(error => {
    console.info('Could not reach ' + url);
    checkUpdates.lastCheckTime = 60 * 25; // allow for retrying to connect if it is in the time window.   
    if (error.message) {
      console.info('error.message', error.message);
    }
    if (error.data && error.data.StatusCode) {
      console.info('error.data.StatusCode', error.data.StatusCode);
    }

  });
}

// return true if the newVersion is newer than the oldVersion.  Needs to be in the format [number].[number].[number] for example 0.1.003 or 2.3.4
function compareReleaseNumber(newVersion, oldVersion = version) {
  newVersion = newVersion.split('.');
  oldVersion = oldVersion.split('.');

  for (let i = 0; i < newVersion.length; i++) {
    let oldVersionElement = oldVersion[i] || '0';
    let newVersionElement = newVersion[i];

    if (i === 0) {  // if the first element has a leading zero, strip it.  For example 01 becomes 1.  Other elements 01 is sm
      oldVersionElement = parseInt(oldVersionElement, 10);
      newVersionElement = parseInt(newVersionElement, 10);
    }

    if (newVersionElement > oldVersionElement) {
      return true;
    }
    else if (newVersion[i] < oldVersionElement) {
      return false;
    }
  }
  return false;
}

function notifyUserOfUpdate(newVersion) {
  logFuncName('openPromptDisplay');
  let text = 'Current version: ' + version + ', New version: ' + newVersion + ', Available at:<br> https://github.com/vtjoeh/webex-desk-slide-presenter';
  console.info('WebexDeskSlidePresenter.js version check.', text);
  if (text.length > 254) {
    text = ' New version available at:<br/> https://github.com/vtjoeh/webex-desk-slide-presenter';
  }
  xapi.Command.UserInterface.Message.Prompt.Display({
    Title: 'WebexDeskSlidePresenter.js       macro update available.',
    Text: text,
    Duration: 0,
    FeedbackId: 'Notify user of update for WebexDeskSlidePresenter.js cleared',
    'Option.1': 'Ok',
  });
}

function turnOnTimerUpdateCheck() {
  let arrayCheckTimeStart = checkUpdates.startTime.split(':');
  let checkMinuteOfDayStart = parseInt(arrayCheckTimeStart[0], 10) * 60 + parseInt(arrayCheckTimeStart[1], 10);
  let checkMinuteOfDayEnd = checkMinuteOfDayStart + checkUpdates.window;
  let currentMinuteOfDay = (new Date).getHours() * 60 + (new Date).getMinutes();
  // console.log('checkMinuteOfDayStart', checkMinuteOfDayStart, 'currentMinuteOfDay', currentMinuteOfDay, "checkMinuteOfDayEnd", checkMinuteOfDayEnd, 'checkUpdates.lastCheckTime', checkUpdates.lastCheckTime, 'checkUpdates.window', checkUpdates.window);

  if (checkMinuteOfDayStart < currentMinuteOfDay && currentMinuteOfDay < checkMinuteOfDayEnd && currentMinuteOfDay < checkUpdates.lastCheckTime) {
    checkUpdates.lastCheckTime = currentMinuteOfDay;
    doCheckUpdate();
  }

  setTimeout(turnOnTimerUpdateCheck, checkUpdates.timerInterval * 60 * 1000);
}

// If a Prompt Command with FeedbackId X:<100>Y:<1000>Scale:<1000>Opacity:  and Option 1 (OK) is received, then send the last command.  
function promptCommand(promptFeedback) {
  logFuncName('promptCommand')
  const regex = /X:\d+,Y:\d+,Scale:\d+,Opacity:\d+,Composition:.+/
  const regex2 = /pptImmersive.+/
  if (regex.test(promptFeedback.FeedbackId)) {
    if (promptFeedback.OptionId == '1') {
      virtualBackground(parseCommand(promptFeedback.FeedbackId));
    }
  }
  else if (regex2.test(promptFeedback.FeedbackId)) {
    if (promptFeedback.OptionId == '1') {
      xapi.Command.UserInterface.Message.TextInput.Response(
        {
          FeedbackId: 'pptVideoSquare2',
          Text: promptFeedback.FeedbackId
        }
      );
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
  console.log("updatePresentationLocalSource: event", event); 
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
  logFuncName('screenMessage() message: ' + message)
  xapi.Command.UserInterface.Message.TextLine.Display({
    Text: message,
    X: x,
    Y: y,
    Duration: duration,
  });
}

function resetScreen(uri) {
  logFuncName("resetScreen", "uri: " + uri + " fixScreenUri: " + fixScreenUri);
  if (uri.toLowerCase() === fixScreenUri.toLowerCase()) {
    xapi.Command.Cameras.SpeakerTrack.Diagnostics.Stop();
    xapi.Command.Presentation.Start({ ConnectorId: state.contentId });
    xapi.Command.Video.Selfview.Set({ FullscreenMode: 'Off', Mode: 'Off' });
    xapi.Command.Video.Input.MainVideo.Unmute();
    xapi.Command.Cameras.Background.ForegroundParameters.Reset();
    xapi.Command.Cameras.Background.Set({ Mode: 'Disabled', Image: 'Image1' });
    state.speakerTrackDiag = 'Off';
    state.slideImmersiveShare = 'Off';
    state.cameraOnlyAsContent = 'Off';
    screenMessage('Screen settings reset to default.  Virtual background set to off.', 12);
  }

}

function updateVirtualBackgroundState(event) {
  logFuncName('updateVirtualBackgroundState() event' + JSON.stringify(event));
  if ("Image" in event) {
    state.lastBackground.image = event.Image;
  }
  if ("Mode" in event) {
    if (!(event.Mode === 'UsbC' || event.Mode === "Hdmi") && state.speakerTrackDiag !== 'On') {
      state.lastBackground.mode = event.Mode;
    }
  }
  consoleState();
}

// Takes input from a standard PowerPoint clicker. 
//  Web Interface:  Settings --> Configuration --> Peripherals --> InputDevice Mode: On

let keyState = {};

keyState.KEY_LEFTALT = 'Released';

function usbKeyInput(keyAction) {

  let pptCmd = {};
  if(state.lastFeedbackId){
    pptCmd.FeedbackId = state.lastFeedbackId; 
  }
  else {
    pptCmd.FeedbackId = 'pptVideoSquare';
  }
  
  if (keyAction.Type === 'Released' && keyAction.Key === 'KEY_LEFTALT') {
    keyState.KEY_LEFTALT = 'Released';
  }

  if (keyAction.Type === 'Pressed') {
    if (keyAction.Key == 'KEY_RIGHT' || keyAction.Key == 'KEY_DOWN' || keyAction.Key == 'KEY_PAGEDOWN') {
      pptCmd.Text = 'X:8000,Y:8000,Scale:40,Opacity:100,Composition:Blend';
      powerPointCommand(pptCmd);
    }
    else if (keyAction.Key == 'KEY_LEFT' || keyAction.Key == 'KEY_UP' || keyAction.Key == 'KEY_PAGEUP') {
      pptCmd.Text = 'X:2000,Y:8000,Scale:40,Opacity:100,Composition:Blend';
      powerPointCommand(pptCmd);
    }
    else if (keyAction.Key == 'KEY_VOLUMEUP') {
      pptCmd.Text = 'X:2000,Y:2000,Scale:40,Opacity:100,Composition:Blend';
      powerPointCommand(pptCmd);
    }
    else if (keyAction.Key == 'KEY_VOLUMEDOWN') {
      pptCmd.Text = 'X:8000,Y:2000,Scale:40,Opacity:100,Composition:Blend';
      powerPointCommand(pptCmd);
    }
    else if (keyAction.Key == 'KEY_ENTER' || keyAction.Key == 'KEY_B') {
      pptCmd.Text = 'pptImmersiveSlideShowEnd';
      powerPointCommand(pptCmd);
    }
    else if (keyAction.Key == 'KEY_TAB' && keyState.KEY_LEFTALT == 'Released') {  
      pptCmd.Text = 'X:200,Y:200,Scale:1,Opacity:1,Composition:Blend';
      powerPointCommand(pptCmd);
    }
    else if (keyAction.Key == 'KEY_TAB' && keyState.KEY_LEFTALT == 'Pressed') {
      state.slideImmersiveShare = 'On';  // temporarily change state to state.slideImmersiveShare = 'On' even though it may not be true. 
      pptCmd.Text = 'pptImmersiveStopContentShare';
      powerPointCommand(pptCmd);
    }
    else if (keyAction.Key == 'KEY_ESC' || keyAction.Key == 'KEY_S') {
      // 
    }
    else if (keyAction.Key == 'KEY_LEFTALT') {
      keyState.KEY_LEFTALT = 'Pressed';
    }
  }
}



setTimeout(turnOnTimerUpdateCheck, 7000);

turnOnHttpClient();

setBackgroundModePc(state.contentId);

xapi.Status.Cameras.Background.on(updateVirtualBackgroundState);

xapi.Status.Cameras.Background.get().then(updateVirtualBackgroundState);

xapi.Status.Conference.Presentation.LocalInstance.get().then(updatePresentationLocalSource); // Update the Presentation state.contentId on restart of macro

xapi.Status.Conference.Presentation.LocalInstance.on(updatePresentationLocalSource); // Event to update the state.contentId 

xapi.Event.PresentationStarted.LocalSource.on(updateEventPresentationLocalSource);

xapi.Status.Conference.Presentation.Mode.get().then(updatePresentationMode);

xapi.Status.Conference.Presentation.Mode.on(updatePresentationMode);

xapi.Status.Video.Input.Connector.get().then(updateDefaultPc);

xapi.Status.Video.Input.Connector.on(updateDefaultPc);

xapi.Status.Video.Input.MainVideoMute.get().then(updateVideoMuteState)

xapi.Status.Video.Input.MainVideoMute.on(updateVideoMuteState);

xapi.Status.SystemUnit.State.NumberOfActiveCalls.get().then(determineActiveCalls);

xapi.Status.SystemUnit.State.NumberOfActiveCalls.on(determineActiveCalls);

xapi.Event.UserInterface.Message.TextInput.Response.on(powerPointCommand);

xapi.Event.UserInterface.Message.Prompt.Response.on(promptCommand);

xapi.Event.CallDisconnect.DisplayName.on(resetScreen);

xapi.Event.UserInterface.InputDevice.Key.Action.on(usbKeyInput); 
