# webex-desk-slide-presenter
Overlay your video on a PowerPoint slide using the Webex Desk.  The PowerPoint uses a macro to control the Webex Desk-series Immersive Share feature.  

## Requirement for Webex Desk Slide Presenter:

- A Webex Desk Pro, Webex Desk or Desk Mini
- Local admin access to the Webex Desk
- Add local user account on the Webex Desk with RoomControl permissions
- Install the WebexDeskSlidePresenter.js Javascript macro on the Webex Desk device
- Download the VB PowerPoint macro TemplateDeskProMacro_ver0.1.3.pptm 
- PC or Mac needs to be on the same network as the Webex Desk 

Step-by-step directions setup can be found in [Directions for WebexDeskSlidePresenter.pdf](https://github.com/vtjoeh/webex-desk-slide-presenter/blob/main/Directions%20for%20WebexDeskSlidePresenter_ver_0.1.3.pdf)

## Video Tutorial and demo

[Webex Desk Slide Presenter: Basic Setup (video 1 of 2)](https://app.vidcast.io/share/a56eda21-4818-4dab-a2ff-9448277e7783)

[Webex Desk Slide Presenter: Advanced Features (video 2 of 2)](https://app.vidcast.io/share/e5bff32f-52fd-4977-91f9-23d9bd83e803)

## Does not work with the "webex-presenter-desk-pro" macro

If you are using the "webex-presenter-desk-pro" macro that I wrote on your Webex Desk, please disable it.  The two macros use the same APIs and will probably create issues. 


# Release Notes

### ver 0.1.5.2 

### ver 0.1.5.1

- Added cover slide from Dirk-Jan Uittenbogaard.
- Slide can be in any order (no longer needs to be the first slide). 
- Username, Password and IP Address are now stored in a slide.  
  - Password in slide 1 automatically becomes transparent at start of macro.  
  - Password transparency can be toggled in presentation by clicking on the eye icon 

New Commands:
- StopImmersiveShare

- SpeakerTrackDiagnostic

- SideBySide

- SideBySide_[123x]_u

- Prominent

- StopContentShare



### ver 0.1.4 
Internal only.  Changes listed above in 0.1.4  

### ver 0.1.3 
First release
