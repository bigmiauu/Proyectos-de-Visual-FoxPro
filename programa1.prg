oForm = CREATEOBJECT("clsWebcam")
oForm.show()

DEFINE CLASS clswebcam AS form

Top = 0
Left = 0
Height = 560
Width = 647
DoCreate = .T.
Caption = "Form"
Name = "clswebcam"


ADD OBJECT shape1 AS shape WITH ;
Top = 507, ;
Left = 23, ;
Height = 34, ;
Width = 589, ;
SpecialEffect = 0, ;
Name = "Shape1"


ADD OBJECT ezvidcap1 AS olecontrol WITH ;
Top = 24, ;
Left = 123, ;
Height = 100, ;
Width = 100, ;
Name = "ezVidCap1", ;
OleClass = "vbVidC60.ezVidCap"

ADD OBJECT cmdcapture AS commandbutton WITH ;
Top = 24, ;
Left = 24, ;
Height = 27, ;
Width = 84, ;
Caption = "Capture Video", ;
Name = "cmdCapture"


ADD OBJECT cmdformatdlg AS commandbutton WITH ;
Top = 60, ;
Left = 24, ;
Height = 27, ;
Width = 84, ;
Caption = "Format DLG", ;
Name = "cmdFormatDlg"


ADD OBJECT cmddisplaydlg AS commandbutton WITH ;
Top = 96, ;
Left = 24, ;
Height = 27, ;
Width = 84, ;
Caption = "Display DLG", ;
Name = "cmdDisplayDlg"


ADD OBJECT cmdsourcedlg AS commandbutton WITH ;
Top = 132, ;
Left = 24, ;
Height = 27, ;
Width = 84, ;
Caption = "Source DLG", ;
Name = "cmdSourceDlg"


ADD OBJECT cmdcompdlg AS commandbutton WITH ;
Top = 168, ;
Left = 24, ;
Height = 27, ;
Width = 84, ;
Caption = "Comp DLG", ;
Name = "cmdCompDlg"


ADD OBJECT cmdaudiodlg AS commandbutton WITH ;
Top = 204, ;
Left = 24, ;
Height = 27, ;
Width = 84, ;
Caption = "Audio DLG", ;
Name = "cmdAudioDlg"


ADD OBJECT chkaudio AS checkbox WITH ;
Top = 305, ;
Left = 24, ;
Height = 17, ;
Width = 96, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Capture Audio", ;
Name = "chkAudio"


ADD OBJECT chkuserconfirm AS checkbox WITH ;
Top = 329, ;
Left = 24, ;
Height = 17, ;
Width = 287, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Show VFW defined user dialog to confirm capture", ;
Name = "chkUserConfirm"


ADD OBJECT chkpreroll AS checkbox WITH ;
Top = 353, ;
Left = 24, ;
Height = 17, ;
Width = 230, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Use Precise Capture Controls (Preroll)", ;
Name = "chkPreRoll"


ADD OBJECT chkpreview AS checkbox WITH ;
Top = 377, ;
Left = 24, ;
Height = 17, ;
Width = 93, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Preview video", ;
Name = "chkPreview"


ADD OBJECT chkframecallback AS checkbox WITH ;
Top = 401, ;
Left = 24, ;
Height = 17, ;
Width = 177, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Enable FrameCallback Event", ;
Name = "chkFrameCallback"


ADD OBJECT chkautosize AS checkbox WITH ;
Top = 425, ;
Left = 24, ;
Height = 17, ;
Width = 68, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Auto Size", ;
Name = "chkAutoSize"


ADD OBJECT chkcenter AS checkbox WITH ;
Top = 449, ;
Left = 24, ;
Height = 17, ;
Width = 82, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Auto Center", ;
Name = "chkCenter"


ADD OBJECT chkstretch AS checkbox WITH ;
Top = 473, ;
Left = 24, ;
Height = 17, ;
Width = 102, ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Stretch Preview", ;
Name = "chkStretch"


ADD OBJECT cmdwidth AS commandbutton WITH ;
Top = 423, ;
Left = 204, ;
Height = 27, ;
Width = 84, ;
Caption = "Set Width", ;
Name = "cmdWidth"


ADD OBJECT cmdheight AS commandbutton WITH ;
Top = 459, ;
Left = 204, ;
Height = 27, ;
Width = 84, ;
Caption = "Set Height", ;
Name = "cmdHeight"


ADD OBJECT spnwidth AS spinner WITH ;
Height = 24, ;
Left = 300, ;
Top = 423, ;
Width = 84, ;
Name = "spnWidth"


ADD OBJECT spnheight AS spinner WITH ;
Height = 24, ;
Left = 300, ;
Top = 459, ;
Width = 84, ;
Name = "spnHeight"


ADD OBJECT cmdsaveas AS commandbutton WITH ;
Top = 423, ;
Left = 420, ;
Height = 27, ;
Width = 84, ;
Caption = "Save AVI", ;
Name = "cmdSaveAs"


ADD OBJECT cmdsavedib AS commandbutton WITH ;
Top = 423, ;
Left = 520, ;
Height = 27, ;
Width = 91, ;
Caption = "Save DIB", ;
Name = "cmdSaveDIB"


ADD OBJECT cbdriver AS combobox WITH ;
Height = 24, ;
Left = 420, ;
Top = 459, ;
Width = 192, ;
Name = "cbDriver"


ADD OBJECT lblstatuscode AS label WITH ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Status Code", ;
Height = 17, ;
Left = 47, ;
Top = 516, ;
Width = 70, ;
ForeColor = RGB(255,0,0), ;
Name = "lblStatusCode"


ADD OBJECT lblstatusstring AS label WITH ;
AutoSize = .T., ;
BackStyle = 0, ;
Caption = "Status String", ;
Height = 17, ;
Left = 315, ;
Top = 516, ;
Width = 72, ;
Name = "lblStatusString"


PROCEDURE Load
_vfp.AutoYield = .f.
ENDPROC


PROCEDURE Init
LOCAL i

thisform.EnableButtons() &&check device caps and enable appropriate btns
thisform.Refresh()
If 0 < thisform.ezVidCap1.NumCapDevs
For i = 0 To thisform.ezVidCap1.NumCapDevs - 1
thisform.cbDriver.AddItem (thisform.ezVidCap1.GetDriverName(i))
endfor
thisform.cbDriver.ListIndex = thisform.ezVidCap1.DriverIndex
Else
thisform.cbDriver.AddItem ("<none>")
thisform.cbDriver.ListIndex = 0
MESSAGEBOX("No Video Capture Device!",64,thisform.Caption)
EndIf

*init form with current properties
thisform.lblStatusCode.caption = "Status Panel"
thisform.lblStatusString.caption = thisform.ezVidCap1.GetDriverVersion()
thisform.spnWidth.value = thisform.ezVidCap1.Width
thisform.spnHeight.value = thisform.ezVidCap1.Height
thisform.chkAutoSize.Value = thisform.ezVidCap1.AutoSize
thisform.chkCenter.Value = thisform.ezVidCap1.CenterVideo
thisform.chkStretch.Value = thisform.ezVidCap1.StretchPreview
thisform.chkAudio.Value = thisform.ezVidCap1.CaptureAudio
thisform.chkPreRoll.Value = thisform.ezVidCap1.UsePreciseCaptureControls
thisform.chkUserConfirm.Value = thisform.ezVidCap1.MakeUserConfirmCapture
thisform.chkPreview.Value = thisform.ezVidCap1.Preview
thisform.chkFrameCallback.Value = thisform.ezVidCap1.FrameEventEnabled
thisform.Refresh()
ENDPROC

PROCEDURE MessWithVidBits
*Still working on this sorry
ENDPROC

PROCEDURE enablebuttons
WITH thisform
.cmdAudioDlg.Enabled = .F.
.cmdFormatDlg.Enabled = .F.
.cmdDisplayDlg.Enabled = .F.
.cmdSourceDlg.Enabled = .F.
.cmdCapture.Enabled = .F.
.cmdCompDlg.Enabled = .F.
With .ezVidCap1
If .NumCapDevs > 0
thisform.cmdCapture.Enabled = .T.
thisform.cmdCompDlg.Enabled = .T.
EndIf
If .HasAudio
thisform.cmdAudioDlg.Enabled = .T.
endif
If .HasDlgFormat
thisform.cmdFormatDlg.Enabled = .T.
endif
If .HasDlgDisplay
thisform.cmdDisplayDlg.Enabled = .T.
endif
If .HasDlgSource
thisform.cmdSourceDlg.Enabled = .T.
endif
EndWith
endwith
ENDPROC


PROCEDURE ezvidcap1.StatusMessage
*** ActiveX Control Event ***
LPARAMETERS statcode, statstring
thisform.lblStatusCode.Caption = "StatusCode: " + StatCode
If StatCode <> 0
thisform.lblStatusString.Caption = StatString
EndIf
thisform.Refresh()
ENDPROC


PROCEDURE ezvidcap1.ErrorMessage
*** ActiveX Control Event ***
LPARAMETERS errcode, errstring
If ErrCode <> 0
thisform.lblStatusString.caption = "Error " + ErrString
EndIf
ENDPROC


PROCEDURE ezvidcap1.CaptureYield
*** ActiveX Control Event ***
DOEVENTS
ENDPROC


PROCEDURE ezvidcap1.FrameCallback
*** ActiveX Control Event ***
LPARAMETERS lpvhdr
=thisform.MessWithVidBits(lpVHdr)
ENDPROC


PROCEDURE ezvidcap1.PreRollComplete
*** ActiveX Control Event ***
Local userRet

userRet = MessageBox("Using precise capture controls." +CHR(13) + CHR(10) +;
"PreRoll complete - Click OK to start capture immediately.",1, thisform.caption)
If userRet = 1
thisform.ezVidCap1.PreciseCaptureStart
Else
thisform.ezVidCap1.PreciseCaptureCancel
EndIf
ENDPROC


PROCEDURE cmdcapture.Click
thisform.ezVidCap1.CaptureVideo()
ENDPROC


PROCEDURE cmdformatdlg.Click
thisform.ezVidCap1.ShowDlgVideoFormat()
ENDPROC


PROCEDURE cmddisplaydlg.Click
thisform.ezVidCap1.ShowDlgVideoDisplay()
ENDPROC


PROCEDURE cmdsourcedlg.Click
thisform.ezVidCap1.ShowDlgVideoSource()
ENDPROC


PROCEDURE cmdcompdlg.Click
thisform.ezVidCap1.ShowDlgCompressionOptions()
ENDPROC


PROCEDURE cmdaudiodlg.Click
thisform.ezVidCap1.ShowDlgAudioFormat()
ENDPROC


PROCEDURE chkaudio.InteractiveChange
If thisform.chkAudio.Value = .T.
thisform.ezVidCap1.CaptureAudio = .T.
Else
thisform.ezVidCap1.CaptureAudio = .F.
EndIf
ENDPROC


PROCEDURE chkuserconfirm.InteractiveChange
If thisform.chkUserConfirm.Value = .t.
thisform.ezVidCap1.MakeUserConfirmCapture = .T.
Else
thisform.ezVidCap1.MakeUserConfirmCapture = .F.
EndIf
ENDPROC


PROCEDURE chkpreroll.InteractiveChange
If thisform.chkPreRoll.Value = .t.
thisform.ezVidCap1.UsePreciseCaptureControls = .T.
Else
thisform.ezVidCap1.UsePreciseCaptureControls = .F.
EndIf
ENDPROC


PROCEDURE chkpreview.InteractiveChange
If thisform.chkPreview.Value = .t.
thisform.ezVidCap1.Preview = .T.
Else
thisform.ezVidCap1.Preview = .F.
EndIf
ENDPROC


PROCEDURE chkframecallback.InteractiveChange
If thisform.chkFrameCallback.Value = .t.
thisform.ezVidCap1.FrameEventEnabled = .T.
Else
thisform.ezVidCap1.FrameEventEnabled = .F.
EndIf
ENDPROC


PROCEDURE chkautosize.InteractiveChange
If thisform.chkAutoSize.Value = .t.
thisform.ezVidCap1.AutoSize = .T.
Else
thisform.ezVidCap1.AutoSize = .F.
EndIf
ENDPROC


PROCEDURE chkcenter.InteractiveChange
If thisform.chkCenter.Value = .t.
thisform.ezVidCap1.CenterVideo = .T.
Else
thisform.ezVidCap1.CenterVideo = .F.
EndIf
ENDPROC


PROCEDURE chkstretch.InteractiveChange
If thisform.chkStretch.Value = .t.
thisform.ezVidCap1.StretchPreview = .T.
Else
thisform.ezVidCap1.StretchPreview = .F.
EndIf
ENDPROC


PROCEDURE cmdwidth.Click
thisform.ezVidCap1.Width = thisform.spnWidth.value
*show actual size (in case auto size is turned on)
thisform.spnWidth.value = thisform.ezVidCap1.Width
ENDPROC


PROCEDURE cmdheight.Click
thisform.ezVidCap1.Height = thisform.spnHeight.value
*show actual size (in case auto size is turned on)
thisform.spnHeight.value = thisform.ezVidCap1.Height
ENDPROC


PROCEDURE cmdsaveas.Click
LOCAL cFile
cFile = PUTFILE('AVI files', 'capture.avi', 'AVI')
IF !EMPTY(cFile)
thisform.ezVidCap1.SaveAs(cFile)
ENDIF
ENDPROC


PROCEDURE cmdsavedib.Click
LOCAL cFile
cFile = PUTFILE('Bitmap files', 'capture.bmp', 'BMP')
IF !EMPTY(cFile)
thisform.ezVidCap1.SaveDIB(cFile)
ENDIF
ENDPROC


PROCEDURE cbdriver.InteractiveChange
Local oldDriver
oldDriver = thisform.ezVidCap1.DriverIndex

On ERROR *

thisform.ezVidCap1.DriverIndex = thisform.cbDriver.ListIndex

If ERROR()
*restore old settings
thisform.ezVidCap1.DriverIndex = oldDriver
thisform.cbDriver.ListIndex = oldDriver
thisform.lblStatusString.caption = "Could not connect!"
EndIf
ENDPROC


ENDDEFINE