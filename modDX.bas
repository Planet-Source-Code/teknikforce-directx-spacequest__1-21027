Attribute VB_Name = "modDX"
'The Main DirectX Object

'This sets the cooperative levels of Direct Draw and Direct Sound
Sub DX_SetCoopLevel(Hdl As Long)
    Call dsMain.SetCooperativeLevel(Hdl, DSSCL_PRIORITY)
    Call ddMain.SetCooperativeLevel(Hdl, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
End Sub

'This sets the screens display mode
Sub DX_SetDisplay(sWidth As Long, sHeight As Long, sBPP As Long)
    Call ddMain.SetDisplayMode(sWidth, sHeight, sBPP, 0, DDSDM_DEFAULT)
End Sub

'This is the main DirectX Initialization
Sub DX_Init()

    'If already in directx mode, why go in it again
    If InDirectXMode Then Exit Sub

    On Error GoTo errInit

    'Creates Direct Draw and Direct Sound
    Set ddMain = dxMain.DirectDrawCreate("")
    Set dsMain = dxMain.DirectSoundCreate("")
    'Sets Their Cooperative levels
    Call DX_SetCoopLevel(frmMain.hWnd)
    'Sets the screen's display to 640x480x16
    Call DX_SetDisplay(640, 480, 16)

    'Puts computer in directx mode
    InDirectXMode = True
    'Hides the cursor
    Call ShowCursor(0)
    Exit Sub

errInit:
    MsgBox "DirectX couldn't be initialized! Please check to see if it has been installed correctly!", vbExclamation, "Error!"
End Sub

'Restores Cooperative levels of Direct Draw and Direct Sound at programs end
Sub DX_RestoreCoopLevel(Hdl As Long)
    Call dsMain.SetCooperativeLevel(Hdl, DSSCL_NORMAL)
    Call ddMain.SetCooperativeLevel(Hdl, DDSCL_NORMAL)
End Sub
