Attribute VB_Name = "modDM"
'These are used to create direct music
Public dmLoader As DirectMusicLoader
Public dmPerformance As DirectMusicPerformance

'These are the segments the music will be stored in
Public dmSegment As DirectMusicSegment

'This creates direct music
Sub DM_CreateLoaderPerformance(Hdl As Long)
    'Sets the loader as the directx music loader
    Set dmLoader = dxMain.DirectMusicLoaderCreate
    'Sets the performance as the directx music performance
    Set dmPerformance = dxMain.DirectMusicPerformanceCreate

    'Initializes the performance
    Call dmPerformance.Init(Nothing, Hdl)
    'Sets the active port for the performance
    Call dmPerformance.SetPort(-1, 1)
End Sub

'This loads direct music and plays the selected file
Sub DM_LoadPlayMidi(FileName As String)

    'this tells direct music to search for the files in the programs directory
    Call dmLoader.SetSearchDirectory(App.Path)
    'This creates the segment from a file
    Set dmSegment = dmLoader.LoadSegment(FileName)

    'this says that if the file ends in .mid that it is a standard midi file
    If StrConv(Right(FileName, 4), vbLowerCase) = ".mid" Then
    Call dmSegment.SetStandardMidiFile
    End If

    'this turns automatic downloading of instruments on
    Call dmPerformance.SetMasterAutoDownload(True)
    'downloads the selection for the current segment
    Call dmSegment.Download(dmPerformance)

    'this plays the segment
    Call dmPerformance.PlaySegment(dmSegment, 0, 0)
End Sub

'This unloads the midis and therefore stops them from playing
Sub DM_UnloadStopMidi()
    Set dmSegment = Nothing
    Set dmPerformance = Nothing
    Set dmLoader = Nothing
End Sub
