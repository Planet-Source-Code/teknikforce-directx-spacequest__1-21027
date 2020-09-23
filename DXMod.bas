Attribute VB_Name = "DXMod"
Option Explicit

Public DXMain As New DirectX7 'Main Object
Public ddMain As DirectDraw7 ' DirectDraw Object
Public diMain As DirectInput 'DirectInput Object
Public diDev As DirectInputDevice 'DirectInput Device
Public diState As DIKEYBOARDSTATE 'Direct Input Keyboard
Public sdMain As DDSURFACEDESC2  'Main Surface
Public dsPrim As DirectDrawSurface7 'Primary Surface
Public dsBbuf As DirectDrawSurface7 'Buffer Surface
Public sdStar As DDSURFACEDESC2 'Star Surface Description
Public dsStar As DirectDrawSurface7 'Star Surface
Public sAngle As Single 'Current Star Angle
Public sSpd As Single 'Speed Multiplier
Public sStar(150, 3) As Single 'Star Array

Public dsShip As DirectDrawSurface7 'Ship Surface
Public sdShip As DDSURFACEDESC2 'Ship Description
Public rShip As RECT

Public dsMissile As DirectDrawSurface7 'Missile Surface
Public sdMissile As DDSURFACEDESC2 'Missile Description
Public rMissile As RECT

Public dsStone As DirectDrawSurface7 'Stone Surface
Public sdStone As DDSURFACEDESC2 'Stone Description
Public rStone As RECT


