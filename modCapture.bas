Attribute VB_Name = "modCapture"
Option Explicit

Declare Sub SetCapture Lib "user32" (ByVal hWnd As Long)
Declare Sub ReleaseCapture Lib "user32" ()
