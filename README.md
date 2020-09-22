<div align="center">

## Screenshot Without Using Clipboard Or DIB


</div>

### Description

Seen many posts using api for printscreen keypress, but what about the info on clipboard? Seen others that copy clipboard to temp var, and set back after, but why waste resources and processing when you can do it the right way? Also seen classes using DIB, wayyy slow. This = best way! =D

This uses two tricks that arent very known, so here you go! When using GetDC(0) as our source, we actually capture the screen, and when using picture.image=picture.picture.
 
### More Info
 
Always make sure the Picturebox has AutoRedraw enabled. Create a Command1, and in it call ScreenShot()

A screenshot in the app folder called screenshot.bmp


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[DanDaMan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dandaman.md)
**Level**          |Intermediate
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dandaman-screenshot-without-using-clipboard-or-dib__1-61855/archive/master.zip)

### API Declarations

Bitblt and GetDC


### Source Code

```
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Sub ScreenShot()
  'Important Precursors
  'All these can be set in form load
  'The Picturebox can also be Visible=False
  'and this will still work
  Picture1.Width = Screen.Width
  Picture1.Height = Screen.Height
  Picture1.AutoRedraw = True
  BitBlt Picture1.hDC, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, GetDC(0), 0, 0, vbSrcCopy
  Picture1.Picture = Picture1.Image
  SavePicture Picture1.Picture, App.Path & "/screen.bmp"
End Sub
```

