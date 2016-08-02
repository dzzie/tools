Attribute VB_Name = "MTransBlt"
' *************************************************************
'  Copyright (C)1997, Karl E. Peterson
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************
Option Explicit
'
' Win32 API Declarations, Structures, and Constants
'
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
'
' Constants used by new transparent support in NT.
'
Private Const CAPS1 = 94                 '  other caps
Private Const C1_TRANSPARENT = &H1       '  new raster cap
Private Const NEWTRANSPARENT = 3         '  use with SetBkMode()
Private Const OBJ_BITMAP = 7             '  used to retrieve HBITMAP from HDC
'
' Ternary raster operations
'
Private Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
'
' API structure definition for Rectangle
'
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'
' Bitmap Header Definition
'
Private Type BITMAP '14 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Public Sub TileBlt(ByVal hWndDest As Long, ByVal hBmpSrc As Long)
   '
   ' 32-Bit Tiling BitBlt Function
   ' Written by Karl E. Peterson, 9/22/96.
   ' Tiles a bitmap across the client area of destination window.
   '
   ' Parameters ************************************************************
   '   hWndDest:     hWnd of destination
   '   hBmpSrc:      hBitmap of source
   ' ***********************************************************************
   '
   Dim bmp As BITMAP     ' Header info for passed bitmap handle
   Dim hDCSrc As Long    ' Device context for source
   Dim hDCDest As Long   ' Device context for destination
   Dim hBmpTmp As Long   ' Holding space for temporary bitmap
   Dim dRect As RECT     ' Holds coordinates of destination rectangle
   Dim Rows As Long      ' Number of rows in destination
   Dim Cols As Long      ' Number of columns in destination
   Dim dX As Long        ' CurrentX in destination
   Dim dY As Long        ' CurrentY in destination
   Dim i As Long, j As Long
   '
   ' Get destination rectangle and device context.
   '
   Call GetClientRect(hWndDest, dRect)
   hDCDest = GetDC(hWndDest)
   '
   ' Create source DC and select passed bitmap into it.
   '
   hDCSrc = CreateCompatibleDC(hDCDest)
   hBmpTmp = SelectObject(hDCSrc, hBmpSrc)
   '
   ' Get size information about passed bitmap, and
   ' Calc number of rows and columns to paint.
   '
   Call GetObj(hBmpSrc, Len(bmp), bmp)
   Rows = dRect.Right \ bmp.bmWidth
   Cols = dRect.Bottom \ bmp.bmHeight
   '
   ' Spray out across destination.
   '
   For i = 0 To Rows
      dX = i * bmp.bmWidth
      For j = 0 To Cols
         dY = j * bmp.bmHeight
         Call BitBlt(hDCDest, dX, dY, bmp.bmWidth, bmp.bmHeight, hDCSrc, 0, 0, SRCCOPY)
      Next j
   Next i
   '
   ' and clean up
   '
   Call SelectObject(hDCSrc, hBmpTmp)
   Call DeleteDC(hDCSrc)
   Call ReleaseDC(hWndDest, hDCDest)
End Sub

Public Sub TransBlt(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal TransColor As Long)
   '
   ' 32-Bit Transparent BitBlt Function
   ' Written by Karl E. Peterson, 9/20/96.
   ' Portions borrowed and modified from KB.
   ' Other portions modified following input from users. <g>
   '
   ' Parameters ************************************************************
   '   hDestDC:     Destination device context
   '   x, y:        Upper-left destination coordinates (pixels)
   '   nWidth:      Width of destination
   '   nHeight:     Height of destination
   '   hSrcDC:      Source device context
   '   xSrc, ySrc:  Upper-left source coordinates (pixels)
   '   TransColor:  RGB value for transparent pixels, typically &HC0C0C0.
   ' ***********************************************************************
   '
   Dim OrigColor As Long      ' Holds original background color
   Dim OrigMode As Long       ' Holds original background drawing mode
   
   If (GetDeviceCaps(hDestDC, CAPS1) And C1_TRANSPARENT) Then
      '
      ' Some NT machines support this *super* simple method!
      ' Save original settings, Blt, restore settings.
      '
      OrigMode = SetBkMode(hDestDC, NEWTRANSPARENT)
      OrigColor = SetBkColor(hDestDC, TransColor)
      Call BitBlt(hDestDC, x, y, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY)
      Call SetBkColor(hDestDC, OrigColor)
      Call SetBkMode(hDestDC, OrigMode)
      
   Else
      Dim saveDC As Long         ' Backup copy of source bitmap
      Dim maskDC As Long         ' Mask bitmap (monochrome)
      Dim invDC As Long          ' Inverse of mask bitmap (monochrome)
      Dim resultDC As Long       ' Combination of source bitmap & background
      Dim hSaveBmp As Long       ' Bitmap stores backup copy of source bitmap
      Dim hMaskBmp As Long       ' Bitmap stores mask (monochrome)
      Dim hInvBmp As Long        ' Bitmap holds inverse of mask (monochrome)
      Dim hResultBmp As Long     ' Bitmap combination of source & background
      Dim hSavePrevBmp As Long   ' Holds previous bitmap in saved DC
      Dim hMaskPrevBmp As Long   ' Holds previous bitmap in the mask DC
      Dim hInvPrevBmp As Long    ' Holds previous bitmap in inverted mask DC
      Dim hDestPrevBmp As Long   ' Holds previous bitmap in destination DC
      '
      ' Create DCs to hold various stages of transformation.
      '
      saveDC = CreateCompatibleDC(hDestDC)
      maskDC = CreateCompatibleDC(hDestDC)
      invDC = CreateCompatibleDC(hDestDC)
      resultDC = CreateCompatibleDC(hDestDC)
      '
      ' Create monochrome bitmaps for the mask-related bitmaps.
      '
      hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
      hInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
      '
      ' Create color bitmaps for final result & stored copy of source.
      '
      hResultBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
      hSaveBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
      '
      ' Select bitmaps into DCs.
      '
      hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
      hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
      hInvPrevBmp = SelectObject(invDC, hInvBmp)
      hDestPrevBmp = SelectObject(resultDC, hResultBmp)
      '
      ' Create mask: set background color of source to transparent color.
      '
      OrigColor = SetBkColor(hSrcDC, TransColor)
      Call BitBlt(maskDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, vbSrcCopy)
      TransColor = SetBkColor(hSrcDC, OrigColor)
      '
      ' Create inverse of mask to AND w/ source & combine w/ background.
      '
      Call BitBlt(invDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbNotSrcCopy)
      '
      ' Copy background bitmap to result.
      '
      Call BitBlt(resultDC, 0, 0, nWidth, nHeight, hDestDC, x, y, vbSrcCopy)
      '
      ' AND mask bitmap w/ result DC to punch hole in the background by
      ' painting black area for non-transparent portion of source bitmap.
      '
      Call BitBlt(resultDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbSrcAnd)
      '
      ' get overlapper
      '
      Call BitBlt(saveDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, vbSrcCopy)
      '
      ' AND with inverse monochrome mask
      '
      Call BitBlt(saveDC, 0, 0, nWidth, nHeight, invDC, 0, 0, vbSrcAnd)
      '
      ' XOR these two
      '
      Call BitBlt(resultDC, 0, 0, nWidth, nHeight, saveDC, 0, 0, vbSrcInvert)
      '
      ' Display transparent bitmap on background.
      '
      Call BitBlt(hDestDC, x, y, nWidth, nHeight, resultDC, 0, 0, vbSrcCopy)
      '
      ' Select original objects back.
      '
      Call SelectObject(saveDC, hSavePrevBmp)
      Call SelectObject(resultDC, hDestPrevBmp)
      Call SelectObject(maskDC, hMaskPrevBmp)
      Call SelectObject(invDC, hInvPrevBmp)
      '
      ' Deallocate system resources.
      '
      Call DeleteObject(hSaveBmp)
      Call DeleteObject(hMaskBmp)
      Call DeleteObject(hInvBmp)
      Call DeleteObject(hResultBmp)
      Call DeleteDC(saveDC)
      Call DeleteDC(invDC)
      Call DeleteDC(maskDC)
      Call DeleteDC(resultDC)
   End If
End Sub

