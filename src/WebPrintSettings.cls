VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebPrintSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v1.8
' A Selenium wrapper for Edge and Chrome written in Windows VBA based on JSon wire protocol.
'
' (c) GCUser99
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' ==========================================================================
' MIT License
'
' Copyright (c) 2022 GCUser99
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================

Option Explicit

Private mMarginBottom As Single
Private mMarginTop As Single
Private mMarginLeft As Single
Private mMarginRight As Single
Private mPageHeight  As Single
Private mPageWidth  As Single
Private mBackground As Boolean
Private mOrientation  As String
Private mPageRanges() As Variant
Private mPrintScale  As Single
Private mShrinkToFit As Boolean
Private mUnits As svbaUnits

Public Enum svbaOrientation
    svbaLandscape = 0
    svbaPortrait = 1
End Enum

Public Enum svbaUnits
    svbaCentimeters = 0
    svbaInches = 1
End Enum

Public Property Get Units() As svbaUnits
    Units = mUnits
End Property
Public Property Let Units(ByVal val As svbaUnits)
    mUnits = val
End Property

Public Property Get MarginBottom() As Single
    MarginBottom = mMarginBottom
End Property
Public Property Let MarginBottom(ByVal val As Single)
    mMarginBottom = val
End Property

Public Property Get MarginTop() As Single
    MarginTop = mMarginTop
End Property
Public Property Let MarginTop(ByVal val As Single)
    mMarginTop = val
End Property

Public Property Get MarginLeft() As Single
    MarginLeft = mMarginLeft
End Property
Public Property Let MarginLeft(ByVal val As Single)
    mMarginLeft = val
End Property

Public Property Get MarginRight() As Single
    MarginRight = mMarginRight
End Property
Public Property Let MarginRight(ByVal val As Single)
    mMarginRight = val
End Property

Public Property Let MarginsAll(ByVal val As Single)
    mMarginBottom = val
    mMarginTop = val
    mMarginLeft = val
    mMarginRight = val
End Property

Public Property Get PageHeight() As Single
    PageHeight = mPageHeight
End Property
Public Property Let PageHeight(ByVal val As Single)
    mPageHeight = val
End Property

Public Property Get PageWidth() As Single
    PageWidth = mPageWidth
End Property
Public Property Let PageWidth(ByVal val As Single)
    mPageWidth = val
End Property

Public Property Get Background() As Boolean
    Background = mBackground
End Property
Public Property Let Background(ByVal val As Boolean)
    mBackground = val
End Property

Public Property Get Orientation() As svbaOrientation
    Orientation = mOrientation
End Property
Public Property Let Orientation(ByVal val As svbaOrientation)
    mOrientation = val
End Property

Public Function PageRanges(ParamArray pages() As Variant)
    mPageRanges = pages
End Function

Public Property Get PrintScale() As Single
    PrintScale = mPrintScale
End Property
Public Property Let PrintScale(ByVal val As Single)
    mPrintScale = val
End Property

Public Property Get ShrinkToFit() As Boolean
    ShrinkToFit = mShrinkToFit
End Property
Public Property Let ShrinkToFit(ByVal val As Boolean)
    mShrinkToFit = val
End Property

Friend Function ToDictionary() As Dictionary
    Dim margins As New Dictionary
    Dim page As New Dictionary
    Dim Data As New Dictionary
    
    Dim ufact As Single
    If mUnits = svbaInches Then ufact = 2.54 Else ufact = 1
    
    margins.Add "bottom", mMarginBottom * ufact
    margins.Add "left", mMarginLeft * ufact
    margins.Add "right", mMarginRight * ufact
    margins.Add "top", mMarginTop * ufact
    
    page.Add "height", mPageHeight * ufact
    page.Add "width", mPageWidth * ufact
    
    Data.Add "background", mBackground
    
    If mOrientation = svbaPortrait Then
        Data.Add "orientation", "portrait"
    Else
        Data.Add "orientation", "landscape"
    End If
    
    Dim i As Long
    For i = 0 To UBound(mPageRanges)
        mPageRanges(i) = CStr(mPageRanges(i))
    Next i
    
    Data.Add "margin", margins
    Data.Add "page", page
    Data.Add "pageRanges", mPageRanges 'eg. Array("1-2", "4-7", 13, "15")
    Data.Add "scale", mPrintScale
    Data.Add "shrinkToFit", mShrinkToFit
    
    Set ToDictionary = Data
End Function

Private Sub Class_Initialize()
    mUnits = svbaInches
    mMarginBottom = 1 / 2.54
    mMarginTop = 1 / 2.54
    mMarginLeft = 1 / 2.54
    mMarginRight = 1 / 2.54
    mPageHeight = 11
    mPageWidth = 8.5
    mBackground = False
    mOrientation = svbaPortrait
    mPageRanges = Array()
    mPrintScale = 1#
    mShrinkToFit = True
End Sub

