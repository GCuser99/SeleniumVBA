VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebPrintSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to manage print settings for generating pdf files using PrintToPDF method of the WebDriver class"
'@ModuleDescription "This class is used to manage print settings for generating pdf files using PrintToPDF method of the WebDriver class"
'@Exposed
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v6.3
'
' A Selenium wrapper for browser automation developed for MS Office VBA
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' Contact Info:
'
' https://github.com/6DiegoDiego9
' https://github.com/GCUser99
' ==========================================================================
' MIT License
'
' Copyright (c) 2023-2024, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
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
Private mOrientation  As svbaOrientation
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

'@Description("Sets or gets the measurement units in centimeters or inches")
Public Property Get Units() As svbaUnits
Attribute Units.VB_Description = "Sets or gets the measurement units in centimeters or inches"
    Units = mUnits
End Property
Public Property Let Units(ByVal val As svbaUnits)
    mUnits = val
End Property

'@Description("Sets or gets the bottom margin")
Public Property Get MarginBottom() As Single
Attribute MarginBottom.VB_Description = "Sets or gets the bottom margin"
    MarginBottom = mMarginBottom
End Property
Public Property Let MarginBottom(ByVal val As Single)
    mMarginBottom = val
End Property

'@Description("Sets or gets the top margin")
Public Property Get MarginTop() As Single
Attribute MarginTop.VB_Description = "Sets or gets the top margin"
    MarginTop = mMarginTop
End Property
Public Property Let MarginTop(ByVal val As Single)
    mMarginTop = val
End Property

'@Description("Sets or gets the left margin")
Public Property Get MarginLeft() As Single
Attribute MarginLeft.VB_Description = "Sets or gets the left margin"
    MarginLeft = mMarginLeft
End Property
Public Property Let MarginLeft(ByVal val As Single)
    mMarginLeft = val
End Property

'@Description("Sets or gets the right margin")
Public Property Get MarginRight() As Single
Attribute MarginRight.VB_Description = "Sets or gets the right margin"
    MarginRight = mMarginRight
End Property
Public Property Let MarginRight(ByVal val As Single)
    mMarginRight = val
End Property

'@Description("Sets all margins to same value")
Public Property Let MarginsAll(ByVal val As Single)
Attribute MarginsAll.VB_Description = "Sets all margins to same value"
    mMarginBottom = val
    mMarginTop = val
    mMarginLeft = val
    mMarginRight = val
End Property

'@Description("Sets or gets the page height")
Public Property Get PageHeight() As Single
Attribute PageHeight.VB_Description = "Sets or gets the page height"
    PageHeight = mPageHeight
End Property
Public Property Let PageHeight(ByVal val As Single)
    mPageHeight = val
End Property

'@Description("Sets or gets the page width")
Public Property Get PageWidth() As Single
Attribute PageWidth.VB_Description = "Sets or gets the page width"
    PageWidth = mPageWidth
End Property
Public Property Let PageWidth(ByVal val As Single)
    mPageWidth = val
End Property

'@Description("Sets or gets the background property")
Public Property Get Background() As Boolean
Attribute Background.VB_Description = "Sets or gets the background property"
    Background = mBackground
End Property
Public Property Let Background(ByVal val As Boolean)
    mBackground = val
End Property

'@Description("Sets or gets the page orientation")
Public Property Get Orientation() As svbaOrientation
Attribute Orientation.VB_Description = "Sets or gets the page orientation"
    Orientation = mOrientation
End Property
Public Property Let Orientation(ByVal val As svbaOrientation)
    mOrientation = val
End Property

'@Description("Sets a list of pages and page ranges to print")
Public Sub PageRanges(ParamArray pages() As Variant)
Attribute PageRanges.VB_Description = "Sets a list of pages and page ranges to print"
    mPageRanges = pages
End Sub

'@Description("Sets or gets the print scale (fractional)")
Public Property Get PrintScale() As Single
Attribute PrintScale.VB_Description = "Sets or gets the print scale (fractional)"
    PrintScale = mPrintScale
End Property
Public Property Let PrintScale(ByVal val As Single)
    mPrintScale = val
End Property

'@Description("Sets or gets the shrink-to-fit property")
Public Property Get ShrinkToFit() As Boolean
Attribute ShrinkToFit.VB_Description = "Sets or gets the shrink-to-fit property"
    ShrinkToFit = mShrinkToFit
End Property
Public Property Let ShrinkToFit(ByVal val As Boolean)
    mShrinkToFit = val
End Property

Friend Function ToDictionary() As Dictionary
    Dim margins As New Dictionary
    Dim page As New Dictionary
    Dim data As New Dictionary
    
    Dim ufact As Single
    If mUnits = svbaInches Then ufact = 2.54 Else ufact = 1
    
    margins.Add "bottom", mMarginBottom * ufact
    margins.Add "left", mMarginLeft * ufact
    margins.Add "right", mMarginRight * ufact
    margins.Add "top", mMarginTop * ufact
    
    page.Add "height", mPageHeight * ufact
    page.Add "width", mPageWidth * ufact
    
    data.Add "background", mBackground
    
    If mOrientation = svbaPortrait Then
        data.Add "orientation", "portrait"
    Else
        data.Add "orientation", "landscape"
    End If
    
    Dim i As Long
    For i = 0 To UBound(mPageRanges)
        mPageRanges(i) = CStr(mPageRanges(i))
    Next i
    
    data.Add "margin", margins
    data.Add "page", page
    data.Add "pageRanges", mPageRanges 'eg. Array("1-2", "4-7", 13, "15")
    data.Add "scale", mPrintScale
    data.Add "shrinkToFit", mShrinkToFit
    
    Set ToDictionary = data
End Function

Private Sub Class_Initialize()
    Dim iniFilePath As String
    
    'attempt to read global variable values from optional settings file - if not successful,
    'then assign a system default value (last input parameter of ReadIniFileEntry)
    
    iniFilePath = getFullLocalPath(thisLibFolderPath) & "\SeleniumVBA.ini"
    
    mUnits = enumTextToValue(readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "units", svbaInches))
    mMarginBottom = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_bottom", 1 / 2.54)
    mMarginTop = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_top", 1 / 2.54)
    mMarginLeft = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_left", 1 / 2.54)
    mMarginRight = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "margin_right", 1 / 2.54)
    mPageHeight = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "page_height", 11)
    mPageWidth = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "page_width", 8.5)
    mBackground = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "background", False)
    mOrientation = enumTextToValue(readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "orientation", svbaPortrait))
    mPageRanges = Array()
    mPrintScale = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "print_scale", 1)
    mShrinkToFit = readIniFileEntry(iniFilePath, "PDF_DEFAULT_PRINT_SETTINGS", "shrink_to_fit", True)
End Sub
