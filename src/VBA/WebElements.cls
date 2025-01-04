VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used for managing a collection of WebElement objects"
'@ModuleDescription "This class is used for managing a collection of WebElement objects"
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
'
' This class was written to support collection of WebElement objects
'
' ==========================================================================

Option Explicit

Private colWebElement As Collection

'@Description("Adds an element to the collection")
Public Function Add(element As WebElement) As WebElement
Attribute Add.VB_Description = "Adds an element to the collection"
    colWebElement.Add element, element.elementId
    Set Add = element
End Function

'@Description("Returns the number of elements in the collection")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of elements in the collection"
   Count = colWebElement.Count
End Property

'@Description("Returns the number of elements in the collection")
Public Property Get Size() As Long
Attribute Size.VB_Description = "Returns the number of elements in the collection"
   Size = colWebElement.Count
End Property

'@Description("Removes an element corresponding to the given index or element object")
Public Sub Remove(indexOrWebElement As Variant)
Attribute Remove.VB_Description = "Removes an element corresponding to the given index or element object"
    If TypeName(indexOrWebElement) = "WebElement" Then
        Dim welem As WebElement
        Set welem = indexOrWebElement 'Friend property ElementID can not be invoked with late binding
        colWebElement.Remove welem.elementId
    Else
        colWebElement.Remove indexOrWebElement
    End If
End Sub

'@Description("Returns the element corresponding to the given index")
'@DefaultMember
Public Property Get Item(ByVal index As Variant) As WebElement
Attribute Item.VB_Description = "Returns the element corresponding to the given index"
Attribute Item.VB_UserMemId = 0
    Set Item = colWebElement.Item(index)
End Property

'@Description("Returns True if element object is already in the collection, False otherwise")
Public Function Exists(element As WebElement) As Boolean
Attribute Exists.VB_Description = "Returns True if element object is already in the collection, False otherwise"
    On Error GoTo notFound
    Exists = True
    colWebElement.Item element.elementId
    Exit Function
notFound:
    Exists = False
End Function

'@Description("Removes all member element objects from the collection")
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all member element objects from the collection"
    Do While colWebElement.Count > 0
        colWebElement.Remove 1
    Loop
End Sub

'@Description("Highlights this element collection with color")
Public Function Highlight(Optional ByVal borderSizePx As Byte = 4, Optional ByVal borderColor As VBAcolors = Cyan, _
                     Optional ByVal backgroundColor As VBAcolors = VBAcolors.Unchanged, _
                     Optional ByVal ScrollIntoView As Boolean = True, Optional ByVal unHighlightLast As Boolean = True) _
                     As WebElements
Attribute Highlight.VB_Description = "Highlights this element collection with color"
    Dim i As Long
    colWebElement(1).Highlight borderSizePx, borderColor, backgroundColor, ScrollIntoView, unHighlightLast
    For i = 2 To colWebElement.Count
        colWebElement(i).Highlight borderSizePx, borderColor, backgroundColor, ScrollIntoView, False
    Next i
    Set Highlight = Me
End Function

'@Description("Unhighlights this element collection")
Public Sub UnHighlight()
Attribute UnHighlight.VB_Description = "Unhighlights this element collection"
    Dim elem As WebElement
    For Each elem In colWebElement
        elem.UnHighlight
    Next elem
End Sub

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colWebElement.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set colWebElement = New Collection
End Sub

Private Sub Class_Terminate()
    Set colWebElement = Nothing
End Sub
