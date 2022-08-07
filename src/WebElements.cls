VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v1.6
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
'
' This class was written to support collection of WebElement objects
'
' ==========================================================================

Option Explicit

Private colWebElement As Collection

Public Function Add(element As WebElement) As WebElement
    colWebElement.Add element, element.ElementId
    Set Add = element
End Function

Public Property Get Count() As Integer
   Count = colWebElement.Count
End Property

Public Property Get Size() As Integer
   Size = colWebElement.Count
End Property

Public Sub Remove(indexOrWebElement As Variant)
    If TypeName(indexOrWebElement) = "WebElement" Then
        Dim welem As WebElement
        Set welem = indexOrWebElement 'Friend property ElementID can not be invoked with late binding
        colWebElement.Remove welem.ElementId
    Else
        colWebElement.Remove indexOrWebElement
    End If
End Sub

Public Property Get Item(ByVal index As Variant) As WebElement
Attribute Item.VB_UserMemId = 0
    'Attribute Dictionary.VB_UserMemId = 0
    'uncomment line above in text editor to make this the default property
    Set Item = colWebElement.Item(index)
End Property

Friend Sub RemoveAll()
    Do While colWebElement.Count > 0
        colWebElement.Remove 1
    Loop
End Sub

Public Function IsMember(element As WebElement) As Boolean
    Dim i As Long, elemId As String
    elemId = element.ElementId
    IsMember = False
    For i = 1 To Me.Count
        If Me.Item(i).ElementId = elemId Then
            IsMember = True
            Exit For
        End If
    Next i
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_UserMemID = -4
    'uncomment line above in text editor to enable For Each iteration in for loop
    Set NewEnum = colWebElement.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set colWebElement = New Collection
End Sub

Private Sub Class_Terminate()
    RemoveAll
    Set colWebElement = Nothing
End Sub


