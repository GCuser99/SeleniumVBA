VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShadowRoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'this class is written to support Shadow Roots

Private Driver_ As WebDriver
Private SessionId_ As String
Private ShadowRootId_ As String

Friend Property Set Driver(objWebDriver As WebDriver)
    Set Driver_ = objWebDriver
End Property

Friend Property Let SessionId(ByVal val)
    SessionId_ = val
End Property

Friend Property Let ShadowRootId(ByVal val)
    ShadowRootId_ = val
End Property

Friend Property Get ShadowRootId()
    ShadowRootId = ShadowRootId_
End Property

Public Function FindElement(by_ As by, ByVal Value As String) As WebElement
    Set FindElement = Driver_.FindElementFromShadowRoot(by_, Value, Me, SessionId_)
End Function

Public Function FindElements(by_ As by, ByVal Value As String) As WebElement()
    FindElements = Driver_.FindElementsFromShadowRoot(by_, Value, Me, SessionId_)
End Function


