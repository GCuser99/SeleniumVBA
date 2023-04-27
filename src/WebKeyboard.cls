VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "This class is used to automate keyboard actions"
'@ModuleDescription "This class is used to automate keyboard actions"
'@Exposed
'@PredeclaredId
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v3.8
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
' Copyright (c) 2023, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
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
'This class was written to handle special keys for SendKeys and Action Chains

Option Explicit

Private mNullKey As String
Private mCancelKey As String
Private mHelpKey As String
Private mBackspaceKey As String
Private mTabKey As String
Private mClearKey As String
Private mReturnKey As String
Private mEnterKey As String
Private mShiftKey As String
Private mControlKey As String
Private mAltKey As String
Private mPauseKey As String
Private mEscapeKey As String
Private mSpaceKey As String
Private mPageUpKey As String
Private mPageDownKey As String
Private mEndKey As String
Private mHomeKey As String
Private mLeftKey As String
Private mUpKey As String
Private mRightKey As String
Private mDownKey As String
Private mInsertKey As String
Private mDeleteKey As String
Private mSemicolonKey As String
Private mEqualsKey As String
Private mNumpad0Key As String
Private mNumpad1Key As String
Private mNumpad2Key As String
Private mNumpad3Key As String
Private mNumpad4Key As String
Private mNumpad5Key As String
Private mNumpad6Key As String
Private mNumpad7Key As String
Private mNumpad8Key As String
Private mNumpad9Key As String
Private mMultiplyKey As String
Private mAddKey As String
Private mSeparatorKey As String
Private mSubtractKey As String
Private mDecimalKey As String
Private mDivideKey As String
Private mF1Key As String
Private mF2Key As String
Private mF3Key As String
Private mF4Key As String
Private mF5Key As String
Private mF6Key As String
Private mF7Key As String
Private mF8Key As String
Private mF9Key As String
Private mF10Key As String
Private mF11Key As String
Private mF12Key As String
Private mMetaKey As String
Private mCommandKey As String

'@Description("Returns the Null (\\ue000) key")
Public Property Get NullKey() As String
Attribute NullKey.VB_Description = "Returns the Null (\\ue000) key"
    NullKey = mNullKey
End Property
'@Description("Returns the Cancel (\\ue001) key")
Public Property Get CancelKey() As String
Attribute CancelKey.VB_Description = "Returns the Cancel (\\ue001) key"
    CancelKey = mCancelKey
End Property
'@Description("Returns the Help (\\ue002) key")
Public Property Get HelpKey() As String
Attribute HelpKey.VB_Description = "Returns the Help (\\ue002) key"
    HelpKey = mHelpKey
End Property
'@Description("Returns the Backspace (\\ue003) key")
Public Property Get BackspaceKey() As String
Attribute BackspaceKey.VB_Description = "Returns the Backspace (\\ue003) key"
    BackspaceKey = mBackspaceKey
End Property
'@Description("Returns the Tab (\\ue004) key")
Public Property Get TabKey() As String
Attribute TabKey.VB_Description = "Returns the Tab (\\ue004) key"
    TabKey = mTabKey
End Property
'@Description("Returns the Clear (\\ue005) key")
Public Property Get ClearKey() As String
Attribute ClearKey.VB_Description = "Returns the Clear (\\ue005) key"
    ClearKey = mClearKey
End Property
'@Description("Returns the Return (\\ue006) key")
Public Property Get ReturnKey() As String
Attribute ReturnKey.VB_Description = "Returns the Return (\\ue006) key"
    ReturnKey = mReturnKey
End Property
'@Description("Returns the Enter (\\ue007) key")
Public Property Get EnterKey() As String
Attribute EnterKey.VB_Description = "Returns the Enter (\\ue007) key"
    EnterKey = mEnterKey
End Property
'@Description("Returns the Shift (\\ue008) key")
Public Property Get ShiftKey() As String
Attribute ShiftKey.VB_Description = "Returns the Shift (\\ue008) key"
    ShiftKey = mShiftKey
End Property
'@Description("Returns the Control (\\ue009) key")
Public Property Get ControlKey() As String
Attribute ControlKey.VB_Description = "Returns the Control (\\ue009) key"
    ControlKey = mControlKey
End Property
'@Description("Returns the Alt (\\ue00a) key")
Public Property Get AltKey() As String
Attribute AltKey.VB_Description = "Returns the Alt (\\ue00a) key"
    AltKey = mAltKey
End Property
'@Description("Returns the Pause (\\ue00b) key")
Public Property Get PauseKey() As String
Attribute PauseKey.VB_Description = "Returns the Pause (\\ue00b) key"
    PauseKey = mPauseKey
End Property
'@Description("Returns the Escape (\\ue00c) key")
Public Property Get EscapeKey() As String
Attribute EscapeKey.VB_Description = "Returns the Escape (\\ue00c) key"
    EscapeKey = mEscapeKey
End Property
'@Description("Returns the Space (\\ue00d) key")
Public Property Get SpaceKey() As String
Attribute SpaceKey.VB_Description = "Returns the Space (\\ue00d) key"
    SpaceKey = mSpaceKey
End Property
'@Description("Returns the PageUp (\\ue00e) key")
Public Property Get PageUpKey() As String
Attribute PageUpKey.VB_Description = "Returns the PageUp (\\ue00e) key"
    PageUpKey = mPageUpKey
End Property
'@Description("Returns the PageDown (\\ue00f) key")
Public Property Get PageDownKey() As String
Attribute PageDownKey.VB_Description = "Returns the PageDown (\\ue00f) key"
    PageDownKey = mPageDownKey
End Property
'@Description("Returns the End (\\ue010) key")
Public Property Get EndKey() As String
Attribute EndKey.VB_Description = "Returns the End (\\ue010) key"
    EndKey = mEndKey
End Property
'@Description("Returns the Home (\\ue011) key")
Public Property Get HomeKey() As String
Attribute HomeKey.VB_Description = "Returns the Home (\\ue011) key"
    HomeKey = mHomeKey
End Property
'@Description("Returns the Left (\\ue012) key")
Public Property Get LeftKey() As String
Attribute LeftKey.VB_Description = "Returns the Left (\\ue012) key"
    LeftKey = mLeftKey
End Property
'@Description("Returns the Up (\\ue013) key")
Public Property Get UpKey() As String
Attribute UpKey.VB_Description = "Returns the Up (\\ue013) key"
    UpKey = mUpKey
End Property
'@Description("Returns the Right (\\ue014) key")
Public Property Get RightKey() As String
Attribute RightKey.VB_Description = "Returns the Right (\\ue014) key"
    RightKey = mRightKey
End Property
'@Description("Returns the Down (\\ue015) key")
Public Property Get DownKey() As String
Attribute DownKey.VB_Description = "Returns the Down (\\ue015) key"
    DownKey = mDownKey
End Property
'@Description("Returns the Insert (\\ue016) key")
Public Property Get InsertKey() As String
Attribute InsertKey.VB_Description = "Returns the Insert (\\ue016) key"
    InsertKey = mInsertKey
End Property
'@Description("Returns the Delete (\\ue017) key")
Public Property Get DeleteKey() As String
Attribute DeleteKey.VB_Description = "Returns the Delete (\\ue017) key"
    DeleteKey = mDeleteKey
End Property
'@Description("Returns the Semicolon (\\ue018) key")
Public Property Get SemicolonKey() As String
Attribute SemicolonKey.VB_Description = "Returns the Semicolon (\\ue018) key"
    SemicolonKey = mSemicolonKey
End Property
'@Description("Returns the Equals (\\ue019) key")
Public Property Get EqualsKey() As String
Attribute EqualsKey.VB_Description = "Returns the Equals (\\ue019) key"
    EqualsKey = mEqualsKey
End Property
'@Description("Returns the Numpad 0 (\\ue01a) key")
Public Property Get Numpad0Key() As String
Attribute Numpad0Key.VB_Description = "Returns the Numpad 0 (\\ue01a) key"
    Numpad0Key = mNumpad0Key
End Property
'@Description("Returns the Numpad 1 (\\ue01b) key")
Public Property Get Numpad1Key() As String
Attribute Numpad1Key.VB_Description = "Returns the Numpad 1 (\\ue01b) key"
    Numpad1Key = mNumpad1Key
End Property
'@Description("Returns the Numpad 2 (\\ue01c) key")
Public Property Get Numpad2Key() As String
Attribute Numpad2Key.VB_Description = "Returns the Numpad 2 (\\ue01c) key"
    Numpad2Key = mNumpad2Key
End Property
'@Description("Returns the Numpad 3 (\\ue01d) key")
Public Property Get Numpad3Key() As String
Attribute Numpad3Key.VB_Description = "Returns the Numpad 3 (\\ue01d) key"
    Numpad3Key = mNumpad3Key
End Property
'@Description("Returns the Numpad 4 (\\ue01e) key")
Public Property Get Numpad4Key() As String
Attribute Numpad4Key.VB_Description = "Returns the Numpad 4 (\\ue01e) key"
    Numpad4Key = mNumpad4Key
End Property
'@Description("Returns the Numpad 5 (\\ue01f) key")
Public Property Get Numpad5Key() As String
Attribute Numpad5Key.VB_Description = "Returns the Numpad 5 (\\ue01f) key"
    Numpad5Key = mNumpad5Key
End Property
'@Description("Returns the Numpad 6 (\\ue020) key")
Public Property Get Numpad6Key() As String
Attribute Numpad6Key.VB_Description = "Returns the Numpad 6 (\\ue020) key"
    Numpad6Key = mNumpad6Key
End Property
'@Description("Returns the Numpad 7 (\\ue021) key")
Public Property Get Numpad7Key() As String
Attribute Numpad7Key.VB_Description = "Returns the Numpad 7 (\\ue021) key"
    Numpad7Key = mNumpad7Key
End Property
'@Description("Returns the Numpad 8 (\\ue022) key")
Public Property Get Numpad8Key() As String
Attribute Numpad8Key.VB_Description = "Returns the Numpad 8 (\\ue022) key"
    Numpad8Key = mNumpad8Key
End Property
'@Description("Returns the Numpad 9 (\\ue023) key")
Public Property Get Numpad9Key() As String
Attribute Numpad9Key.VB_Description = "Returns the Numpad 9 (\\ue023) key"
    Numpad9Key = mNumpad9Key
End Property
'@Description("Returns the Multiply (\\ue024) key")
Public Property Get MultiplyKey() As String
Attribute MultiplyKey.VB_Description = "Returns the Multiply (\\ue024) key"
    MultiplyKey = mMultiplyKey
End Property
'@Description("Returns the Add (\\ue025) key")
Public Property Get AddKey() As String
Attribute AddKey.VB_Description = "Returns the Add (\\ue025) key"
    AddKey = mAddKey
End Property
'@Description("Returns the Separator (\\ue026) key")
Public Property Get SeparatorKey() As String
Attribute SeparatorKey.VB_Description = "Returns the Separator (\\ue026) key"
    SeparatorKey = mSeparatorKey
End Property
'@Description("Returns the Subtract (\\ue027) key")
Public Property Get SubtractKey() As String
Attribute SubtractKey.VB_Description = "Returns the Subtract (\\ue027) key"
    SubtractKey = mSubtractKey
End Property
'@Description("Returns the Decimal (\\ue028) key")
Public Property Get DecimalKey() As String
Attribute DecimalKey.VB_Description = "Returns the Decimal (\\ue028) key"
    DecimalKey = mDecimalKey
End Property
'@Description("Returns the Divide (\\ue029) key")
Public Property Get DivideKey() As String
Attribute DivideKey.VB_Description = "Returns the Divide (\\ue029) key"
    DivideKey = mDivideKey
End Property
'@Description("Returns the F1 (\\ue031) key")
Public Property Get F1Key() As String
Attribute F1Key.VB_Description = "Returns the F1 (\\ue031) key"
    F1Key = mF1Key
End Property
'@Description("Returns the F2 (\\ue032) key")
Public Property Get F2Key() As String
Attribute F2Key.VB_Description = "Returns the F2 (\\ue032) key"
    F2Key = mF2Key
End Property
'@Description("Returns the F3 (\\ue033) key")
Public Property Get F3Key() As String
Attribute F3Key.VB_Description = "Returns the F3 (\\ue033) key"
    F3Key = mF3Key
End Property
'@Description("Returns the F4 (\\ue034) key")
Public Property Get F4Key() As String
Attribute F4Key.VB_Description = "Returns the F4 (\\ue034) key"
    F4Key = mF4Key
End Property
'@Description("Returns the F5 (\\ue035) key")
Public Property Get F5Key() As String
Attribute F5Key.VB_Description = "Returns the F5 (\\ue035) key"
    F5Key = mF5Key
End Property
'@Description("Returns the F6 (\\ue036) key")
Public Property Get F6Key() As String
Attribute F6Key.VB_Description = "Returns the F6 (\\ue036) key"
    F6Key = mF6Key
End Property
'@Description("Returns the F7 (\\ue037) key")
Public Property Get F7Key() As String
Attribute F7Key.VB_Description = "Returns the F7 (\\ue037) key"
    F7Key = mF7Key
End Property
'@Description("Returns the F8 (\\ue038) key")
Public Property Get F8Key() As String
Attribute F8Key.VB_Description = "Returns the F8 (\\ue038) key"
    F8Key = mF8Key
End Property
'@Description("Returns the F9 (\\ue039) key")
Public Property Get F9Key() As String
Attribute F9Key.VB_Description = "Returns the F9 (\\ue039) key"
    F9Key = mF9Key
End Property
'@Description("Returns the F10 (\\ue03a) key")
Public Property Get F10Key() As String
Attribute F10Key.VB_Description = "Returns the F10 (\\ue03a) key"
    F10Key = mF10Key
End Property
'@Description("Returns the F11 (\\ue03b) key")
Public Property Get F11Key() As String
Attribute F11Key.VB_Description = "Returns the F11 (\\ue03b) key"
    F11Key = mF11Key
End Property
'@Description("Returns the F12 (\\ue03c) key")
Public Property Get F12Key() As String
Attribute F12Key.VB_Description = "Returns the F12 (\\ue03c) key"
    F12Key = mF12Key
End Property
'@Description("Returns the Meta (\\ue03d) key")
Public Property Get MetaKey() As String
Attribute MetaKey.VB_Description = "Returns the Meta (\\ue03d) key"
    MetaKey = mMetaKey
End Property
'@Description("Returns the Command (\\ue03d) key")
Public Property Get CommandKey() As String
Attribute CommandKey.VB_Description = "Returns the Command (\\ue03d) key"
    CommandKey = mCommandKey
End Property

Private Sub Class_Initialize()
    mNullKey = "\ue000"
    mCancelKey = "\ue001"
    mHelpKey = "\ue002"
    mBackspaceKey = "\ue003"
    mTabKey = "\ue004"
    mClearKey = "\ue005"
    mReturnKey = "\ue006"
    mEnterKey = "\ue007"
    mShiftKey = "\ue008"
    mControlKey = "\ue009"
    mAltKey = "\ue00a"
    mPauseKey = "\ue00b"
    mEscapeKey = "\ue00c"
    mSpaceKey = "\ue00d"
    mPageUpKey = "\ue00e"
    mPageDownKey = "\ue00f"
    mEndKey = "\ue010"
    mHomeKey = "\ue011"
    mLeftKey = "\ue012"
    mUpKey = "\ue013"
    mRightKey = "\ue014"
    mDownKey = "\ue015"
    mInsertKey = "\ue016"
    mDeleteKey = "\ue017"
    mSemicolonKey = "\ue018"
    mEqualsKey = "\ue019"
    mNumpad0Key = "\ue01a"
    mNumpad1Key = "\ue01b"
    mNumpad2Key = "\ue01c"
    mNumpad3Key = "\ue01d"
    mNumpad4Key = "\ue01e"
    mNumpad5Key = "\ue01f"
    mNumpad6Key = "\ue020"
    mNumpad7Key = "\ue021"
    mNumpad8Key = "\ue022"
    mNumpad9Key = "\ue023"
    mMultiplyKey = "\ue024"
    mAddKey = "\ue025"
    mSeparatorKey = "\ue026"
    mSubtractKey = "\ue027"
    mDecimalKey = "\ue028"
    mDivideKey = "\ue029"
    mF1Key = "\ue031"
    mF2Key = "\ue032"
    mF3Key = "\ue033"
    mF4Key = "\ue034"
    mF5Key = "\ue035"
    mF6Key = "\ue036"
    mF7Key = "\ue037"
    mF8Key = "\ue038"
    mF9Key = "\ue039"
    mF10Key = "\ue03a"
    mF11Key = "\ue03b"
    mF12Key = "\ue03c"
    mMetaKey = "\ue03d"
    mCommandKey = "\ue03d"
End Sub
