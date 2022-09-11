VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================================================
' SeleniumVBA v2.0
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
Private mPage_UpKey As String
Private mPage_DownKey As String
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

Property Get NullKey() As String
    NullKey = mNullKey
End Property
Property Get CancelKey() As String
    CancelKey = mCancelKey
End Property
Property Get HelpKey() As String
    HelpKey = mHelpKey
End Property
Property Get BackspaceKey() As String
    BackspaceKey = mBackspaceKey
End Property
Property Get TabKey() As String
    TabKey = mTabKey
End Property
Property Get ClearKey() As String
    ClearKey = mClearKey
End Property
Property Get ReturnKey() As String
    ReturnKey = mReturnKey
End Property
Property Get EnterKey() As String
    EnterKey = mEnterKey
End Property
Property Get ShiftKey() As String
    ShiftKey = mShiftKey
End Property
Property Get ControlKey() As String
    ControlKey = mControlKey
End Property
Property Get AltKey() As String
    AltKey = mAltKey
End Property
Property Get PauseKey() As String
    PauseKey = mPauseKey
End Property
Property Get EscapeKey() As String
    EscapeKey = mEscapeKey
End Property
Property Get SpaceKey() As String
    SpaceKey = mSpaceKey
End Property
Property Get Page_UpKey() As String
    Page_UpKey = mPage_UpKey
End Property
Property Get Page_DownKey() As String
    Page_DownKey = mPage_DownKey
End Property
Property Get EndKey() As String
    EndKey = mEndKey
End Property
Property Get HomeKey() As String
    HomeKey = mHomeKey
End Property
Property Get LeftKey() As String
    LeftKey = mLeftKey
End Property
Property Get UpKey() As String
    UpKey = mUpKey
End Property
Property Get RightKey() As String
    RightKey = mRightKey
End Property
Property Get DownKey() As String
    DownKey = mDownKey
End Property
Property Get InsertKey() As String
    InsertKey = mInsertKey
End Property
Property Get DeleteKey() As String
    DeleteKey = mDeleteKey
End Property
Property Get SemicolonKey() As String
    SemicolonKey = mSemicolonKey
End Property
Property Get EqualsKey() As String
    EqualsKey = mEqualsKey
End Property
Property Get Numpad0Key() As String
    Numpad0Key = mNumpad0Key
End Property
Property Get Numpad1Key() As String
    Numpad1Key = mNumpad1Key
End Property
Property Get Numpad2Key() As String
    Numpad2Key = mNumpad2Key
End Property
Property Get Numpad3Key() As String
    Numpad3Key = mNumpad3Key
End Property
Property Get Numpad4Key() As String
    Numpad4Key = mNumpad4Key
End Property
Property Get Numpad5Key() As String
    Numpad5Key = mNumpad5Key
End Property
Property Get Numpad6Key() As String
    Numpad6Key = mNumpad6Key
End Property
Property Get Numpad7Key() As String
    Numpad7Key = mNumpad7Key
End Property
Property Get Numpad8Key() As String
    Numpad8Key = mNumpad8Key
End Property
Property Get Numpad9Key() As String
    Numpad9Key = mNumpad9Key
End Property
Property Get MultiplyKey() As String
    MultiplyKey = mMultiplyKey
End Property
Property Get AddKey() As String
    AddKey = mAddKey
End Property
Property Get SeparatorKey() As String
    SeparatorKey = mSeparatorKey
End Property
Property Get SubtractKey() As String
    SubtractKey = mSubtractKey
End Property
Property Get DecimalKey() As String
    DecimalKey = mDecimalKey
End Property
Property Get DivideKey() As String
    DivideKey = mDivideKey
End Property
Property Get F1Key() As String
    F1Key = mF1Key
End Property
Property Get F2Key() As String
    F2Key = mF2Key
End Property
Property Get F3Key() As String
    F3Key = mF3Key
End Property
Property Get F4Key() As String
    F4Key = mF4Key
End Property
Property Get F5Key() As String
    F5Key = mF5Key
End Property
Property Get F6Key() As String
    F6Key = mF6Key
End Property
Property Get F7Key() As String
    F7Key = mF7Key
End Property
Property Get F8Key() As String
    F8Key = mF8Key
End Property
Property Get F9Key() As String
    F9Key = mF9Key
End Property
Property Get F10Key() As String
    F10Key = mF10Key
End Property
Property Get F11Key() As String
    F11Key = mF11Key
End Property
Property Get F12Key() As String
    F12Key = mF12Key
End Property
Property Get MetaKey() As String
    MetaKey = mMetaKey
End Property
Property Get CommandKey() As String
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
    mPage_UpKey = "\ue00e"
    mPage_DownKey = "\ue00f"
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
