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
'This class was written to handle special keys for SendKeys and Action Chains
'
'For more info, see https://www.w3.org/TR/webdriver/#keyboard-actions

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
Private mCtrlKey As String
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
Private mZenkakuHankakuKey As String
Private mCtrlShiftKeys As String
Private mAltShiftKeys As String
Private mCtrlAltKeys As String

'@Description("Returns the Null (\\uE000) key - triggers release of all previous modifiers used in the string")
Public Property Get NullKey() As String
Attribute NullKey.VB_Description = "Returns the Null (\\uE000) key - triggers release of all previous modifiers used in the string"
    NullKey = mNullKey
End Property
'@Description("Returns the Cancel (\\uE001) key")
Public Property Get CancelKey() As String
Attribute CancelKey.VB_Description = "Returns the Cancel (\\uE001) key"
    CancelKey = mCancelKey
End Property
'@Description("Returns the Help (\\uE002) key")
Public Property Get HelpKey() As String
Attribute HelpKey.VB_Description = "Returns the Help (\\uE002) key"
    HelpKey = mHelpKey
End Property
'@Description("Returns the Backspace (\\uE003) key")
Public Property Get BackspaceKey() As String
Attribute BackspaceKey.VB_Description = "Returns the Backspace (\\uE003) key"
    BackspaceKey = mBackspaceKey
End Property
'@Description("Returns the Tab (\\uE004) key")
Public Property Get TabKey() As String
Attribute TabKey.VB_Description = "Returns the Tab (\\uE004) key"
    TabKey = mTabKey
End Property
'@Description("Returns the Clear (\\uE005) key")
Public Property Get ClearKey() As String
Attribute ClearKey.VB_Description = "Returns the Clear (\\uE005) key"
    ClearKey = mClearKey
End Property
'@Description("Returns the Return (\\uE006) key")
Public Property Get ReturnKey() As String
Attribute ReturnKey.VB_Description = "Returns the Return (\\uE006) key"
    ReturnKey = mReturnKey
End Property
'@Description("Returns the Enter (\\uE007) key")
Public Property Get EnterKey() As String
Attribute EnterKey.VB_Description = "Returns the Enter (\\uE007) key"
    EnterKey = mEnterKey
End Property
'@Description("Returns the Shift (\\uE008) key")
Public Property Get ShiftKey() As String
Attribute ShiftKey.VB_Description = "Returns the Shift (\\uE008) key"
    ShiftKey = mShiftKey
End Property
'@Description("Returns the Ctrl (\\uE009) key")
Public Property Get CtrlKey() As String
Attribute CtrlKey.VB_Description = "Returns the Ctrl (\\uE009) key"
    CtrlKey = mCtrlKey
End Property
'@Description("Returns the Alt (\\uE00a) key")
Public Property Get AltKey() As String
Attribute AltKey.VB_Description = "Returns the Alt (\\uE00a) key"
    AltKey = mAltKey
End Property
'@Description("Returns the Pause (\\uE00b) key")
Public Property Get PauseKey() As String
Attribute PauseKey.VB_Description = "Returns the Pause (\\uE00b) key"
    PauseKey = mPauseKey
End Property
'@Description("Returns the Escape (\\uE00c) key")
Public Property Get EscapeKey() As String
Attribute EscapeKey.VB_Description = "Returns the Escape (\\uE00c) key"
    EscapeKey = mEscapeKey
End Property
'@Description("Returns the Space (\\uE00d) key")
Public Property Get SpaceKey() As String
Attribute SpaceKey.VB_Description = "Returns the Space (\\uE00d) key"
    SpaceKey = mSpaceKey
End Property
'@Description("Returns the PageUp (\\uE00e) key")
Public Property Get PageUpKey() As String
Attribute PageUpKey.VB_Description = "Returns the PageUp (\\uE00e) key"
    PageUpKey = mPageUpKey
End Property
'@Description("Returns the PageDown (\\uE00f) key")
Public Property Get PageDownKey() As String
Attribute PageDownKey.VB_Description = "Returns the PageDown (\\uE00f) key"
    PageDownKey = mPageDownKey
End Property
'@Description("Returns the End (\\uE010) key")
Public Property Get EndKey() As String
Attribute EndKey.VB_Description = "Returns the End (\\uE010) key"
    EndKey = mEndKey
End Property
'@Description("Returns the Home (\\uE011) key")
Public Property Get HomeKey() As String
Attribute HomeKey.VB_Description = "Returns the Home (\\uE011) key"
    HomeKey = mHomeKey
End Property
'@Description("Returns the Left (\\uE012) key")
Public Property Get LeftKey() As String
Attribute LeftKey.VB_Description = "Returns the Left (\\uE012) key"
    LeftKey = mLeftKey
End Property
'@Description("Returns the Up (\\uE013) key")
Public Property Get UpKey() As String
Attribute UpKey.VB_Description = "Returns the Up (\\uE013) key"
    UpKey = mUpKey
End Property
'@Description("Returns the Right (\\uE014) key")
Public Property Get RightKey() As String
Attribute RightKey.VB_Description = "Returns the Right (\\uE014) key"
    RightKey = mRightKey
End Property
'@Description("Returns the Down (\\uE015) key")
Public Property Get DownKey() As String
Attribute DownKey.VB_Description = "Returns the Down (\\uE015) key"
    DownKey = mDownKey
End Property
'@Description("Returns the Insert (\\uE016) key")
Public Property Get InsertKey() As String
Attribute InsertKey.VB_Description = "Returns the Insert (\\uE016) key"
    InsertKey = mInsertKey
End Property
'@Description("Returns the Delete (\\uE017) key")
Public Property Get DeleteKey() As String
Attribute DeleteKey.VB_Description = "Returns the Delete (\\uE017) key"
    DeleteKey = mDeleteKey
End Property
'@Description("Returns the Semicolon (\\uE018) key")
Public Property Get SemicolonKey() As String
Attribute SemicolonKey.VB_Description = "Returns the Semicolon (\\uE018) key"
    SemicolonKey = mSemicolonKey
End Property
'@Description("Returns the Equals (\\uE019) key")
Public Property Get EqualsKey() As String
Attribute EqualsKey.VB_Description = "Returns the Equals (\\uE019) key"
    EqualsKey = mEqualsKey
End Property
'@Description("Returns the Numpad 0 (\\uE01a) key")
Public Property Get Numpad0Key() As String
Attribute Numpad0Key.VB_Description = "Returns the Numpad 0 (\\uE01a) key"
    Numpad0Key = mNumpad0Key
End Property
'@Description("Returns the Numpad 1 (\\uE01b) key")
Public Property Get Numpad1Key() As String
Attribute Numpad1Key.VB_Description = "Returns the Numpad 1 (\\uE01b) key"
    Numpad1Key = mNumpad1Key
End Property
'@Description("Returns the Numpad 2 (\\uE01c) key")
Public Property Get Numpad2Key() As String
Attribute Numpad2Key.VB_Description = "Returns the Numpad 2 (\\uE01c) key"
    Numpad2Key = mNumpad2Key
End Property
'@Description("Returns the Numpad 3 (\\uE01d) key")
Public Property Get Numpad3Key() As String
Attribute Numpad3Key.VB_Description = "Returns the Numpad 3 (\\uE01d) key"
    Numpad3Key = mNumpad3Key
End Property
'@Description("Returns the Numpad 4 (\\uE01e) key")
Public Property Get Numpad4Key() As String
Attribute Numpad4Key.VB_Description = "Returns the Numpad 4 (\\uE01e) key"
    Numpad4Key = mNumpad4Key
End Property
'@Description("Returns the Numpad 5 (\\uE01f) key")
Public Property Get Numpad5Key() As String
Attribute Numpad5Key.VB_Description = "Returns the Numpad 5 (\\uE01f) key"
    Numpad5Key = mNumpad5Key
End Property
'@Description("Returns the Numpad 6 (\\uE020) key")
Public Property Get Numpad6Key() As String
Attribute Numpad6Key.VB_Description = "Returns the Numpad 6 (\\uE020) key"
    Numpad6Key = mNumpad6Key
End Property
'@Description("Returns the Numpad 7 (\\uE021) key")
Public Property Get Numpad7Key() As String
Attribute Numpad7Key.VB_Description = "Returns the Numpad 7 (\\uE021) key"
    Numpad7Key = mNumpad7Key
End Property
'@Description("Returns the Numpad 8 (\\uE022) key")
Public Property Get Numpad8Key() As String
Attribute Numpad8Key.VB_Description = "Returns the Numpad 8 (\\uE022) key"
    Numpad8Key = mNumpad8Key
End Property
'@Description("Returns the Numpad 9 (\\uE023) key")
Public Property Get Numpad9Key() As String
Attribute Numpad9Key.VB_Description = "Returns the Numpad 9 (\\uE023) key"
    Numpad9Key = mNumpad9Key
End Property
'@Description("Returns the Multiply (\\uE024) key")
Public Property Get MultiplyKey() As String
Attribute MultiplyKey.VB_Description = "Returns the Multiply (\\uE024) key"
    MultiplyKey = mMultiplyKey
End Property
'@Description("Returns the Add (\\uE025) key")
Public Property Get AddKey() As String
Attribute AddKey.VB_Description = "Returns the Add (\\uE025) key"
    AddKey = mAddKey
End Property
'@Description("Returns the Separator (\\uE026) key")
Public Property Get SeparatorKey() As String
Attribute SeparatorKey.VB_Description = "Returns the Separator (\\uE026) key"
    SeparatorKey = mSeparatorKey
End Property
'@Description("Returns the Subtract (\\uE027) key")
Public Property Get SubtractKey() As String
Attribute SubtractKey.VB_Description = "Returns the Subtract (\\uE027) key"
    SubtractKey = mSubtractKey
End Property
'@Description("Returns the Decimal (\\uE028) key")
Public Property Get DecimalKey() As String
Attribute DecimalKey.VB_Description = "Returns the Decimal (\\uE028) key"
    DecimalKey = mDecimalKey
End Property
'@Description("Returns the Divide (\\uE029) key")
Public Property Get DivideKey() As String
Attribute DivideKey.VB_Description = "Returns the Divide (\\uE029) key"
    DivideKey = mDivideKey
End Property
'@Description("Returns the F1 (\\uE031) key")
Public Property Get F1Key() As String
Attribute F1Key.VB_Description = "Returns the F1 (\\uE031) key"
    F1Key = mF1Key
End Property
'@Description("Returns the F2 (\\uE032) key")
Public Property Get F2Key() As String
Attribute F2Key.VB_Description = "Returns the F2 (\\uE032) key"
    F2Key = mF2Key
End Property
'@Description("Returns the F3 (\\uE033) key")
Public Property Get F3Key() As String
Attribute F3Key.VB_Description = "Returns the F3 (\\uE033) key"
    F3Key = mF3Key
End Property
'@Description("Returns the F4 (\\uE034) key")
Public Property Get F4Key() As String
Attribute F4Key.VB_Description = "Returns the F4 (\\uE034) key"
    F4Key = mF4Key
End Property
'@Description("Returns the F5 (\\uE035) key")
Public Property Get F5Key() As String
Attribute F5Key.VB_Description = "Returns the F5 (\\uE035) key"
    F5Key = mF5Key
End Property
'@Description("Returns the F6 (\\uE036) key")
Public Property Get F6Key() As String
Attribute F6Key.VB_Description = "Returns the F6 (\\uE036) key"
    F6Key = mF6Key
End Property
'@Description("Returns the F7 (\\uE037) key")
Public Property Get F7Key() As String
Attribute F7Key.VB_Description = "Returns the F7 (\\uE037) key"
    F7Key = mF7Key
End Property
'@Description("Returns the F8 (\\uE038) key")
Public Property Get F8Key() As String
Attribute F8Key.VB_Description = "Returns the F8 (\\uE038) key"
    F8Key = mF8Key
End Property
'@Description("Returns the F9 (\\uE039) key")
Public Property Get F9Key() As String
Attribute F9Key.VB_Description = "Returns the F9 (\\uE039) key"
    F9Key = mF9Key
End Property
'@Description("Returns the F10 (\\uE03a) key")
Public Property Get F10Key() As String
Attribute F10Key.VB_Description = "Returns the F10 (\\uE03a) key"
    F10Key = mF10Key
End Property
'@Description("Returns the F11 (\\uE03b) key")
Public Property Get F11Key() As String
Attribute F11Key.VB_Description = "Returns the F11 (\\uE03b) key"
    F11Key = mF11Key
End Property
'@Description("Returns the F12 (\\uE03c) key")
Public Property Get F12Key() As String
Attribute F12Key.VB_Description = "Returns the F12 (\\uE03c) key"
    F12Key = mF12Key
End Property
'@Description("Returns the Meta (\\uE03d) key")
Public Property Get MetaKey() As String
Attribute MetaKey.VB_Description = "Returns the Meta (\\uE03d) key"
    MetaKey = mMetaKey
End Property
'@Description("Returns the Command (\\uE03d) key")
Public Property Get CommandKey() As String
Attribute CommandKey.VB_Description = "Returns the Command (\\uE03d) key"
    CommandKey = mCommandKey
End Property
'@Description("Returns the Zenkaku to Hankaku (\\uE040) key for Japanese keyboard")
Public Property Get ZenkakuHankakuKey() As String
Attribute ZenkakuHankakuKey.VB_Description = "Returns the Zenkaku to Hankaku (\\uE040) key for Japanese keyboard"
    ZenkakuHankakuKey = mZenkakuHankakuKey
End Property
'@Description("Returns the CtrlShift (\\uE009 + \\uE008) keys")
Public Property Get CtrlShiftKeys() As String
Attribute CtrlShiftKeys.VB_Description = "Returns the CtrlShift (\\uE009 + \\uE008) keys"
    CtrlShiftKeys = mCtrlShiftKeys
End Property
'@Description("Returns the AltShift (\\uE00a + \\uE008) keys")
Public Property Get AltShiftKeys() As String
Attribute AltShiftKeys.VB_Description = "Returns the AltShift (\\uE00a + \\uE008) keys"
    AltShiftKeys = mCtrlShiftKeys
End Property
'@Description("Returns the CtrlAlt (\\uE009 + \\uE00a) keys")
Public Property Get CtrlAltKeys() As String
Attribute CtrlAltKeys.VB_Description = "Returns the CtrlAlt (\\uE009 + \\uE00a) keys"
    CtrlAltKeys = mCtrlAltKeys
End Property

'@Description("Returns modifier key(s) down, keys to send, modifier key(s) up sequence")
Public Function Chord(ByVal modifierKeys As String, ByVal keysToSend As String) As String
Attribute Chord.VB_Description = "Returns modifier key(s) down, keys to send, modifier key(s) up sequence"
    modifierKeys = unEscapeUnicode(modifierKeys)
    'check if modifierKeys has only valid modifiers Shift, Ctrl, and/or Alt
    If Not (modifierKeys Like "*[!" & ChrW$(&HE008&) & "-" & ChrW$(&HE00A&) & "]*") Then
        'combine modifiers and keys with a terminating null key
        Chord = modifierKeys & unEscapeUnicode(keysToSend) & NullKey
    Else
        Err.Raise 1, , "Error in Chord method - valid modifier keys must be Shift, Control, and Alt keys"
    End If
End Function

'@Description("Returns a string consisting of a sub-string repeated the specified number of times")
Public Function Repeat(ByVal keyString As String, ByVal numberOfTimes As Long) As String
Attribute Repeat.VB_Description = "Returns a string consisting of a sub-string repeated the specified number of times"
    Dim i As Long
    keyString = unEscapeUnicode(keyString)
    For i = 1 To numberOfTimes
        Repeat = Repeat & keyString
    Next i
End Function

Private Sub Class_Initialize()
    mNullKey = ChrW$(&HE000&)
    mCancelKey = ChrW$(&HE001&)
    mHelpKey = ChrW$(&HE002&)
    mBackspaceKey = ChrW$(&HE003&)
    mTabKey = ChrW$(&HE004&)
    mClearKey = ChrW$(&HE005&)
    mReturnKey = ChrW$(&HE006&)
    mEnterKey = ChrW$(&HE007&)
    mShiftKey = ChrW$(&HE008&)
    mControlKey = ChrW$(&HE009&)
    mCtrlKey = ChrW$(&HE009&)
    mAltKey = ChrW$(&HE00A&)
    mPauseKey = ChrW$(&HE00B&)
    mEscapeKey = ChrW$(&HE00C&)
    mSpaceKey = ChrW$(&HE00D&)
    mPageUpKey = ChrW$(&HE00E&)
    mPageDownKey = ChrW$(&HE00F&)
    mEndKey = ChrW$(&HE010&)
    mHomeKey = ChrW$(&HE011&)
    mLeftKey = ChrW$(&HE012&)
    mUpKey = ChrW$(&HE013&)
    mRightKey = ChrW$(&HE014&)
    mDownKey = ChrW$(&HE015&)
    mInsertKey = ChrW$(&HE016&)
    mDeleteKey = ChrW$(&HE017&)
    mSemicolonKey = ChrW$(&HE018&)
    mEqualsKey = ChrW$(&HE019&)
    mNumpad0Key = ChrW$(&HE01A&)
    mNumpad1Key = ChrW$(&HE01B&)
    mNumpad2Key = ChrW$(&HE01C&)
    mNumpad3Key = ChrW$(&HE01D&)
    mNumpad4Key = ChrW$(&HE01E&)
    mNumpad5Key = ChrW$(&HE01F&)
    mNumpad6Key = ChrW$(&HE020&)
    mNumpad7Key = ChrW$(&HE021&)
    mNumpad8Key = ChrW$(&HE022&)
    mNumpad9Key = ChrW$(&HE023&)
    mMultiplyKey = ChrW$(&HE024&)
    mAddKey = ChrW$(&HE025&)
    mSeparatorKey = ChrW$(&HE026&)
    mSubtractKey = ChrW$(&HE027&)
    mDecimalKey = ChrW$(&HE028&)
    mDivideKey = ChrW$(&HE029&)
    mF1Key = ChrW$(&HE031&)
    mF2Key = ChrW$(&HE032&)
    mF3Key = ChrW$(&HE033&)
    mF4Key = ChrW$(&HE034&)
    mF5Key = ChrW$(&HE035&)
    mF6Key = ChrW$(&HE036&)
    mF7Key = ChrW$(&HE037&)
    mF8Key = ChrW$(&HE038&)
    mF9Key = ChrW$(&HE039&)
    mF10Key = ChrW$(&HE03A&)
    mF11Key = ChrW$(&HE03B&)
    mF12Key = ChrW$(&HE03C&)
    mMetaKey = ChrW$(&HE03D&)
    mCommandKey = ChrW$(&HE03D&)
    mZenkakuHankakuKey = ChrW$(&HE040&)
    mCtrlShiftKeys = ChrW$(&HE009&) & ChrW$(&HE008&)
    mAltShiftKeys = ChrW$(&HE00A&) & ChrW$(&HE008&)
    mCtrlAltKeys = ChrW$(&HE009&) & ChrW$(&HE00A&)
End Sub
