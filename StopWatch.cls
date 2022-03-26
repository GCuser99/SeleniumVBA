VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2900
Attribute VB_Description = "A programmable timer object."
' ==========================================================================
' SeleniumVBA v0.0.6
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
'   Optional StopWatch class to test performance
'
'   Usage:
'
'   Dim StopWatch as new StopWatch
'   StopWatch.StartIt
'
'   'do some stuff here
'
'   Debug.Print StopWatch.StopIt 'elapse time since StartIt
'

Option Explicit

Private mStartTime As Single
Private mIsRunning As Boolean
Private mStopTime As Single

Public Sub StartIt()
Attribute StartIt.VB_Description = "Starts the timer."
Attribute StartIt.VB_HelpID = 2830
    'Store the starting time
    mStartTime = timer 'timeGetTime
    mIsRunning = True
End Sub

Property Get IsRunning() As Boolean
Attribute IsRunning.VB_Description = "Returns True if the StopWatch object is running."
Attribute IsRunning.VB_HelpID = 1950
    'Return True if the stopwatch is running,
    'False otherwise
    IsRunning = mIsRunning
End Property

Property Get elapsedTime() As Single
Attribute elapsedTime.VB_Description = "Returns the elapse time after the StopWatch has been stopped."
Attribute elapsedTime.VB_HelpID = 1650
    'Return the elapsed time.
    'This is meaningless unless we're running, of course
    If IsRunning Then
        elapsedTime = (timer - mStartTime)
    Else
        elapsedTime = (mStopTime - mStartTime)
    End If
End Property

Property Get startTime() As Single
Attribute startTime.VB_Description = "Returns the start time after the timer has been stopped."
Attribute startTime.VB_HelpID = 2840
    'Return the starting time
    startTime = mStartTime
End Property

Property Get StopTime() As Single
Attribute StopTime.VB_Description = "Returns the stop time after the timer has been stopped."
Attribute StopTime.VB_HelpID = 2890
    'Return the starting time
    If Not IsRunning Then
        StopTime = mStopTime
    Else
        StopTime = 0
    End If
End Property

Public Function StopIt() As Single
Attribute StopIt.VB_Description = "Stops the timer and returns the elapse time."
Attribute StopIt.VB_HelpID = 2880
    'Retrieve the elapsed time, and reset
    'the start time to 0 (so IsRunning will
    'think you're not currently running)
    If IsRunning Then
        mStopTime = timer
        StopIt = (mStopTime - mStartTime)
        mIsRunning = False
    Else
        StopIt = 0
    End If
End Function

