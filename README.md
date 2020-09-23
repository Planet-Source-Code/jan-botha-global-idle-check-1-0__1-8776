<div align="center">

## Global Idle Check 1\.0


</div>

### Description

Global Idle Check is used to monitor your system for activity. If found that your system has been inactive (idling) for the amount of time you specify, the code will call the IdleStateEngaged sub, which is where you can put your code.

The code demonstrates the use of the GetAsyncKeyState as well as the GetCursorPos API's.

The code continually monitors the state of your keyboard and mouse buttons, as well as you mouse position.

This submission is commented quite heavily, and I hope it is easy to follow. Please vote for me if you think I deserve it. Give me comments as well, please.

If you've always wanted to create a screen saver that is independent of windows, you can do so now.
 
### More Info
 
INTERVAL - set the number of seconds the system

must be inactive before the idle-state is reached.

BEFORE you start:

Put 2 Timers on a form. Name the one tmrPeriod

and the other tmrStateMonitor. Set both

Timers' Interval property to 1.

This will screw up if you run it at midnight, because the Timer object resets to 0 then.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Botha](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-botha.md)
**Level**          |Advanced
**User Rating**    |4.6 (51 globes from 11 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-botha-global-idle-check-1-0__1-8776/archive/master.zip)

### API Declarations

```
'I've put them in the code itself!
```


### Source Code

```
'BEFORE you start:
'Put 2 Timers on a form
'Name the one tmrPeriod and the other
'tmrStateMonitor. Set both Timers' Interval
'property to 1
'Paste all the below code into the form!
'Global Idle Check 1
'==============
'Copyright > Jan Botha 1998-2000
'Release Date > 9 June 2000
'Email > ja_botha@hotmail.com
'
'This code monitors the state of the keys on the
'keyboard and the mousebuttons as well as the
'position of the mouse. Whenever the 'tmrStateMonitor'
'finds that no keys or mousebuttons is pressed
'and that the mouse is still in the same position,
'it sets the IsIdle variable to True and the
'startOfIdle variable to the = the system timer.
'
'Throughout the form, comments/documentation
'are given either on the same line as the statement
'it is commenting on, or on the line preceeding the
'statement. The code is quite well commented
'to make beginners or any one else understand
'what's going on. This code IS on a beginner level,
'but the result is quite useful.
'
'Contact me if you have any ideas.
'You can use and modify this as much as you like,
'BUT:
'1. Please let me know how you modified this, just
'  'cause I'd like to see where I maybe went wrong.
'2. Give me some credit. Even if you only tell me
'  about an app that you used this for!
'
'Now I'll shut up, so you can actually see what this
'is about.
'Enjoy!
'Jan Botha
'email: ja_botha@hotmail.com
'==========================
'IMPORTANT NOTE:
'This code will probably screw up completely if you
'try to run it while the midnight rollover occurr.
'That's 'cause the Timer object resets to 0 at
'midnight.
'You could try and run something to wait until midnight
'has passed, before continuing the idle check
'=============================
'START OF ACTUAL CODE:
'all variables must be declared explicitly (this is simply
'a good programming "principle", if you want :-)
Option Explicit
'type declaration for the mouse (cursor) position
Private Type POINTAPI
    x As Long
    y As Long
End Type
'API function to get the cursor position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'API function to check the state of the mouse buttons
'as well as the keyboard.
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'set the length of time the computer must idle, before
'the so-called "idle-state" is reached. unit: seconds
'You'd probably want to change this value!!!
Private Const INTERVAL As Long = 10
Dim IsIdle As Boolean 'True when idling or while in idle-state
Dim MousePos As POINTAPI 'holds mouse position
'holds time (in seconds) when the idle started
'used to calculate if the computer has been idle for INTERVAL
Dim startOfIdle As Long
Private Sub tmrStateMonitor_Timer()
  'holds the state of the key that is being monitored
  Dim state As Integer
  'holds the CURRENT mouse position.
  'It's to compare the current position with the previous
  'position
  Dim tmpPos As POINTAPI
  Dim ret As Long 'simply holds the return value of the API
  'this checks if a key/button is pressed, or
  'if the mouse has moved.
  Dim IdleFound As Boolean
  Dim i As Integer 'the counter uses by the For loop
  IdleFound = False
  'Here I'm not sure about myself:
  'I don't know to what to set the value
  '256 to. It works as is, though!
  'And, what it does, is retrieve the state of each
  'individual key.
  For i = 1 To 256
    'call the API
    state = GetAsyncKeyState(i)
    'state will = -32767 if the 'i' key/button is
    'currently being pressed:
    If state = -32767 Then
      'if it is pressed, then this is the end of any idles
      IdleFound = True 'means that something is withholding the computer of idling
      IsIdle = False 'thus, it is not idling, so set the value
    End If
  Next
  'get the position of the mouse cursor
  ret = GetCursorPos(tmpPos)
  'if the coordinates of the mouse are different than
  'last time or when the idle started, then the system
  'is not idling:
  If tmpPos.x <> MousePos.x Or tmpPos.y <> MousePos.y Then
    IsIdle = False 'set the...
    IdleFound = True 'values
    'store the current coordinates so that we
    'can compare next time round
    MousePos.x = tmpPos.x
    MousePos.y = tmpPos.y
  End If
  'if something did not withhold the idle then...
  If Not IdleFound Then
    'if isIdle not equals false, then don't reset the
    'startOfIdle!!
    If Not IsIdle Then
      'if it is false, then the idle is beginning
      IsIdle = True
      startOfIdle = Timer
    End If
  End If
End Sub
Private Sub tmrPeriod_Timer()
  'this timer continuesly monitors the
  'value of IsIdle to see if the system has been
  'idle for INTERVAL
  If IsIdle Then
    'if the difference between now (timer) and the
    'time the idle started, is => INTERVAL, then
    'the 'idle state' has been reached
    If Timer - startOfIdle >= INTERVAL Then
      'call the sub that will handle any code at this stage
      'this is merely to seperate the idle check code
      'from your own code
      'NOTE: I advise you to perform some sort of
      'check here to see if the idle state has been
      'reached for the first time, or if the system
      'has just been idling ever since the idle state
      'was reached
      Call IdleStateEngaged(Timer)
      'important: set the values
      startOfIdle = Timer
      IsIdle = True
    End If
   Else ' not idling, or the idlestate has been left
    'call the sub
    'NOTE: I advise you to perform some sort of
    'check here to see if the system was in the
    'idle state, or if the system
    'has not been idling anyway
    Call IdleStateDisengaged(Timer)
  End If
End Sub
Public Sub IdleStateEngaged(ByVal IdleStartTime As Long)
  'PUT YOUR CODE HERE:
  'This is where you will put the code that you want
  'to execute now that the system has been idling
  'for INTERVAL seconds
  'Example:
  Caption = "Idle state reached - " & IdleStartTime
  'If you use the Global Idle Check for a screen
  'saver (thereby overruling the window$ screensaver),
  'you would put the start code here
End Sub
Public Sub IdleStateDisengaged(ByVal IdleStopTime As Long)
  'PUT YOUR CODE HERE:
  'This is where you will put the code that you want
  'to execute now as soon as the system stops idling
  'or while the user is active
  'Example:
  Caption = "No idling - " & IdleStopTime
  'If you use the Global Idle Check for a screen
  'saver (thereby overruling the window$ screensaver),
  'you would put the end code here
End Sub
```

