Attribute VB_Name = "j0k4w1ld"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd&)
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function Paintdesktop Lib "user32" Alias "PaintDesktop" (ByVal hdc As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const SMTO_ABORTIFHUNG = &H2

' Registry - Used to get and set registry.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number

' Used to get the Directories in Registry
Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Function Loader() 'loads your program to chat or pm
Call Set_SYNCaption

YSend "j0k4w1lds Module"
Pause 1
ClickSend
YSend "Get it @"
Pause 1
ClickSend
YSend "www.Klownin-Inc.com"
Pause 1
ClickSend
YSend "for more stuff  " & Time
Pause 1
ClickSend
YSend "If U AiNt A KlOwN U AiNt DoWn MoFo !!  " & Date
Pause 1
ClickSend
YSend "http://Klownin-Inc.com/ Klownin Inc.</url>"
Pause 1
ClickSend
YSend "By: Marx ( j0k4w1ld )"
Pause 1
ClickSend
End Function

Function Unloader() ' unloads ur program on chat or pm
Call Yahoo_Caption

YSend "j0k4w1lds Module"
Pause 1
ClickSend
YSend "Get it @"
Pause 1
ClickSend
YSend "www.Klownin-Inc.com"
Pause 1
ClickSend
YSend "for more stuff  " & Time
Pause 1
ClickSend
YSend "If U AiNt A KlOwN U AiNt DoWn MoFo !!  " & Date
Pause 1
ClickSend
YSend "http://Klownin-Inc.com/ Klownin Inc.</url>"
Pause 1
ClickSend
YSend "By: Marx ( j0k4w1ld )"
Pause 1
ClickSend
End Function

Function YSend(txt As String)
Dim imclass As Long, richedit As Long
imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SendMessageByString(richedit, WM_SETTEXT, 0&, txt$)
End Function




Sub ChangeStatus(Stat As String) ' changes your status to anything u want to change it lol
Dim name As String
name = GetString(HKEY_CURRENT_USER, "Software\Yahoo\Pager", "Yahoo! user id")

Call SaveString(HKEY_CURRENT_USER, "Software\Yahoo\Pager\profiles\" + name + "\Custom Msgs", 1, Stat)
Dim x As Long
On Error Resume Next
x = FindWindow("YahooBuddyMain", vbNullString)
SendMessageLong x, &H111, 388, 1&
End Sub

Public Function DeleteKey(ByVal hKey As HKeyTypes, ByVal strPath As String)
    'Call DeleteKey(HKEY_CURRENT_USER, "Software\VBW\Registry")
    
    Dim keyhand As Long
    r = RegDeleteKey(hKey, strPath)
End Function

Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Public Sub SaveKey(hKey As HKeyTypes, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strData As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
End Sub

Public Function GetString(hKey As HKeyTypes, strPath As String, strValue As String)
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Function IsYahoo_Open()
   Dim yahoobuddymain As Long
 yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
    'U can keep any of the parameters as null string.
    ' Check if you were able to obtain the Window handle.

    If yahoobuddymain <> 0 Then
    MsgBox "Yahoo is open.", vbInformation + vbOKOnly, "SYN Build 1.5"
    
    Else
    
    MsgBox "Yahoo is not open.", vbInformation + vbOKOnly, "SYN Build 1.5"
    End If
   End Function

Function Yahoo_hWnd()
Dim yahoobuddymain As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
End Function
Function Y_PMhWnd()
Y_PMhWnd = FindWindow("imclass", vbNullString)
End Function

Sub ShrinkForm(frmObj As Form, Optional FramesPerSec As Long = 600, Optional UnloadForm As Boolean = False)
Dim vWidth As Long, vHeight As Long, count As Long

vWidth = frmObj.Width
vHeight = frmObj.Height

With frmObj
    For count = vWidth To 1 Step -1 * FramesPerSec
        frmObj.Move (Screen.Width - count) / 2, (Screen.Height - (vHeight * count / vWidth)) / 2, count, vHeight * count / vWidth
        frmObj.Refresh
        DoEvents
    Next
    
    frmObj.Hide
    If UnloadForm Then Unload frmObj
End With
End Sub


Public Function Open_Pm(User As String)
' Opens up your browser and does a code to send a im to the USER String
If ShellExecute(&O0, "Open", "ymsgr:sendIM?" & User$, vbNullString, vbNullString, SW_NORMAL) < 33 Then
End If
End Function
Public Function OpenPm(User As String)
' Opens up your browser and does a code to send a im to the USER String
If ShellExecute(&O0, "Open", "ymsgr:sendIM?" & User$, vbNullString, vbNullString, SW_NORMAL) < 33 Then
End If
End Function

Function SoundBOMBER(User As String)
On Error Resume Next ' If a error occurs, goto the next line
Open_Pm User$ ' Opens a pm to what ever the user string is
Pause 0.37    ' Pauses before sending the first string
YSend "<snd=pow>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=cowbell>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=chimeup>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=backsp>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=phone>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=sent>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=plybktsp>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=yahoomail>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=yahoo>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=knock>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=door>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
End Function
Function SoundBOMBER2()
On Error Resume Next
YSend "<snd=pow>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=cowbell>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=chimeup>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=backsp>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=phone>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=sent>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=plybktsp>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=yahoomail>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=yahoo>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=knock>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
YSend "<snd=door>"
Pause 0.37
ClickSend
Close_Maximum
Pause 0.37
End Function

Function Chat_or_PM()
Dim imclass As Long, Button As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If Right(TheText, 7) = "-- Chat" Then
Chat_or_PM = "chat"
ElseIf Right(TheText, 18) = "-- Instant Message" Then
Chat_or_PM = "pm"
ElseIf Left(TheText, 10) = "Conference" Then
Chat_or_PM = "chat"
ElseIf Left(TheText, 10) = "Voice Conf" Then
Chat_or_PM = "chat"
Else
imclass = FindWindow("imclass", vbNullString)
Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
TL = SendMessageLong(Button, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(Button, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If Not TheText = "" Then
Chat_or_PM = "pm"
Else
Chat_or_PM = "none"
End If
End If
End Function

Function Close_Yahoo()
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
Do
    DoEvents
       
   Loop Until imclass <= 0
End Function


Sub RunMenubystring(Window, mnuCap)
Dim ToSearch As Long
Dim MenuCount As Integer
Dim FindString
Dim ToSearchSub As Long
Dim MenuItemCount As Integer
Dim GetString
Dim SubCount As Long
Dim MenuString As String                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             'CBM RULES BITCH
Dim GetStringMenu As Integer
Dim MenuItem As Long
Dim RunTheMenu As Integer


ToSearch& = GetMenu(Window)
MenuCount% = GetMenuItemCount(ToSearch&)

For FindString = 0 To MenuCount% - 1
ToSearchSub& = GetSubMenu(ToSearch&, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub&)
For GetString = 0 To MenuItemCount% - 1
SubCount& = GetMenuItemID(ToSearchSub&, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub&, SubCount&, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(mnuCap)) Then
MenuItem& = SubCount&
GoTo MatchString
End If
Next GetString
Next FindString
MatchString:
RunTheMenu% = SendMessage(Window, WM_COMMAND, MenuItem&, 0)
End Sub

Function VoiceBomber(who As String)
If Chat_or_PM = "chat" Then Exit Function
Open_Pm who$
Call RunMenubystring(Y_PMhWnd, "Enable &Voice")
End Function

Public Function Close_Maximum()
Dim MessageBox As Long

MessageBox& = FindWindow("#32770", "Maximum message rate exceeded")
If MessageBox& Then
    Call SetForegroundWindow(MessageBox&)
    SendKeys "{enter}"
End If
End Function

Sub Kill_ChatAdd()
'this will change that chat ad on YIM
SaveString HKEY_CURRENT_USER, "Software\Yahoo\Pager\yurl", "Chat Adurl", "about:<body bgcolor=#FFFFFF></a>"
End Sub

Function Kill_CamAdd()
SaveString HKEY_CURRENT_USER, "Software\Yahoo\Pager\yurl", "Cam Adurl", "about:<body bgcolor=#FFFFFF></a>"
End Function

Function Kill_ConfAdd()
SaveString HKEY_CURRENT_USER, "Software\Yahoo\Pager\yurl", "Conf Adurl", "about:<body bgcolor=#FFFFFF></a>"
End Function

Public Function Status_Blue(txt As String)
ChangeStatus txt & Chr(160) & "!"
End Function

Public Function Set_SYNCaption()
Dim yahoobuddymain As Long, ysearchbar As Long, ysearchbox As Long, Button As Long
Dim editx As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
ysearchbar = FindWindowEx(yahoobuddymain, 0&, "ysearchbar", vbNullString)
ysearchbox = FindWindowEx(ysearchbar, 0&, "ysearchbox", vbNullString)
editx = FindWindowEx(ysearchbox, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, "Ø ø  S ¥ N   ¹·5  ø Ø")
Button = FindWindowEx(ysearchbar, 0&, "button", vbNullString)
Call SendMessageByString(Button, WM_SETTEXT, 0&, "S ¥ N")
Call SendMessageByString(yahoobuddymain, WM_SETTEXT, 0&, "Ø ø  S ¥ N   ¹·5  ø Ø")
Call EnableWindow(ysearchbar, 0)
Call EnableWindow(ysearchbox, 0)
Call EnableWindow(Button, 0)
Call EnableWindow(editx, 0)
End Function

Public Function Yahoo_Caption()
Dim yahoobuddymain As Long, ysearchbar As Long, ysearchbox As Long, Button As Long
Dim editx As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
ysearchbar = FindWindowEx(yahoobuddymain, 0&, "ysearchbar", vbNullString)
ysearchbox = FindWindowEx(ysearchbar, 0&, "ysearchbox", vbNullString)
editx = FindWindowEx(ysearchbox, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, "Search Yahoo!")
Button = FindWindowEx(ysearchbar, 0&, "button", vbNullString)
Call SendMessageByString(Button, WM_SETTEXT, 0&, "Search")
Call SendMessageByString(yahoobuddymain, WM_SETTEXT, 0&, "Yahoo! Messenger")
Call EnableWindow(ysearchbar, 1)
Call EnableWindow(ysearchbox, 1)
Call EnableWindow(Button, 1)
Call EnableWindow(editx, 1)
End Function

Public Sub scroll_shouts(text As TextBox)
Dim a As String
a = text
If Len(a) < 1 Then Exit Sub
a = Right(a, Len(a) - 1) & Left(a, 1)
text = a
End Sub
Public Sub scroll_shout1(label As label)
Dim a As String
a = label.Caption
If Len(a) < 1 Then Exit Sub
a = Right(a, Len(a) - 1) & Left(a, 1)
label.Caption = a
End Sub


Public Sub Answering_Machine(List As listbox, Message As String, Lbl As label)
If Y_PMhWnd = 0 Then Exit Sub
If Chat_or_PM = "pm" Then
List.AddItem GetName
Lbl.Caption = Lbl.Caption + 1
YSend Message
Y_Close
End If
End Sub

Sub Y_Close()
Dim imclass As Long
     imclass = FindWindow("imclass", vbNullString)
Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
End Sub



Public Sub Emote()
On Error Resume Next
Dim LIndex As Long, imclass As Long, listbox As Long

imclass = FindWindow("imclass", vbNullString)
listbox = FindWindowEx(imclass, 0&, "listbox", vbNullString)

Y_Emote_Count = SendMessageLong(listbox, LB_GETCOUNT, 0&, 0&)

LIndex = Int(Rnd * Y_Emote_Count)

Call SendMessageLong(listbox, LB_SETCURSEL, LIndex, 0&)
Call SendMessageLong(listbox, WM_LBUTTONDBLCLK, 0&, 0&)
DoEvents
End Sub
Sub GotoSite(url As String)
On Error GoTo Error
If Left(url, 4) = "www." Then url = "http://" + url
Shell ("explorer.exe " + url), vbNormalFocus
Exit Sub
Error:
Beep
Exit Sub
End Sub
Sub BuzzBomb(who As String)
'make shure you have a stopprog botton :P
If LCase(who) = "T0n3" Then
End
End If
Do
Open_Pm who
Pause 0.01
SendChat ("a/s/l")
Pause 0.01
ClickPmMenu "&Buzz Friend"
Pause 0.01
ClosePm
Pause 0.2
Loop
End Sub
Sub CenterForm(frm As Form)
If frm.WindowState = 0 Then
frm.Top = Screen.Height / 2 - frm.Height / 2
frm.Left = Screen.Width / 2 - frm.Width / 2
End If
End Sub
Sub ChageChatName(text As String)
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Call SendMessageByString(imclass, WM_SETTEXT, 0&, text)
End Sub

Sub ChangeHandsFree(text As String)
Dim atlcf As Long, x As Long, Button As Long
atlcf = FindWindow("atl:0054cf78", vbNullString)
x = FindWindowEx(atlcf, 0&, "#32770", vbNullString)
Button = FindWindowEx(x, 0&, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Call SendMessageByString(Button, WM_SETTEXT, 0&, text)
End Sub
Sub ChangeSearch(text As String)
Dim yahoobuddymain As Long, ysearchbar As Long, ysearchbox As Long
Dim editx As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
ysearchbar = FindWindowEx(yahoobuddymain, 0&, "ysearchbar", vbNullString)
ysearchbox = FindWindowEx(ysearchbar, 0&, "ysearchbox", vbNullString)
editx = FindWindowEx(ysearchbox, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, text)
End Sub
Sub ChangeSearchBotton(text As String)
Dim yahoobuddymain As Long, ysearchbar As Long, Button As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
ysearchbar = FindWindowEx(yahoobuddymain, 0&, "ysearchbar", vbNullString)
Button = FindWindowEx(ysearchbar, 0&, "button", vbNullString)
Call SendMessageByString(Button, WM_SETTEXT, 0&, text)
End Sub
Sub ChangeTalk(text As String)
Dim atlcf As Long, x As Long, Button As Long
atlcf = FindWindow("atl:0054cf78", vbNullString)
x = FindWindowEx(atlcf, 0&, "#32770", vbNullString)
Button = FindWindowEx(x, 0&, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Call SendMessageByString(Button, WM_SETTEXT, 0&, text)
End Sub

Sub ChangeYStatus(text As String)
Dim yahoobuddymain As Long, msctlsstatusbar As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
msctlsstatusbar = FindWindowEx(yahoobuddymain, 0&, "msctls_statusbar32", vbNullString)
Call SendMessageByString(msctlsstatusbar, WM_SETTEXT, 0&, text)
End Sub
Sub ChatInvite(who As String)
SendChat "/invite " + who
End Sub
Sub ChangeYMSGName(text As String)
Dim yahoobuddymain As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Call SendMessageByString(yahoobuddymain, WM_SETTEXT, 0&, text)
End Sub
Sub ClickChatMenu(TextToClick As String)
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = imclass
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase(TextToClick)) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&

End Sub
Sub ClickPmMenu(TextToClick As String)
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = imclass
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase(TextToClick)) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&

End Sub
Sub ClickSend()
Dim imclass As Long, Button As Long
imclass = FindWindow("imclass", vbNullString)
Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub ClickYMSGMenu(TextToClick As String)
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = imclass
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase(TextToClick)) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&

End Sub
Sub CloseAddFirend()
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
End Sub
Sub CloseChatInvite()
Dim yalertclass As Long
yalertclass = FindWindow("yalertclass", vbNullString)
Call SendMessageLong(yalertclass, WM_CLOSE, 0&, 0&)
End Sub
Sub CloseConInvite()

Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
End Sub
Sub CloseFile()
Dim yalertclass As Long
yalertclass = FindWindow("yalertclass", vbNullString)
Call SendMessageLong(yalertclass, WM_CLOSE, 0&, 0&)
End Sub

Sub ClosePm()
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Call SendMessageLong(imclass, WM_CLOSE, 0&, 0&)
End Sub
Sub CloseSharedFiles()
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
End Sub
Sub CloseVoice()
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
Call ClosePm
End Sub
Sub CloseWebcam()
Dim webcamclass As Long
webcamclass = FindWindow("webcamclass", vbNullString)
Call SendMessageLong(webcamclass, WM_CLOSE, 0&, 0&)
End Sub
Sub CloseWebcamDecline()
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
End Sub
Sub CloseYahooBomb()
'BOMBERS START THIS WORKS ON YOUR/THERE YAHOO MESSAGER HEHE:>
'this refrsh bombers there buddy list :))
'to stop just put pause 1E+271
Do
ClickYMSGMenu ("Close")
Pause 0.009
Loop
End Sub
Sub Copy(TextToCopy As String)
Clipboard.Clear
Clipboard.SetText TextToCopy
End Sub
Sub TheDate()
'just put text1.text = date
'or something like that
End Sub
Sub TheTime2()
'just put text1.text = time
'or something like that
End Sub
Function FilePath()
Dim light As String
If Right(App.Path, 1) = "\" Then
light = App.Path
Else
light = App.Path + "\"
FilePath = light
End If
End Function
Function FindPmBuddy()
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TL + 1, TheText)
FindPmBuddy = Left(TheText, TL)
End Function
Sub Form_ExitRight(Form As Form)
'Makes your form fly right
Do Until Form.Left >= 13000
Form.Left = Trim(Str(Int(Form.Left) + 25))
Loop
End Sub
Sub FriendTxt(text As String)
Dim x As Long, editx As Long
x = FindWindow("#32770", vbNullString)
editx = FindWindowEx(x, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, text)
End Sub
Sub GetFont()
Dim TheText As String, TL As Long
Dim imclass As Long
Dim toolbarparent As Long
Dim toolbarwindow As Long
Dim combobox As Long

imclass = FindWindow("imclass", vbNullString)
toolbarparent = FindWindowEx(imclass, 0&, "toolbarparent", vbNullString)
toolbarwindow = FindWindowEx(toolbarparent, 0&, "toolbarwindow32", vbNullString)
combobox = FindWindowEx(toolbarwindow, 0&, "combobox", vbNullString)
TL = SendMessageLong(combobox&, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(combobox&, WM_GETTEXT, TL + 1, TheText)
Y_FontName = Left(TheText, TL)

If combobox = 0 Then
    Exit Sub
End If

Do
    DoEvents
    imclass = FindWindow("imclass", vbNullString)
    toolbarparent = FindWindowEx(imclass, 0&, "toolbarparent", vbNullString)
    toolbarwindow = FindWindowEx(toolbarparent, 0&, "toolbarwindow32", vbNullString)
    combobox = FindWindowEx(toolbarwindow, 0&, "combobox", vbNullString)
    combobox = FindWindowEx(toolbarwindow, combobox, "combobox", vbNullString)
    TL = SendMessageLong(combobox&, WM_GETTEXTLENGTH, 0&, 0&)
    TheText = String(TL + 1, " ")
    Call SendMessageByString(combobox&, WM_GETTEXT, TL + 1, TheText)
    Y_FontSize = Left(TheText, TL)
Loop Until combobox <> 0

End Sub
Sub HideChat()
'this hides the chat room lol
Dim imclass As Long

imclass = FindWindow("imclass", vbNullString)
Call ShowWindow(imclass, SW_HIDE)
End Sub
Sub HideChatBanner()
'This hides the chat banner
Dim imclass As Long, atlb As Long, atlaxwin As Long
Dim shellembedding As Long, shelldocobjectview As Long, internetexplorerserver As Long
imclass = FindWindow("imclass", vbNullString)
atlb = FindWindowEx(imclass, 0&, "atl:0054b360", vbNullString)
atlaxwin = FindWindowEx(atlb, 0&, "atlaxwin7", vbNullString)
shellembedding = FindWindowEx(atlaxwin, 0&, "shell embedding", vbNullString)
shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
internetexplorerserver = FindWindowEx(shelldocobjectview, 0&, "internet explorer_server", vbNullString)
Call SendMessageLong(internetexplorerserver, WM_CLOSE, 0&, 0&)
End Sub
Sub HideCtrlAltDel()
App.TaskVisible = False
End Sub
Sub HideMessegeText()
Dim imclass As Long, richedit As Long
imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call ShowWindow(richedit, SW_HIDE)
End Sub
Sub HideReprot()
'This hides the Reprot Abuse in the chat room hehe
Dim imclass As Long, atleb As Long
imclass = FindWindow("imclass", vbNullString)
atleb = FindWindowEx(imclass, 0&, "atl:0054e0b8", vbNullString)
atleb = FindWindowEx(imclass, atleb, "atl:0054e0b8", vbNullString)
atleb = FindWindowEx(imclass, atleb, "atl:0054e0b8", vbNullString)
Call ShowWindow(atleb, SW_HIDE)
End Sub
Sub HideVoiceChatBanner()
'This hides the voice chat banner
Dim imclass As Long, atlb As Long, atlaxwin As Long
Dim shellembedding As Long, shelldocobjectview As Long, internetexplorerserver As Long
imclass = FindWindow("imclass", vbNullString)
atlb = FindWindowEx(imclass, 0&, "atl:0054b360", vbNullString)
atlaxwin = FindWindowEx(atlb, 0&, "atlaxwin7", vbNullString)
shellembedding = FindWindowEx(atlaxwin, 0&, "shell embedding", vbNullString)
shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
internetexplorerserver = FindWindowEx(shelldocobjectview, 0&, "internet explorer_server", vbNullString)
Call SendMessageLong(internetexplorerserver, WM_CLOSE, 0&, 0&)
End Sub
Sub HitEnterKey()
Dim imclass As Long, richedit As Long
imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SendMessageLong(richedit, WM_CHAR, 13, 0&)
End Sub
Sub IgnoreUser(Ignore As String)
OpenPm (Ignore)
SendChat ("iggy time bitch")
ClickPmMenu ("&Ignore User...")
End Sub
Sub MoveForm(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Function OnOff() As Long
'this finds the on off line window :P
Dim staticx As Long
Dim x As Long
x = FindWindow("#32770", vbNullString)
staticx = FindWindowEx(x, 0&, "static", vbNullString)
End Function
Function OpenSetting(ProgName As String, Setting As String)
Dim that As String
On Error GoTo Error
that = GetSetting("ToneModuleSettings", ProgName, Setting)
If that = "" Then
Error:
OpenSetting = "(Error getting setting)"
Else
OpenSetting = that
End If
End Function
Function OpenTextFile(Filename As String)
Dim F As Integer

        On Error GoTo Error
            F = FreeFile
            Open Filename For Input As F
            OpenTextFile = Input(LOF(F), F)
            Close F
            Exit Function

Error:
OpenTextFile = ""
Close F
        Exit Function
End Function
Function PastText()
PastText = Clipboard.GetText
End Function

Sub TextToPut(text As String)
Dim imclass As Long, richedit As Long
imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Call SendMessageByString(richedit, WM_SETTEXT, 0&, text)
End Sub
Sub RefreshBomb()
'BOMBERS START THIS WORKS ON YOUR/THERE YAHOO MESSAGER HEHE:>
'this refrsh bombers there buddy list :))
'to stop just put pause 1E+271
Do
ClickYMSGMenu ("Refresh")
Pause 0.009
Loop
End Sub
Sub RoomBust(Room As String)
'brakes in a room that is full :P
SendChat ("/join " + Room)
Pause 0.2
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
Pause 0.2
End Sub
Public Sub RunMenuItem(hwnd As Long, MenuCaption As String)
On Error GoTo Error
Dim intLoop         As Integer
Dim intSubLoop      As Integer
Dim intSub2Loop         As Integer
Dim intSub3Loop         As Integer
Dim intSub4Loop         As Integer
Dim lngmenu(1 To 5)     As Long
Dim lngcount(1 To 5)    As Long
Dim lngSubMenuID(1 To 4)    As Long
Dim MnCaption(1 To 4)   As String
lngmenu(1) = GetMenu(hwnd)
lngcount(1) = GetMenuItemCount(lngmenu(1))
For intLoop = 0 To lngcount(1) - 1
DoEvents
lngmenu(2) = GetSubMenu(lngmenu(1), intLoop)
lngcount(2) = GetMenuItemCount(lngmenu(2))
For intSubLoop = 0 To lngcount(2) - 1
DoEvents
lngSubMenuID(1) = GetMenuItemID(lngmenu(2), intSubLoop)
MnCaption(1) = String(75, " ")
Call GetMenuString(lngmenu(2), lngSubMenuID(1), MnCaption(1), 75, 1)


If InStr(LCase(MnCaption(1)), LCase(MenuCaption$)) Then
Call SendMessage(hwnd, WM_COMMAND, lngSubMenuID(1), 0)
Exit Sub
End If

lngmenu(3) = GetSubMenu(lngmenu(2), intSubLoop)
lngcount(3) = GetMenuItemCount(lngmenu(3))
If lngcount(3) > 0 Then
For intSub2Loop = 0 To lngcount(3) - 1
DoEvents
lngSubMenuID(2) = GetMenuItemID(lngmenu(3), intSub2Loop)
MnCaption(2) = String(75, " ")
Call GetMenuString(lngmenu(3), lngSubMenuID(2), MnCaption(2), 75, 1)

'MsgBox MnCaption(2)
'Form1.List1.AddItem MnCaption(2)

If InStr(LCase(MnCaption(2)), LCase(MenuCaption$)) Then
Call SendMessage(hwnd, WM_COMMAND, lngSubMenuID(2), 0)
Exit Sub
End If

lngmenu(4) = GetSubMenu(lngmenu(3), intSub2Loop)
lngcount(4) = GetMenuItemCount(lngmenu(4))
If lngcount(4) > 0 Then
For intSub3Loop = 0 To lngcount(4) - 1
DoEvents
lngSubMenuID(3) = GetMenuItemID(lngmenu(4), intSub3Loop)
MnCaption(3) = String(75, " ")
Call GetMenuString(lngmenu(4), lngSubMenuID(3), MnCaption(3), 75, 1)
If InStr(LCase(MnCaption(3)), LCase(MenuCaption$)) Then
Call SendMessage(hwnd, WM_COMMAND, lngSubMenuID(3), 0)
Exit Sub
End If
lngmenu(5) = GetSubMenu(lngmenu(4), intSub3Loop)
lngcount(5) = GetMenuItemCount(lngmenu(5))
If lngcount(5) > 0 Then
For intSub4Loop = 0 To lngcount(5) - 1
DoEvents
lngSubMenuID(4) = GetMenuItemID(lngmenu(5), intSub4Loop)
MnCaption(4) = String(75, " ")
Call GetMenuString(lngmenu(5), lngSubMenuID(4), MnCaption(4), 75, 1)
If InStr(LCase(MnCaption(4)), LCase(MenuCaption$)) Then
Call SendMessage(hwnd, WM_COMMAND, lngSubMenuID(4), 0)
Exit Sub
End If
Next intSub4Loop
End If
Next intSub3Loop
End If
Next intSub2Loop
End If
Next intSubLoop
Next intLoop
Error:
End Sub
Sub SaveTextFile(Filename As String, WhatToSave As String)
Dim F As Integer
On Error GoTo CloseError
    F = FreeFile
    Open Filename For Output As F
    Print #F, WhatToSave
    Close F
        Exit Sub
CloseError:

        Exit Sub
End Sub
Sub SaveTheSetting(ProgName As String, Setting As String, NewSetting As String)

On Error GoTo Error
SaveSetting "YahooToneModuleSettings", ProgName, Setting, NewSetting
Error:
Beep
End Sub
Sub ScrollDown(Box As Object)
'put ScrollDown Text1
Box.SelStart = Len(Box)
End Sub
Sub SendChat(text As String)
Dim x As Long, Button As Long
Y_GetPMWind
TextToPut text
Pause 0.06
ClickSend
Pause 0.06
If Y_GetText = text Then
ClickSend
End If
End Sub
Sub SendFile(FileToSend As String)
'this sends a file
'put the code like this:
'OpenPM(Text1)
'SendFile("C:\folders\some file.exe")
'(only YIM)
Dim imclass As Long
Dim tabtoolbar As Long
Dim yahoobuddymain As Long
Dim yalertclass As Long
Dim editx As Long
Dim Button As Long

imclass = FindWindow("imclass", vbNullString)
tabtoolbar = FindWindowEx(imclass, 0&, "tabtoolbar", vbNullString)
Call SendMessageLong(tabtoolbar, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(tabtoolbar, WM_LBUTTONUP, 0&, 0&)

If tabtoolbar = 0 Then
    Exit Sub
End If

Pause 0.1

Do
    DoEvents
    yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
    yalertclass = FindWindow("yalertclass", vbNullString)
    editx = FindWindowEx(yalertclass, 0&, "edit", vbNullString)
    Call SendMessageByString(editx, WM_SETTEXT, 0&, FileToSend)
Loop Until editx <> 0

Pause 0.1

Do
    DoEvents
    yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
    yalertclass = FindWindow("yalertclass", vbNullString)
    Button = FindWindowEx(yalertclass, 0&, "button", vbNullString)
    Button = FindWindowEx(yalertclass, Button, "button", vbNullString)
    Call SendMessageLong(Button, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)
Loop Until Button <> 0

End Sub

Sub MarxAnti()
'anit boot
Dim imclass As Long, atlce As Long, internetexplorerserver As Long
imclass = FindWindow("imclass", vbNullString)
atlce = FindWindowEx(imclass, 0&, "atl:0054c0e0", vbNullString)
internetexplorerserver = FindWindowEx(atlce, 0&, "internet explorer_server", vbNullString)
Call SendMessageLong(internetexplorerserver, WM_CLOSE, 0&, 0&)

End Sub
Function Typedin_SN()
If Not Y_ChatorPM = "pm" Then
Typedin_SN = ""
GoTo Error
End If
Dim imclass As Long, editx As Long
imclass = FindWindow("imclass", vbNullString)
editx = FindWindowEx(imclass, 0&, "edit", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(editx, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(editx, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
Typedin_SN = TheText
Error:
End Function


Function Y_ChatorPM()
'put If Y_ChatOrPM = "none or pm or chat or con" then
Dim imclass As Long, Button As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If Right(TheText, 7) = "-- Chat" Then
Y_ChatorPM = "chat"
ElseIf Right(TheText, 18) = "-- Instant Message" Then
Y_ChatorPM = "pm"
ElseIf Left(TheText, 10) = "Conference" Then
Y_ChatorPM = "con"
ElseIf Left(TheText, 10) = "Voice Conf" Then
Y_ChatorPM = "con"
Else
imclass = FindWindow("imclass", vbNullString)
Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
TL = SendMessageLong(Button, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(Button, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If TheText = "&Send" Then Y_ChatorPM = "pm"
End If
End Function

Function Y_GetText()
Dim imclass As Long, richedit As Long
imclass = FindWindow("imclass", vbNullString)
richedit = FindWindowEx(imclass, 0&, "richedit", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(richedit, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(richedit, WM_GETTEXT, TL + 1, TheText)
Y_GetText = Left(TheText, TL)
End Function

Function YesNo()
Dim It As Integer
It = Rnd
It = Right(It, 1)
If It = 1 Then
YesNo = "yes"
Else
YesNo = "no"
End If
End Function
Function AfterID(a As String)
On Error GoTo Error
If Len(a) > 1000 Then a = Right(a, 1000)
'Dim Pos As Integer
'Pos = InStrRev(A, ": ")
'If Pos = 0 Then
'A = A + " "
'Else
'A = Mid(A, Pos + 2) + " "
'End If
'AfterID = A
For This = 1 To Len(a)
a = Right(a, Len(a) - 1)
If Left(a, 2) = ": " Then
a = Right(a, Len(a) - 2)
AfterID = a
Exit Function
End If
Next
Error:
End Function
Function ASCII()
Dim Bla As String
Dim la As String
Dim d As String
Dim u As Integer
For u = 1 To (Right(Rnd, 1) + 5)
It:
Bla = Rnd + Right(Time$, 1)
Bla = Right(Bla, 1)
If Bla = "1" Or Bla = 2 Then
la = Bla
ElseIf Bla = 7 Then
la = "0"
Else
GoTo It
End If
d:
Bla = Rnd + Right(Time$, 1)
Bla = Right(Bla, 2)
Bla = Left(Bla, 1)
If la = "2" And Bla > 5 Then
GoTo d
Else
la = la + Bla
End If
s:
Bla = Rnd + Right(Time$, 1)
Bla = Right(Bla, 2)
Bla = Left(Bla, 1)
If la = "25" And Bla > 4 Then
GoTo s
Else
la = la + Bla
End If
la = la + 0
If la > 255 Then GoTo It
If la < 32 Then GoTo It
la = Chr(la)
d = d + la
Next
ASCII = d
End Function

Function GetHexColor(Color As String)
'this will turn a RGB(255,10,33) color to a #40DS33 color (like a windows color to a HTML color)
' i know this is coded slopply, but what i did was made VB print them all out like this then copyed/pasted it in here, lmao
If Color = "0" Then
 Color = "00"
GoTo h
ElseIf Color = "1" Then
 Color = "01"
GoTo h
ElseIf Color = "2" Then
 Color = "02"
GoTo h
ElseIf Color = "3" Then
 Color = "03"
GoTo h
ElseIf Color = "4" Then
 Color = "04"
GoTo h
ElseIf Color = "5" Then
 Color = "05"
GoTo h
ElseIf Color = "6" Then
 Color = "06"
GoTo h
ElseIf Color = "7" Then
 Color = "07"
GoTo h
ElseIf Color = "8" Then
 Color = "08"
GoTo h
ElseIf Color = "9" Then
 Color = "09"
GoTo h
ElseIf Color = "10" Then
 Color = "0A"
GoTo h
ElseIf Color = "11" Then
 Color = "0B"
GoTo h
ElseIf Color = "12" Then
 Color = "0C"
GoTo h
ElseIf Color = "13" Then
 Color = "0D"
GoTo h
ElseIf Color = "14" Then
 Color = "0E"
GoTo h
ElseIf Color = "15" Then
 Color = "0F"
GoTo h
ElseIf Color = "16" Then
 Color = "10"
GoTo h
ElseIf Color = "17" Then
 Color = "11"
GoTo h
ElseIf Color = "18" Then
 Color = "12"
GoTo h
ElseIf Color = "19" Then
 Color = "13"
GoTo h
ElseIf Color = "20" Then
 Color = "14"
GoTo h
ElseIf Color = "21" Then
 Color = "15"
GoTo h
ElseIf Color = "22" Then
 Color = "16"
GoTo h
ElseIf Color = "23" Then
 Color = "17"
GoTo h
ElseIf Color = "24" Then
 Color = "18"
GoTo h
ElseIf Color = "25" Then
 Color = "19"
GoTo h
ElseIf Color = "26" Then
 Color = "1A"
GoTo h
ElseIf Color = "27" Then
 Color = "1B"
GoTo h
ElseIf Color = "28" Then
 Color = "1C"
GoTo h
ElseIf Color = "29" Then
 Color = "1D"
GoTo h
ElseIf Color = "30" Then
 Color = "1E"
GoTo h
ElseIf Color = "31" Then
 Color = "1F"
GoTo h
ElseIf Color = "32" Then
 Color = "20"
GoTo h
ElseIf Color = "33" Then
 Color = "21"
GoTo h
ElseIf Color = "34" Then
 Color = "22"
GoTo h
ElseIf Color = "35" Then
 Color = "23"
GoTo h
ElseIf Color = "36" Then
 Color = "24"
GoTo h
ElseIf Color = "37" Then
 Color = "25"
GoTo h
ElseIf Color = "38" Then
 Color = "26"
GoTo h
ElseIf Color = "39" Then
 Color = "27"
GoTo h
ElseIf Color = "40" Then
 Color = "28"
GoTo h
ElseIf Color = "41" Then
 Color = "29"
GoTo h
ElseIf Color = "42" Then
 Color = "2A"
GoTo h
ElseIf Color = "43" Then
 Color = "2B"
GoTo h
ElseIf Color = "44" Then
 Color = "2C"
GoTo h
ElseIf Color = "45" Then
 Color = "2D"
GoTo h
ElseIf Color = "46" Then
 Color = "2E"
GoTo h
ElseIf Color = "47" Then
 Color = "2F"
GoTo h
ElseIf Color = "48" Then
 Color = "30"
GoTo h
ElseIf Color = "49" Then
 Color = "31"
GoTo h
ElseIf Color = "50" Then
 Color = "32"
GoTo h
ElseIf Color = "51" Then
 Color = "33"
GoTo h
ElseIf Color = "52" Then
 Color = "34"
GoTo h
ElseIf Color = "53" Then
 Color = "35"
GoTo h
ElseIf Color = "54" Then
 Color = "36"
GoTo h
ElseIf Color = "55" Then
 Color = "37"
GoTo h
ElseIf Color = "56" Then
 Color = "38"
GoTo h
ElseIf Color = "57" Then
 Color = "39"
GoTo h
ElseIf Color = "58" Then
 Color = "3A"
GoTo h
ElseIf Color = "59" Then
 Color = "3B"
GoTo h
ElseIf Color = "60" Then
 Color = "3C"
GoTo h
ElseIf Color = "61" Then
 Color = "3D"
GoTo h
ElseIf Color = "62" Then
 Color = "3E"
GoTo h
ElseIf Color = "63" Then
 Color = "3F"
GoTo h
ElseIf Color = "64" Then
 Color = "40"
GoTo h
ElseIf Color = "65" Then
 Color = "41"
GoTo h
ElseIf Color = "66" Then
 Color = "42"
GoTo h
ElseIf Color = "67" Then
 Color = "43"
GoTo h
ElseIf Color = "68" Then
 Color = "44"
GoTo h
ElseIf Color = "69" Then
 Color = "45"
GoTo h
ElseIf Color = "70" Then
 Color = "46"
GoTo h
ElseIf Color = "71" Then
 Color = "47"
GoTo h
ElseIf Color = "72" Then
 Color = "48"
GoTo h
ElseIf Color = "73" Then
 Color = "49"
GoTo h
ElseIf Color = "74" Then
 Color = "4A"
GoTo h
ElseIf Color = "75" Then
 Color = "4B"
GoTo h
ElseIf Color = "76" Then
 Color = "4C"
GoTo h
ElseIf Color = "77" Then
 Color = "4D"
GoTo h
ElseIf Color = "78" Then
 Color = "4E"
GoTo h
ElseIf Color = "79" Then
 Color = "4F"
GoTo h
ElseIf Color = "80" Then
 Color = "50"
GoTo h
ElseIf Color = "81" Then
 Color = "51"
GoTo h
ElseIf Color = "82" Then
 Color = "52"
GoTo h
ElseIf Color = "83" Then
 Color = "53"
GoTo h
ElseIf Color = "84" Then
 Color = "54"
GoTo h
ElseIf Color = "85" Then
 Color = "55"
GoTo h
ElseIf Color = "86" Then
 Color = "56"
GoTo h
ElseIf Color = "87" Then
 Color = "57"
GoTo h
ElseIf Color = "88" Then
 Color = "58"
GoTo h
ElseIf Color = "89" Then
 Color = "59"
GoTo h
ElseIf Color = "90" Then
 Color = "5A"
GoTo h
ElseIf Color = "91" Then
 Color = "5B"
GoTo h
ElseIf Color = "92" Then
 Color = "5C"
GoTo h
ElseIf Color = "93" Then
 Color = "5D"
GoTo h
ElseIf Color = "94" Then
 Color = "5E"
GoTo h
ElseIf Color = "95" Then
 Color = "5F"
GoTo h
ElseIf Color = "96" Then
 Color = "60"
GoTo h
ElseIf Color = "97" Then
 Color = "61"
GoTo h
ElseIf Color = "98" Then
 Color = "62"
GoTo h
ElseIf Color = "99" Then
 Color = "63"
GoTo h
ElseIf Color = "100" Then
 Color = "64"
GoTo h
ElseIf Color = "101" Then
 Color = "65"
GoTo h
ElseIf Color = "102" Then
 Color = "66"
GoTo h
ElseIf Color = "103" Then
 Color = "67"
GoTo h
ElseIf Color = "104" Then
 Color = "68"
GoTo h
ElseIf Color = "105" Then
 Color = "69"
GoTo h
ElseIf Color = "106" Then
 Color = "6A"
GoTo h
ElseIf Color = "107" Then
 Color = "6B"
GoTo h
ElseIf Color = "108" Then
 Color = "6C"
GoTo h
ElseIf Color = "109" Then
 Color = "6D"
GoTo h
ElseIf Color = "110" Then
 Color = "6E"
GoTo h
ElseIf Color = "111" Then
 Color = "6F"
GoTo h
ElseIf Color = "112" Then
 Color = "70"
GoTo h
ElseIf Color = "113" Then
 Color = "71"
GoTo h
ElseIf Color = "114" Then
 Color = "72"
GoTo h
ElseIf Color = "115" Then
 Color = "73"
GoTo h
ElseIf Color = "116" Then
 Color = "74"
GoTo h
ElseIf Color = "117" Then
 Color = "75"
GoTo h
ElseIf Color = "118" Then
 Color = "76"
GoTo h
ElseIf Color = "119" Then
 Color = "77"
GoTo h
ElseIf Color = "120" Then
 Color = "78"
GoTo h
ElseIf Color = "121" Then
 Color = "79"
GoTo h
ElseIf Color = "122" Then
 Color = "7A"
GoTo h
ElseIf Color = "123" Then
 Color = "7B"
GoTo h
ElseIf Color = "124" Then
 Color = "7C"
GoTo h
ElseIf Color = "125" Then
 Color = "7D"
GoTo h
ElseIf Color = "126" Then
 Color = "7E"
GoTo h
ElseIf Color = "127" Then
 Color = "7F"
GoTo h
ElseIf Color = "128" Then
 Color = "80"
GoTo h
ElseIf Color = "129" Then
 Color = "81"
GoTo h
ElseIf Color = "130" Then
 Color = "82"
GoTo h
ElseIf Color = "131" Then
 Color = "83"
GoTo h
ElseIf Color = "132" Then
 Color = "84"
GoTo h
ElseIf Color = "133" Then
 Color = "85"
GoTo h
ElseIf Color = "134" Then
 Color = "86"
GoTo h
ElseIf Color = "135" Then
 Color = "87"
GoTo h
ElseIf Color = "136" Then
 Color = "88"
GoTo h
ElseIf Color = "137" Then
 Color = "89"
GoTo h
ElseIf Color = "138" Then
 Color = "8A"
GoTo h
ElseIf Color = "139" Then
 Color = "8B"
GoTo h
ElseIf Color = "140" Then
 Color = "8C"
GoTo h
ElseIf Color = "141" Then
 Color = "8D"
GoTo h
ElseIf Color = "142" Then
 Color = "8E"
GoTo h
ElseIf Color = "143" Then
 Color = "8F"
GoTo h
ElseIf Color = "144" Then
 Color = "90"
GoTo h
ElseIf Color = "145" Then
 Color = "91"
GoTo h
ElseIf Color = "146" Then
 Color = "92"
GoTo h
ElseIf Color = "147" Then
 Color = "93"
GoTo h
ElseIf Color = "148" Then
 Color = "94"
GoTo h
ElseIf Color = "149" Then
 Color = "95"
GoTo h
ElseIf Color = "150" Then
 Color = "96"
GoTo h
ElseIf Color = "151" Then
 Color = "97"
GoTo h
ElseIf Color = "152" Then
 Color = "98"
GoTo h
ElseIf Color = "153" Then
 Color = "99"
GoTo h
ElseIf Color = "154" Then
 Color = "9A"
GoTo h
ElseIf Color = "155" Then
 Color = "9B"
GoTo h
ElseIf Color = "156" Then
 Color = "9C"
GoTo h
ElseIf Color = "157" Then
 Color = "9D"
GoTo h
ElseIf Color = "158" Then
 Color = "9E"
GoTo h
ElseIf Color = "159" Then
 Color = "9F"
GoTo h
ElseIf Color = "160" Then
 Color = "A0"
GoTo h
ElseIf Color = "161" Then
 Color = "A1"
GoTo h
ElseIf Color = "162" Then
 Color = "A2"
GoTo h
ElseIf Color = "163" Then
 Color = "A3"
GoTo h
ElseIf Color = "164" Then
 Color = "A4"
GoTo h
ElseIf Color = "165" Then
 Color = "A5"
GoTo h
ElseIf Color = "166" Then
 Color = "A6"
GoTo h
ElseIf Color = "167" Then
 Color = "A7"
GoTo h
ElseIf Color = "168" Then
 Color = "A8"
GoTo h
ElseIf Color = "169" Then
 Color = "A9"
GoTo h
ElseIf Color = "170" Then
 Color = "AA"
GoTo h
ElseIf Color = "171" Then
 Color = "AB"
GoTo h
ElseIf Color = "172" Then
 Color = "AC"
GoTo h
ElseIf Color = "173" Then
 Color = "AD"
GoTo h
ElseIf Color = "174" Then
 Color = "AE"
GoTo h
ElseIf Color = "175" Then
 Color = "AF"
GoTo h
ElseIf Color = "176" Then
 Color = "B0"
GoTo h
ElseIf Color = "177" Then
 Color = "B1"
GoTo h
ElseIf Color = "178" Then
 Color = "B2"
GoTo h
ElseIf Color = "179" Then
 Color = "B3"
GoTo h
ElseIf Color = "180" Then
 Color = "B4"
GoTo h
ElseIf Color = "181" Then
 Color = "B5"
GoTo h
ElseIf Color = "182" Then
 Color = "B6"
GoTo h
ElseIf Color = "183" Then
 Color = "B7"
GoTo h
ElseIf Color = "184" Then
 Color = "B8"
GoTo h
ElseIf Color = "185" Then
 Color = "B9"
GoTo h
ElseIf Color = "186" Then
 Color = "BA"
GoTo h
ElseIf Color = "187" Then
 Color = "BB"
GoTo h
ElseIf Color = "188" Then
 Color = "BC"
GoTo h
ElseIf Color = "189" Then
 Color = "BD"
GoTo h
ElseIf Color = "190" Then
 Color = "BE"
GoTo h
ElseIf Color = "191" Then
 Color = "BF"
GoTo h
ElseIf Color = "192" Then
 Color = "C0"
GoTo h
ElseIf Color = "193" Then
 Color = "C1"
GoTo h
ElseIf Color = "194" Then
 Color = "C2"
GoTo h
ElseIf Color = "195" Then
 Color = "C3"
GoTo h
ElseIf Color = "196" Then
 Color = "C4"
GoTo h
ElseIf Color = "197" Then
 Color = "C5"
GoTo h
ElseIf Color = "198" Then
 Color = "C6"
GoTo h
ElseIf Color = "199" Then
 Color = "C7"
GoTo h
ElseIf Color = "200" Then
 Color = "C8"
GoTo h
ElseIf Color = "201" Then
 Color = "C9"
GoTo h
ElseIf Color = "202" Then
 Color = "CA"
GoTo h
ElseIf Color = "203" Then
 Color = "CB"
GoTo h
ElseIf Color = "204" Then
 Color = "CC"
GoTo h
ElseIf Color = "205" Then
 Color = "CD"
GoTo h
ElseIf Color = "206" Then
 Color = "CE"
GoTo h
ElseIf Color = "207" Then
 Color = "CF"
GoTo h
ElseIf Color = "208" Then
 Color = "D0"
GoTo h
ElseIf Color = "209" Then
 Color = "D1"
GoTo h
ElseIf Color = "210" Then
 Color = "D2"
GoTo h
ElseIf Color = "211" Then
 Color = "D3"
GoTo h
ElseIf Color = "212" Then
 Color = "D4"
GoTo h
ElseIf Color = "213" Then
 Color = "D5"
GoTo h
ElseIf Color = "214" Then
 Color = "D6"
GoTo h
ElseIf Color = "215" Then
 Color = "D7"
GoTo h
ElseIf Color = "216" Then
 Color = "D8"
GoTo h
ElseIf Color = "217" Then
 Color = "D9"
GoTo h
ElseIf Color = "218" Then
 Color = "DA"
GoTo h
ElseIf Color = "219" Then
 Color = "DB"
GoTo h
ElseIf Color = "220" Then
 Color = "DC"
GoTo h
ElseIf Color = "221" Then
 Color = "DD"
GoTo h
ElseIf Color = "222" Then
 Color = "DE"
GoTo h
ElseIf Color = "223" Then
 Color = "DF"
GoTo h
ElseIf Color = "224" Then
 Color = "E0"
GoTo h
ElseIf Color = "225" Then
 Color = "E1"
GoTo h
ElseIf Color = "226" Then
 Color = "E2"
GoTo h
ElseIf Color = "227" Then
 Color = "E3"
GoTo h
ElseIf Color = "228" Then
 Color = "E4"
GoTo h
ElseIf Color = "229" Then
 Color = "E5"
GoTo h
ElseIf Color = "230" Then
 Color = "E6"
GoTo h
ElseIf Color = "231" Then
 Color = "E7"
GoTo h
ElseIf Color = "232" Then
 Color = "E8"
GoTo h
ElseIf Color = "233" Then
 Color = "E9"
GoTo h
ElseIf Color = "234" Then
 Color = "EA"
GoTo h
ElseIf Color = "235" Then
 Color = "EB"
GoTo h
ElseIf Color = "236" Then
 Color = "EC"
GoTo h
ElseIf Color = "237" Then
 Color = "ED"
GoTo h
ElseIf Color = "238" Then
 Color = "EE"
GoTo h
ElseIf Color = "239" Then
 Color = "EF"
GoTo h
ElseIf Color = "240" Then
 Color = "F0"
GoTo h
ElseIf Color = "241" Then
 Color = "F1"
GoTo h
ElseIf Color = "242" Then
 Color = "F2"
GoTo h
ElseIf Color = "243" Then
 Color = "F3"
GoTo h
ElseIf Color = "244" Then
 Color = "F4"
GoTo h
ElseIf Color = "245" Then
 Color = "F5"
GoTo h
ElseIf Color = "246" Then
 Color = "F6"
GoTo h
ElseIf Color = "247" Then
 Color = "F7"
GoTo h
ElseIf Color = "248" Then
 Color = "F8"
GoTo h
ElseIf Color = "249" Then
 Color = "F9"
GoTo h
ElseIf Color = "250" Then
 Color = "FA"
GoTo h
ElseIf Color = "251" Then
 Color = "FB"
GoTo h
ElseIf Color = "252" Then
 Color = "FC"
GoTo h
ElseIf Color = "253" Then
 Color = "FD"
GoTo h
ElseIf Color = "254" Then
 Color = "FE"
GoTo h
ElseIf Color = "255" Then
 Color = "FF"
GoTo h
End If
h:
GetHexColor = Color
End Function


Function IsLettersThere(Latters As String, OverAll As String)
'this will see if certon  letters r in a string, like a command for a BOT
'put like: If IsLettersThere("bot",GetYahooText) = "yes" then YahooSend "whos the bot??"
Dim a As String
Dim This As Integer
For This = 1 To Len(OverAll)
OverAll = Right(OverAll, 1) + Left(OverAll, Len(OverAll) - 1)
a = Left(OverAll, Len(Latters))
If LCase(a) = LCase(Latters) Then
IsLettersThere = "yes"
Exit Function
End If
Next
IsLettersThere = "no"
End Function
Sub SafeList(Him As String)
Him = LCase(Him) ' so its not case censitive
If Him = "(victim)" Then Him = "chat"
If Him = "" Then Him = "chat"
If Him = "chat" Then Him = LCase(Y_SN)
If Him = "" Then Exit Sub
'put the names u wnat to protect from ur prog booting/bombing them here:
If Him = "j0k4w1ld" Then End
If Him = "o0o_m4d_cl0wn_o0o" Then End
If Him = "Klownin_Inc" Then End
If Him = "Dr.Boy" Then End
If Him = "illusions_of_jesus" Then End
If Him = "bstc111144" Then End
If Him = "l_dark_shadow_l" Then End
If Him = "__doomsday__" Then End
If Him = "addalaide" Then End
If Him = "cisco_tm" Then End
If Him = "do-gg" Then End
If Him = "genocide_on_earth" Then End
If Him = "gucci_soldier" Then End
If Him = "hstwarning" Then End
If Him = "llll_o0_chat_error_0o_llll" Then End
If Him = "m--a--g--i--c@prodigy.net" Then End
If Him = "marxs_baby" Then End
If Him = "rl-brian" Then End
If Him = "sweet_bitch__" Then End
If Him = "system_acc3ss" Then End
If Him = "zeguine_04" Then End
If Him = "cisco_tm" Then End
If Him = "haunted_by_angels" Then End
If Him = "x0o_cybermafia_styles_man_o0x" Then End
If Him = "tank_dogg_2004" Then End
If Him = "Juggalo_Joker2003" Then End

End Sub
Function IsLogin()
Dim x As Long, Button As Long
x = FindWindow("#32770", vbNullString)
Button = FindWindowEx(x, 0&, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(Button, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(Button, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If TheText = "Get a &Yahoo! ID" Then
IsLogin = "yes"
Else
IsLogin = "no"
End If
End Function
Function Y_GetLastLogin()
'you could use this for like a BOT or sogin
'put like: Text1.selText = "Hello, " + Y_GetLastLogin + "!"
Y_GetLastLogin = GetSettingString(HKEY_CURRENT_USER, "Software\Yahoo\Pager", "Yahoo! User ID", "") ' takes it off the registry :>
End Function
Function GetWinamp2Song()
'this will get current song on winamp
On Error GoTo Error
Dim winampvx As Long
winampvx = FindWindow("winamp v1.x", vbNullString) ' says V1 but only works on V2... wtf?
Dim TheText As String, TL As Long
TL = SendMessageLong(winampvx, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(winampvx, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL - 9)
If Right(TheText, 10) = " - Winamp " Then TheText = Left(TheText, Len(TheText) - 10)
GetWinamp2Song = TheText
Error:
End Function
Sub ClickPagerMenu(Ack As String)
'clicks menu on the pager
On Error GoTo Error
Dim yahoobuddymain As Long
yahoobuddymain = FindWindow("YahooBuddyMain", vbNullString)
Call RunMenuItem(yahoobuddymain, Right(Ack, Len(Ack) - 1)) ' Note: u might wonna change the 1 to like 5 is what u wonna click starts with the same word as something b4 it, like 'Be' , but 1 is good cause it fucks up the first work lol
Error:
End Sub


Function Y_IsInvisible()
'will tell u if ur inviisble :P
'put like: if Y_IsInvisible = "no" then bla
Dim yahoobuddymain As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(yahoobuddymain, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(yahoobuddymain, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If Right(TheText, 11) = "(Invisible)" Then
TheText = "yes"
Else
TheText = "no"
End If
Y_IsInvisible = TheText
End Function

Sub Y_ChangeChatAdd()
'this will change that chat ad on YIM
Dim HTMLCode As String
HTMLCode = "<BODY BGCOLOR=#000000><Center><a href=http://www.Klownin-Inc.com target=_newwin><img src=URL OF THE PIC HERE LIKE HTTP:// BALH BALH....JPG/.GIF border=0></a>" 'put the HTML for what ever u want the ad to be here :P
SaveSettingString HKEY_CURRENT_USER, "Software\Yahoo\Pager\yurl", "Chat Adurl", "about:" + HTMLCode + "<" ' this edits messy and changed the way the chatroom Ad works :P
End Sub
Function GetStaticx()
Dim staticx As Long
staticx = FindWindow("static", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(staticx, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(staticx, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If Not Right(TheText, 3) = "line" Then TheText = ""
GetStaticx = TheText
End Function

Public Function CloseRequest1()
Dim yahoobuddymain As Long
Dim x As Long
Dim Button As Long
yahoobuddymain = FindWindow("yahoobuddymain", vbNullString)
x = FindWindow("#32770", vbNullString)
Button = FindWindowEx(x, 0&, "button", vbNullString)
Button = FindWindowEx(x, Button, "button", vbNullString)
Call PostMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(Button, WM_KEYUP, VK_SPACE, 0&)
End Function
Function CountTimeInString(OverAll As String, ToFind As String)
' this will find how meny times a string is in
'put like: If CountTimeInString(A,"<") > 10 then Exit Sub
Dim a As String
Dim T As Integer
T = 0
For This = 1 To Len(OverAll)
a = Right(OverAll, This)
a = Left(a, Len(ToFind))
If a = ToFind Then T = T + 1
Next
CountTimeInString = T
End Function

Function LastLine(a As String)
On Error GoTo Error
If Len(a) > 1000 Then a = Right(a, 1000)
Dim Pos As Integer
Pos = InStrRev(a, vbCrLf)
If Pos = 0 Then
a = a + " "
Else
a = Mid(a, Pos + 2) + " "
End If
LastLine = a
Error:
End Function
Function OnXP()
'this sees if ur on XP cause it screws up sizes...
'(this will also say 'no' if ur on XP but have it set to look like the older windows)
'out it like this: If OnXP = "yes" Then Width = 1400 Else Width = 1100
Dim shelltraywnd As Long, Button As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
Button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(Button, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(Button, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If TheText = "start" Then ' sees if u have XP stat buttion
OnXP = "yes"
Else ' sees if u dont have XP start buttion
OnXP = "no"
End If
End Function


Sub Y_Anti()
Dim imclass As Long
Dim atlce As Long

imclass = FindWindow("imclass", vbNullString)
atlce = FindWindowEx(imclass, 0&, "atl:0054c0e0", vbNullString)
Call SendMessageLong(atlce, WM_CLOSE, 0&, 0&)

If atlce = 0 Then
End If
End Sub
Sub Y_ChangeFontText(Bla As String)
Dim imclass As Long, toolbarparent As Long, toolbarwindow As Long
Dim combobox As Long
imclass = FindWindow("imclass", vbNullString)
toolbarparent = FindWindowEx(imclass, 0&, "toolbarparent", vbNullString)
toolbarwindow = FindWindowEx(toolbarparent, 0&, "toolbarwindow32", vbNullString)
combobox = FindWindowEx(toolbarwindow, 0&, "combobox", vbNullString)
Call SendMessageByString(combobox, WM_SETTEXT, 0&, Bla)
End Sub

Function Y_Chat()
'this gets the chat name
'put YahooSend "bla bla bla " + Y_Chat + " bla bla bla"
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
Y_Chat = Left(TheText, Len(TheText) - 8)
End Function

Function Y_GetPMWind()
'put Y_GetPMWind to activate the pm window
'(only works in PMs on YIM)
Y_GetPMWind = FindWindow("imclass", vbNullString)
End Function



Sub Y_KillChatAdd()
'this will kill the chat ad :P
'put Y_KillChatAdd in a timer
If Not Y_ChatorPM = "chat" Then Exit Sub
Dim imclass As Long, atla As Long, atlaxwin As Long
Dim shellembedding As Long, shelldocobjectview As Long, internetexplorerserver As Long
imclass = FindWindow("imclass", vbNullString)
atla = FindWindowEx(imclass, 0&, "atl:0054a350", vbNullString)
atlaxwin = FindWindowEx(atla, 0&, "atlaxwin7", vbNullString)
shellembedding = FindWindowEx(atlaxwin, 0&, "shell embedding", vbNullString)
shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
internetexplorerserver = FindWindowEx(shelldocobjectview, 0&, "internet explorer_server", vbNullString)
Call SendMessageLong(internetexplorerserver, WM_CLOSE, 0&, 0&)
End Sub

Function Y_LastAlart()
'this gets teh last like 'Bla Is Online' thingy off the coner :D
'i figured this out on axedent after i had givin up on it, so meby my brain sub-conchisly did it! :O
Dim x As Long, staticx As Long
x = FindWindow("#32770", vbNullString)
staticx = FindWindowEx(x, 0&, "static", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(staticx, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(staticx, WM_GETTEXT, TL + 1, TheText)
Y_LastAlart = Left(TheText, TL)
End Function

Sub Y_PMCaption(Bla As String)
'this will add ur caption to the PM title bar, and y_sn wills till work :>
If Not Y_ChatorPM = "pm" Then Exit Sub
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
If Right(Y_Caption, 1) = Chr(32) Then Exit Sub
If Right(Y_Caption, Len(Bla)) = Bla Then Exit Sub ' sees if its alreatty changed
If Right(Y_Caption, 7) = "PM Fixd" Then Exit Sub ' sees if they have my PM Fix installed, this tends to screw it up lol
Call SendMessageByString(imclass, WM_SETTEXT, 0&, Y_SN + " - " + (Bla + Chr(32))) ' the " - " is NEDDED so the y_snw ill still work !!
End Sub

Function Y_SN()
'this sees the SN of who evers PM is open
'only YIM
'put: Text1 = Y_SN
If IsYahooOpen = "no" Then GoTo hell 'sees if u have YIM open
If Y_ChatorPM = "chat" Then ' if its a chat window
Y_SN = Y_Chat ' makes y_sn the chat name
Exit Function
End If
If Not Y_ChatorPM = "pm" Then 'sees if u have  a PM opened ;)
Y_SN = ""
GoTo hell
End If
Dim imclass As Long
imclass = FindWindow("imclass", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TL + 1, TheText)
Dim Thesn As String
Thesn = Left(TheText, TL)
Dim a As String
Dim B As String
B = Thesn
Dim This As Integer
For This = 2 To Len(Thesn)
a = Left(Thesn, This)
a = Right(a, 2)
If a = " -" Then
B = Left(B, This - 2)
If IsLettersThere("(", B) = "yes" And IsLettersThere(")", B) = "yes" Then ' if u have nicknames on :P
Dim FASDF As Integer
For FASDF = 1 To Len(B)
B = Right(B, Len(B) - 1)
If Left(B, 1) = Chr(32) Then
B = Right(B, Len(B) - 2)
GoTo BBS
End If
Next
End If
BBS:
Y_SN = B
GoTo hell
End If
Next
hell:
End Function



Public Sub OpenLink(Form As Form, url As String)
ShellExecute Form.hwnd, "Open", url, "", "", 1
End Sub

Public Sub Drag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub



Public Function GetFileDate(Filename As String) As String
Dim strTime As String, strDate As String
    On Error Resume Next
    GetFileDate = FileDateTime(Filename)
    strDate = "Date: " & Split(GetFileDate, " ")(0)
    strTime = "Time: " & Split(GetFileDate, " ")(1)
GetFileDate = strDate & vbCrLf & strTime
End Function

Public Function GetFileSize(Filename) As String
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(Filename)


    If TempStr >= "1024" Then
        'KB
        TempStr = CCur(TempStr / 1024) & "KB"
    Else


        If TempStr >= "1048576" Then
            'MB
            TempStr = CCur(TempStr / (1024 * 1024)) & "KB"
        Else
            TempStr = CCur(TempStr) & "B"
        End If
    End If
    GetFileSize = TempStr
    Exit Function
Gfserror:
    GetFileSize = "0B"
    Resume
End Function
Public Function DeleteFile(Filename As String)

On Error Resume Next
SetAttr Filename, vbNormal
Kill Filename
End Function
Sub AntiBoot1244()
Dim imclass As Long, atlefb As Long
imclass& = FindWindow("IMClass", vbNullString)
atlefb = FindWindowEx(imclass&, 0&, "ATL:0054E0E0", vbNullString)
Call SendMessageLong(atlefb, WM_CLOSE, 0&, 0&)
imclass& = FindWindow("IMClass", vbNullString)
atlefb = FindWindowEx(imclass&, 0&, "ATL:0054E0E0", vbNullString)
Call SendMessageLong(atlefb, WM_CLOSE, 0&, 0&)
imclass& = FindWindow("IMClass", vbNullString)
atlefb = FindWindowEx(imclass&, 0&, "ATL:0054E0E0", vbNullString)
Call SendMessageLong(atlefb, WM_CLOSE, 0&, 0&)
End Sub
Sub Y_SetStatus(Blah As String)
Dim It As Long
It = FindWindow("yahoobuddymain", vbNullString)
SaveSettingString HKEY_CURRENT_USER, "Software\Yahoo\Pager\profiles\" + Y_GetLastLogin + "\Custom Msgs", 1, Blah ' saves the message to registry
SendMessage It, WM_COMMAND, 388, 1&
End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Default As String) As String
    Dim lngValueType As Long
    Dim strBuffer As String
    Dim lngDataBufferSize As Long
    Dim intZeroPos As Integer

    
    If Not IsEmpty(Default) Then
      GetSettingString = Default
    Else
      GetSettingString = ""
    End If

    
    RegOpenKey hKey, strPath, hCurKey
    RegQueryValueEx hCurKey, strValue, 0&, lngValueType, ByVal 0&, lngDataBufferSize

    If lngValueType = REG_SZ Then
        
        strBuffer = String(lngDataBufferSize, " ")
        RegQueryValueEx hCurKey, strValue, 0&, 0&, ByVal strBuffer, lngDataBufferSize
    
        intZeroPos = InStr(strBuffer, Chr$(0))
        
        If intZeroPos > 0 Then
            GetSettingString = Left$(strBuffer, intZeroPos - 1)
        Else
            GetSettingString = strBuffer
        End If
    End If
    RegCloseKey hCurKey
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    RegCreateKey hKey, strPath, hCurKey
    RegSetValueEx hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    RegCloseKey hCurKey
End Sub

Sub Status_ScrollLeft(txt As String)
Dim x As Integer
For x = 0 To Len(txt)
Y_SetStatus Left(txt, x)
Pause 0.5
Next
End Sub

Sub Status_ScrollRight(txt As String)
Dim x As Integer
For x = 0 To Len(txt)
Y_SetStatus Right(txt, x)
Pause 0.5
Next
End Sub

Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub




Function GetName()
On Error GoTo hell
Dim x As Integer
x = InStr(Get_Caption, " -") - 1
GetName = Left(Get_Caption, x)
hell:
End Function

Function Get_Caption()
Dim imclass As Long, text As String, TLn As Long
imclass = FindWindow("imclass", vbNullString)
TLn = SendMessageLong(imclass, WM_GETTEXTLENGTH, 0&, 0&)
text = String(TLn + 1, " ")
Call SendMessageByString(imclass, WM_GETTEXT, TLn + 1, text)
Get_Caption = Left(text, TLn)
End Function

Sub Sleep(interval)
Dim atime
atime = Timer
Do While Timer - atime < Val(interval)
DoEvents
Loop
End Sub


Sub MarxMove(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub


Sub clearchat() 'clears 1254 and lower
Dim text As String
text = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
SendChat text
Pause 0.4
SendChat text
Pause 0.4
SendChat text
Pause 0.4
SendChat text
Pause 0.4
SendChat text
Pause 0.4
SendChat text
Pause 0.4
SendChat text
Pause 0.4
SendChat text
Pause 0.4
SendChat text


End Sub
Public Function FileExists(ByVal Filename As String) As Boolean
    
    Dim FileInfo As Variant

    'Set Default
    FileExists = True
    
    'Set up error handler
    On Error Resume Next

    'Attempt to grab date and time
    FileInfo = FileDateTime(Filename)

    'Process errors
    Select Case Err
        Case 53, 76, 68   'File Does Not Exist
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, vbOKOnly, "Error"
                End
            End If
    End Select
    
End Function
Sub WinAMP_SetEQValue(EQ_Index As Long, EQ_Value As Long)
    
    '-------------------------------------------------------------'
    'Sets the EQ slider value for the slider specified by EQ_Index'
    '-------------------------------------------------------------'
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Sub
    End If
    
    'eq values will range from 0(top) to 63(bottom)
    
    'range check the EQ Index
    If EQ_Index < 0 Or EQ_Index > 9 Then Exit Sub
    
    'range check the new EQ value
    If EQ_Value < 0 Or EQ_Value > 63 Then Exit Sub
    
    'we have to query the eq line we want to change first
    SendMessage hWndWinAMP, WM_USER, EQ_Index, WA_GETEQDATA
    
    'now we send the new eq value to the selected eq line
    SendMessage hWndWinAMP, WM_USER, EQ_Value, WA_SETEQDATA

End Sub
Sub WinAMP_SetPreAmpValue(PreAmp_Value As Long)
    
    '----------------------'
    'Sets the Pre-amp value'
    '----------------------'
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Sub
    End If
    
    'pre-amp values will range from 0(top) to 63(bottom)
    
    'range check the new pre-amp value
    If PreAmp_Value < 0 Or PreAmp_Value > 63 Then Exit Sub
    
    'we have to query the pre-amp first
    SendMessage hWndWinAMP, WM_USER, 10, WA_GETEQDATA
    
    'now we send the new pre-amp value
    SendMessage hWndWinAMP, WM_USER, PreAmp_Value, WA_SETEQDATA

End Sub
Sub WinAMP_ClearPlaylist()
    
    '------------------------'
    'Clears WinAMP's playlist'
    '------------------------'
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Sub
    End If

    SendMessage hWndWinAMP, WM_USER, 0, WA_CLEARPLAYLIST

End Sub


Function WinAMP_FindWindow() As Boolean

    '-------------------------------------'
    'Retrieves a handle to WinAMP's window'
    '-------------------------------------'

    hWndWinAMP = FindWindow("Winamp v1.x", vbNullString)
    
    If hWndWinAMP <> 0 Then
       WinAMP_FindWindow = True
    Else
       WinAMP_FindWindow = False
    End If

End Function


Function WinAMP_GetStatus() As String

    '----------------------------------------------------------'
    'Retrieves the status of WinAMP: PLAYING, PAUSED or STOPPED'
    '----------------------------------------------------------'

    Dim Status As Long
    Dim i As Long

    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Function
    End If

    Status = SendMessage(hWndWinAMP, WM_USER, 0, WA_GETSTATUS)
    
    Select Case Status
       Case 1
          WinAMP_GetStatus = "PLAYING"
       Case 3
          WinAMP_GetStatus = "PAUSED"
       Case Else
          WinAMP_GetStatus = "STOPPED"
    End Select

End Function

Function WinAMP_GetTrackPosition() As Long

    '---------------------------------------------------'
    'Retrieves the position of the current track in secs'
    '---------------------------------------------------'

    Dim ReturnPos As Long
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Function
    End If

    'ReturnPos will contain the current track pos in milliseconds
    'or -1 if no track is playing or an error occurs
    ReturnPos = SendMessage(hWndWinAMP, WM_USER, 0, WA_GETTRACKPOSITION)

    If ReturnPos <> -1 Then
       WinAMP_GetTrackPosition = ReturnPos \ 1000   'convert ReturnPos to secs
    Else
       WinAMP_GetTrackPosition = -1
    End If

End Function

Function WinAMP_GetEQValue(EQ_Index As Long) As Long

    '------------------------------------------------------------------'
    'Retrieves the EQ slider value for the slider specified by EQ_Index'
    '------------------------------------------------------------------'

    Dim ReturnValue As Long
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Function
    End If
    
    'range check the EQ Index (between 0 and 9, inclusive)
    If EQ_Index < 0 Or EQ_Index > 9 Then
       WinAMP_GetEQValue = -1
       Exit Function
    End If
    
    'return value will hold a value for the selected EQ slider
    'values will range from 0(top) to 63(bottom)
    ReturnValue = SendMessage(hWndWinAMP, WM_USER, EQ_Index, WA_GETEQDATA)

    WinAMP_GetEQValue = ReturnValue

End Function
Function WinAMP_GetPreAmpValue() As Long

    '----------------------------------'
    'Retrieves the Pre-amp slider value'
    '----------------------------------'

    Dim ReturnValue As Long
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Function
    End If
    
    'values will range from 0(top) to 63(bottom)
    ReturnValue = SendMessage(hWndWinAMP, WM_USER, 10, WA_GETEQDATA)

    WinAMP_GetPreAmpValue = ReturnValue

End Function
Sub WinAMP_SeekToPosition(PositionInSec As Long)

    '----------------------------------------------------'
    'Seeks to the specified position in the current track'
    '----------------------------------------------------'

    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Sub
    End If

    'range check the new position
    If PositionInSec < 0 Or PositionInSec > WinAMP_GetTrackLength Then
       Exit Sub
    End If

    SendMessage hWndWinAMP, WM_USER, CLng(PositionInSec * 1000), WA_SEEKTOPOSITION

End Sub
Sub WinAMP_SetVolume(VolumeValue As Long)

    '----------------------'
    'Sets the volume slider'
    '----------------------'

    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Sub
    End If

    'volume values will range from 0(bottom) to 255(top)

    'range check the new volume
    If VolumeValue < 0 Or VolumeValue > 255 Then
       Exit Sub
    End If

    SendMessage hWndWinAMP, WM_USER, VolumeValue, WA_SETVOLUME

End Sub

Sub WinAMP_SetBalance(BalanceValue As Long)

    '-----------------------'
    'Sets the Balance slider'
    '-----------------------'

    Dim ScaledBalanceValue As Long

    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Sub
    End If

    'we want the range of balance to be from -127 to 127, but the
    'real range is as follows
    
    '|LEFT                       CENTER                       RIGHT|
    '|128 ----------------------255(or 0)-----------------------127|
    '|_____________________________________________________________|
    
    'range check the new balance
    If BalanceValue < -127 Or BalanceValue > 127 Then
       Exit Sub
    End If
    
    'here we do our shifting to correct the range of values
    If BalanceValue < 0 Then
       ScaledBalanceValue = 255 + BalanceValue
    Else
       ScaledBalanceValue = BalanceValue
    End If

    SendMessage hWndWinAMP, WM_USER, ScaledBalanceValue, WA_SETBALANCE

End Sub
Function WinAMP_GetTrackLength() As Long

    '-------------------------------------------------'
    'Retrieves the length of the current track in secs'
    '-------------------------------------------------'

    Dim ReturnLength As Long
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Function
    End If

    'ReturnLength will contain the current track length in seconds
    'or -1 if no track is playing or an error occurs
    ReturnLength = SendMessage(hWndWinAMP, WM_USER, 1, WA_GETTRACKLENGTH)

    If ReturnLength <> -1 Then
       WinAMP_GetTrackLength = ReturnLength
    Else
       WinAMP_GetTrackLength = -1
    End If

End Function

Function WinAMP_GetVersion() As String

    '---------------------------------------'
    'Retrieves the version of WinAMP running'
    '---------------------------------------'

    Dim VersionNum As Long
    Dim ReturnVersion As String

    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Function
    End If

    VersionNum = SendMessage(hWndWinAMP, WM_USER, 0, WA_GETVERSION)
    
    If Len(Hex(VersionNum)) > 3 Then
       ReturnVersion = Left(Hex(VersionNum), 1) & "."
       ReturnVersion = ReturnVersion & Mid(Hex(VersionNum), 2, 1)
       ReturnVersion = ReturnVersion & Right$(Hex(VersionNum), Len(Hex(VersionNum)) - 3)
       WinAMP_GetVersion = ReturnVersion
    Else
       WinAMP_GetVersion = "UNKNOWN"
    End If

End Function

Sub WinAMP_SendCommandMessage(CommandMessage As Long)
    
    '--------------------------------------------------'
    'Used to send any of the Command messages to WinAMP'
    '--------------------------------------------------'
    
    If hWndWinAMP = 0 Then
       MsgBox "WinAMP window not found yet...", vbOKOnly + vbCritical, "WinAMP Not Found"
       Exit Sub
    End If
    
    SendMessage hWndWinAMP, WM_COMMAND, CommandMessage, 0

End Sub


Public Function WinAMP_Start() As Boolean

    '---------------------------------------------------------------'
    'Runs an instance of WinAMP if an instance isn't already running'
    '---------------------------------------------------------------'

    Dim ReturnValue As Double
    
    ReturnValue = Shell(WINAMP_PATH, vbMinimizedNoFocus)
    
    If ReturnValue <> 0 Then
       WinAMP_Start = True
    Else
       WinAMP_Start = False
    End If

End Function


Public Function WinAMP_OpenFile(strFilename As String) As Boolean

    '----------------------------------------------------'
    'Causes WinAMP to open the file specified by filename'
    '----------------------------------------------------'

    Dim ReturnValue As Double
    
    If FileExists(strFilename) Then
       ReturnValue = Shell(WINAMP_PATH & " /ADD " & vbQuote & strFilename & vbQuote, vbMinimizedNoFocus)
    
       If ReturnValue <> 0 Then
          WinAMP_OpenFile = True
       Else
          WinAMP_OpenFile = False
       End If
    Else
       WinAMP_OpenFile = False
    End If

End Function
Sub YahooSend(txt As String)
Dim imclass As Long
Dim Rich1   As Long
Dim Rich2   As Long
Dim Button  As Long
imclass = FindWindow("imclass", vbNullString)
Rich1& = FindWindowEx(imclass, 1, "RICHEDIT", vbNullString)
Rich2& = FindWindowEx(imclass, Rich1&, "RICHEDIT", vbNullString)
Call SendMessageByString(Rich2&, WM_SETTEXT, 1, txt$)
Button = FindWindowEx(imclass, 0&, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Button = FindWindowEx(imclass, Button, "button", vbNullString)
Call SendMessageLong(Button, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)
If Button = 0 Then
Exit Sub
End If
End Sub
