Attribute VB_Name = "SpoolIcon"
'Module created 10/5/2001
'api
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As TRAYCON) As Long

Public Type TRAYCON
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    sTip As String * 64
    End Type
    
    'contants
    Public Const IC_ADD = &H0
    Public Const IC_MODIFY = &H1
    Public Const IC_DELETE = &H2
    Public Const IC_MESSAGE = &H1
    Public Const IC_ICON = &H2
    Public Const IC_TIP = &H4
    Public Const IC_DOALL = IC_MESSAGE Or IC_ICON Or IC_TIP
    Public Const IC_RESTORE = 9
    Public Const IC_MINIMIZE = 6
    Public Const IC_MOUSEMOVE = &H200
    Public Const IC_LBUTTONDCLK = &H201
    Public Const IC_LBUTTONDBLCLK = &H203
    Public Const IC_RBUTTONUP = &H205

Public IC As TRAYCON
Dim DC As Long
Public Sub AddIcon()
IC.cbSize = Len(IC)
                IC.hwnd = FrmSmtp.hwnd                      'set form handle as systray handle
                IC.uFlags = IC_DOALL                       'do all features
                IC.uCallbackMessage = IC_MOUSEMOVE         'call back on mouse move
                IC.hIcon = FrmMain.Picture1(0).Picture     'default icon
                IC.sTip = "Hattech Mail Server" & vbNullChar 'change tooltip
                DC = Shell_NotifyIcon(IC_ADD, IC)          'adds icon to tray
                Beep
              
End Sub
Public Sub RemoveIcon() 'removes icon from tray
    DC = Shell_NotifyIcon(IC_DELETE, IC)
End Sub
Public Sub ChangeIcon(IconImage As PictureBox) 'usage: changeicon (picture1.picture)
             IC.hIcon = IconImage
             DC = Shell_NotifyIcon(IC_MODIFY, IC)
            

End Sub

