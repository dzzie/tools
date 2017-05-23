Attribute VB_Name = "Startup"

Public Enum listStyle
    inbox = 1
    outbx = 2
    trash = 3
    saved = 4
End Enum

Public Type parsed
   body As String
   from As String
   subj As String
   atch As String
   to As String
End Type

Public Type profile
    user As String
    pass As String
    Server As String
    port As String
End Type

Public Type font
    size As Integer
    color As String
    face As String
    bold As Boolean
    backcolor As String
End Type

Public Type Prefs
    saveSent As Boolean
    MsgInReply As Boolean
    ReplyChar As String
    useTrash As Boolean
    AutoCheckDelay As Long
End Type

Public Type sendInfo
    sender As String
    Server As String
End Type

Public Type folders
    inbox As String
    oubox As String
    trash As String
    saved As String
    attach As String
    sigFile As String
    browser As String
    editor As String
    saveTo As String
    bookmark As String
    IniFile As String
End Type

'------------------------------------------------------------
Public Type Config           'global configuration object
    Users() As profile       'composed of sub objects
    Prefs As Prefs
    Send As sendInfo
    xHeaders() As String
    recipants() As String
    bitfiles() As String
    folders As folders
    fonts As font
    SentMessages As Integer
End Type

'-------------------------------------------------------------
Global uc As Config      'user config
Global isFirstRun As Boolean
'--------------------------------------------------------------

'--------------------------------------------------------------
'--               Main Entry Point into Program              --
'--------------------------------------------------------------
Sub Main()
    Call setConfigPaths
    Call loadConfig
    Call b64.InitAlpha
    frmMessages.Show
    If isFirstRun Then Shell "notepad """ & App.path & "\readme.txt""", vbNormalFocus
    If Len(Command) > 0 Then Call Library.ParseCommandLine(Command)
End Sub

Public Sub setConfigPaths()
      mainConfig = App.path & "\qmail.ini"
      backupConfig = App.path & "\demo copy qmail.ini"
      
      If fso.FileExists(mainConfig) Then
            uc.folders.IniFile = mainConfig
            Call Ini.LoadFile(CStr(mainConfig))
      ElseIf fso.FileExists(backupConfig) Then
            MsgBox "Warning ! Loading Using demo copy of ini file!" & vbCrLf & "Make sure to rename a copy of the demo file qmail.ini and set with your info." & vbCrLf & "Choose Options from main right click menu to edit ini file. You can also add frmOptions to the project to configure most options, but it is not 100% complete", vbExclamation
            uc.folders.IniFile = backupConfig
            Call Ini.LoadFile(CStr(backupConfig))
            isFirstRun = True
      Else
         MsgBox "Sorry cant find the INI config file..exiting", vbCritical
         End
      End If
      
End Sub

Public Sub loadConfig()
            
      ReDim uc.Users(1)
      ReDim uc.bitfiles(0)
      ReDim uc.recipants(0)
      ReDim uc.xHeaders(0)
      
      n = Ini.GetValue("profile", "number")
      For j = 1 To n
        ub = UBound(uc.Users)
        With uc.Users(ub)
            .Server = Ini.GetValue("profile", "server" & j)
            .port = Ini.GetValue("profile", "port" & j)
            .pass = Ini.GetValue("profile", "pass" & j)
            .user = Ini.GetValue("profile", "user" & j)
        End With
        ReDim Preserve uc.Users(ub + 1)
      Next
            
      With uc.Send
        .sender = Ini.GetValue("sending", "from")
        .Server = Ini.GetValue("sending", "server")
      End With
           
      n = Ini.GetValue("recipients", "number")
      For j = 1 To n
          push uc.recipants, Ini.GetValue("recipients", "recp" & j)
      Next
            
      n = Ini.GetValue("bitfiles", "number")
      For j = 1 To n
         push uc.bitfiles, Ini.GetValue("bitfiles", "bitfile" & j)
      Next
      push uc.bitfiles, "- Add New -"
         
      With uc.Prefs
         .saveSent = Ini.GetValue("preferences", "saveSent")
         .MsgInReply = Ini.GetValue("preferences", "MsgInReply")
         .ReplyChar = Ini.GetValue("preferences", "ReplyChar")
         .useTrash = Ini.GetValue("preferences", "usetrash")
         .AutoCheckDelay = Ini.GetValue("preferences", "AutoCheckDelay")
      End With
            
      n = Ini.GetValue("x-headers", "number")
      For j = 1 To n
         push uc.xHeaders, Ini.GetValue("x-headers", "header" & j)
      Next
      
      With uc.folders
          .attach = Ini.GetValue("folders", "attach")
          .inbox = Ini.GetValue("folders", "inbox")
          .oubox = Ini.GetValue("folders", "outbox")
          .trash = Ini.GetValue("folders", "trash")
          .saved = Ini.GetValue("folders", "saved")
          .saveTo = Ini.GetValue("folders", "saveto")
          .sigFile = Ini.GetValue("folders", "sigfile")
          .bookmark = Ini.GetValue("folders", "bookmarks")
          .browser = """" & Ini.GetValue("folders", "browser") & """"
          .editor = """" & Ini.GetValue("folders", "editor") & """"
      End With
            
      With uc.fonts
          .color = Ini.GetValue("fonts", "msgcolor")
          .face = Ini.GetValue("fonts", "font")
          .size = Ini.GetValue("fonts", "size")
          .bold = Ini.GetValue("fonts", "bold")
          .backcolor = Ini.GetValue("fonts", "backcolor")
      End With
       
      Ini.Release
End Sub

Sub reloadConfig()
    Call setConfigPaths
    Call loadConfig
    frmMessages.loadDynamicMenus
End Sub
