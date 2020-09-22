Attribute VB_Name = "const_mod"
Option Explicit

' -------------------- '
' constanta key fungsi '
' -------------------- '
Public Const MostUsedkey = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

' key fungsi visual1 '
Public Const vis_f_0_kiri = "HKEY_CURRENT_USER\Control Panel\Desktop"
Public Const vis_f_1_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis_f_2_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis_f_3_kiri = "HKEY_CURRENT_USER\Control Panel\Desktop"
Public Const vis_f_4_kiri = "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics"
Public Const vis_f_5_kiri = "HKEY_CURRENT_USER\Control Panel\Desktop"
Public Const vis_f_7_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\CleanupWiz"
Public Const vis_f_9_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis_f_10_kiri = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction"
Public Const vis_f_11_kiri = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer"
Public Const vis_f_12_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis_f_13_kiri = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PCHealth\ErrorReporting"
Public Const vis_f_0_kanan = "DragFullWindows"
Public Const vis_f_1_kanan = "ListviewShadow"
Public Const vis_f_2_kanan = "WebView"
Public Const vis_f_3_kanan = "PaintDesktopVersion"
Public Const vis_f_4_kanan = "MinAnimate"
Public Const vis_f_5_kanan = "MenuShowDelay"
Public Const vis_f_6_kanan = "NoLowDiskSpaceChecks"
Public Const vis_f_7_kanan = "NoRun"
Public Const vis_f_8_kanan = "NoDesktop"
Public Const vis_f_9_kanan = "ListviewWatermark"
Public Const vis_f_10_kanan = "Enable"
Public Const vis_f_11_kanan = "NoSaveSettings"
Public Const vis_f_12_kanan = "EnableBalloonTips"
Public Const vis_f_13_kanan = "DoReport"

' key fungsi visual2 '
Public Const vis2_f_0_kiri = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer"
Public Const vis2_f_1_kiri = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\CrashControl"
Public Const vis2_f_2_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis2_f_3_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis2_f_4_kiri = "HKEY_CURRENT_USER\Control Panel\Mouse"
Public Const vis2_f_6_kiri = "HKEY_CURRENT_USER\Control Panel\Desktop"
Public Const vis2_f_7_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Tour"
Public Const vis2_f_8_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis2_f_9_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis2_f_10_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const vis2_f_11_kiri = "HKEY_USERS\.DEFAULT\Control Panel\Keyboard"
Public Const vis2_f_12_kiri = "HKEY_CURRENT_USER\Control Panel\Mouse"
Public Const vis2_f_13_kiri = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem"
Public Const vis2_f_0_kanan = "NoTrayItemsDisplay"
Public Const vis2_f_1_kanan = "AutoReboot"
Public Const vis2_f_2_kanan = "Start_LargeMFUIcons"
Public Const vis2_f_3_kanan = "IntelliMenus"
Public Const vis2_f_4_kanan = "SwapMouseButtons"
Public Const vis2_f_5_kanan = "NoDriveTypeAutoRun"
Public Const vis2_f_6_kanan = "AutoEndTasks"
Public Const vis2_f_7_kanan = "RunCount"
Public Const vis2_f_8_kanan = "DisableThumbnailCache"
Public Const vis2_f_9_kanan = "ShowInfoTip"
Public Const vis2_f_10_kanan = "Start_ScrollPrograms"
Public Const vis2_f_11_kanan = "InitialKeyboardIndicators"
Public Const vis2_f_12_kanan = "MouseTrails"
Public Const vis2_f_13_kanan = "NtfsDisableLastAccessUpdate"

' key fungsi security1 '
Public Const sec_f_0_kiri = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
Public Const sec_f_2_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
Public Const sec_f_3_1_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const sec_f_6_0_kiri = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Control Panel\Desktop"
Public Const sec_f_6_1_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
Public Const sec_f_8_kiri = "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\System"
Public Const sec_f_9_kiri = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System"
Public Const sec_f_10_kiri = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
Public Const sec_f_11_kiri = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices"
Public Const sec_f_12_kiri = "HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\Installer"
Public Const sec_f_13_kiri = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Public Const sec_f_0_kanan = "DisableTaskMgr"
Public Const sec_f_1_kanan = "NoViewContextMenu"
Public Const sec_f_2_kanan = "NoDispCPL"
Public Const sec_f_3_0_kanan = "NoRecentDocsHistory"
Public Const sec_f_3_1_kanan = "Start_ShowRecentDocs"
Public Const sec_f_4_kanan = "NoFolderOptions"
Public Const sec_f_5_kanan = "NoPropertiesMyComputer"
Public Const sec_f_6_0_kanan = "ScreenSaveActive"
Public Const sec_f_6_1_kanan = "NoDispScrSavPage"
Public Const sec_f_7_kanan = "NoRun"
Public Const sec_f_8_kanan = "DisableGPO"
Public Const sec_f_9_kanan = "DisableCMD"
Public Const sec_f_10_kanan = "DisableRegistryTools"
Public Const sec_f_11_kanan = "SchedulingAgent"
Public Const sec_f_12_kanan = "DisableMSI"
Public Const sec_f_13_kanan = "EncryptionContextMenu"

' key fungsi security2 '
Public Const sec2_f_9_kiri = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
Public Const sec2_f_10_kiri = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Cache"
Public Const sec2_f_12_0_kiri = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore"
Public Const sec2_f_0_kanan = "NoWinKeys"
Public Const sec2_f_1_kanan = "NoChangeStartMenu"
Public Const sec2_f_2_kanan = "NoControlPanel"
Public Const sec2_f_3_kanan = "NoSetTaskbar"
Public Const sec2_f_4_kanan = "NoAddPrinter"
Public Const sec2_f_5_kanan = "NoSecurityTab"
Public Const sec2_f_6_kanan = "NoToolbarCustomize"
Public Const sec2_f_7_kanan = "NoFileMenu"
Public Const sec2_f_8_kanan = "StartMenuLogoff"
Public Const sec2_f_9_kanan = "ClearPageFileAtShutdown"
Public Const sec2_f_10_kanan = "Persistent"
Public Const sec2_f_11_kanan = "NoInstrumentation"
Public Const sec2_f_12_0_kanan = "DisableSR"
Public Const sec2_f_12_1_kanan = "DisableConfig"
Public Const sec2_f_13_kanan = "NoSimpleStartMenu"

' ------------------------ '
' constanta system command '
' ------------------------ '
Public Const shutdown_ = "shutdown -s -f -t 0"
Public Const restart_ = "shutdown -r -f -t 0"
Public Const logoff_ = "logoff"

' ----------------------------------- '
' constanta system properties command '
' ----------------------------------- '
Public Const sys_prop_0 = "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0"
Public Const sys_prop_1 = "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,2"
Public Const sys_prop_2 = "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,3"
Public Const sys_prop_3 = "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,1"
Public Const sys_prop_4 = "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,4"
Public Const sys_prop_5 = "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,5"
Public Const sys_prop_6 = "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,6"

' -------------------------------- '
' constanta system general command '
' -------------------------------- '
Public Const sys_gnrl_0 = "rundll32.exe shell32.dll,Control_RunDLL main.cpl @1"
Public Const sys_gnrl_1 = "rundll32.exe shell32.dll,Control_RunDLL main.cpl @0"
Public Const sys_gnrl_2 = "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0"
Public Const sys_gnrl_3 = "rundll32.exe shell32.dll,Control_RunDLL modem.cpl"
Public Const sys_gnrl_4 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,5"
Public Const sys_gnrl_5 = "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1"
Public Const sys_gnrl_6 = "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"
Public Const sys_gnrl_7 = "rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0"
Public Const sys_gnrl_8 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0"
Public Const sys_gnrl_9 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,1"
Public Const sys_gnrl_10 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2"
Public Const sys_gnrl_11 = "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3"
Public Const sys_gnrl_12 = "rundll32.exe shell32.dll,Control_RunDLL access.cpl"
Public Const sys_gnrl_13 = "rundll32.exe shell32.dll,Control_RunDLL odbccp32.cpl"
Public Const sys_gnrl_14 = "rundll32.exe shell32.dll,Control_RunDLL irprops.cpl"
Public Const sys_gnrl_15 = "rundll32.exe shell32.dll,Control_RunDLL joy.cpl"
Public Const sys_gnrl_16 = "rundll32.exe shell32.dll,Control_RunDLL nusrmgr.cpl"
Public Const sys_gnrl_17 = "rundll32.exe shell32.dll,Control_RunDLL powercfg.cpl"
Public Const sys_gnrl_18 = "rundll32.exe shell32.dll,Control_RunDLL hdwwiz.cpl"
Public Const sys_gnrl_19 = "rundll32.exe shell32.dll,Control_RunDLL ncpa.cpl"

' ------------------------------ '
' constanta system tools command '
' ------------------------------ '
Public Const sys_tools_0 = "taskmgr"
Public Const sys_tools_1 = "gpedit.msc"
Public Const sys_tools_2 = "cmd"
Public Const sys_tools_3 = "rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,0"
Public Const sys_tools_4 = "msconfig"
Public Const sys_tools_5 = "regedit"
Public Const sys_tools_6 = "dxdiag"

' --------------------------- '
' constanta key startup entry '
' --------------------------- '
Public Const startup_entry_1 = "Software\Microsoft\Windows\CurrentVersion\Run"
Public Const startup_entry_1_ = "Software\Microsoft\Windows\CurrentVersion\Run-"
Public Const startup_entry_2 = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
Public Const startup_entry_2_ = "Software\Microsoft\Windows\CurrentVersion\RunOnce-"
Public Const startup_entry_3 = "Software\Microsoft\Windows\CurrentVersion\RunServices"
Public Const startup_entry_3_ = "Software\Microsoft\Windows\CurrentVersion\RunServices-"
Public Const startup_capt_0 = "Current User  Run"
Public Const startup_capt_1 = "Current User  Run Once"
Public Const startup_capt_2 = "Local Machine  Run"
Public Const startup_capt_3 = "Local Machine  Run Once"
Public Const startup_capt_4 = "Local Machine  Run Services"

' ----------------------------------------- '
' constanta dan enum bwt baca startup entry '
' ----------------------------------------- '
Enum RegistryKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum
Enum RegDataTypes
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_DWORD = 4
End Enum
Enum ValKey
    Values = 0
    Keys = 1
End Enum
Enum ref_semua_ato_tidak
    Tidak = 0
    Semua = 1
End Enum
Enum Folders
    Desktop = &H0
    Internet = &H1
    Programs = &H2
    ControlsFolder = &H3
    Printers = &H4
    Personal = &H5
    Favorites = &H6
    StartUp = &H7
    Recent = &H8
    SendTo = &H9
    RecycleBin = &HA
    StartMenu = &HB
    DesktopDirectory = &H10
    Drives = &H11
    Network = &H12
    Nethood = &H13
    Fonts = &H14
    Templates = &H15
    Common_StartMenu = &H16
    Common_Programs = &H17
    Common_StartUp = &H18
    Common_DesktopDirectory = &H19
    ApplicationData = &H1A
    PrintHood = &H1B
    AltStartUp = &H1D
    Common_AltStartUp = &H1E
    Common_Favorites = &H1F
    InternetCache = &H20
    Cookies = &H21
    History = &H22
End Enum
Public Const S_OK = &H0
Public Const S_FALSE = &H1
Public Const E_INVALIDARG = &H80070057
Public Const CSIDL_LOCAL_APPDATA = &H1C&
Public Const CSIDL_FLAG_CREATE = &H8000&
Public Const SHGFP_TYPE_CURRENT = 0
Public Const SHGFP_TYPE_DEFAULT = 1
Public Const MAX_PATH = 260
Public Const KEY_QUERY_VALUE = &H1
