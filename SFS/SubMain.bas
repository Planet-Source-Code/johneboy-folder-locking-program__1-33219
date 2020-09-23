Attribute VB_Name = "SubMain"
Option Explicit
Public Sub Main()
Dim EXEpath As String
Dim Path As String
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\DefaultIcon"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\InProcServer32"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentAddinsRegistered"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentAddinsRegistered\{89BCB740-6119-101A-BCB7-00DD010655AF}"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentHandler"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\ProgID"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\shellex"
CreateKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\shellex\MayChangeDefaultMenu"
EXEpath = "" + Chr(34) + "" + App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\DefaultIcon", "", "" + App.Path + "\" + App.EXEName + ".exe ,0"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\InProcServer32", "", "shell32.dll"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\InProcServer32", "ThreadingModel", "Apartment"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentAddinsRegistered\{89BCB740-6119-101A-BCB7-00DD010655AF}", "", "{00021401-0000-0000-C000-000000000045}"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentHandler", "", "{00021401-0000-0000-C000-000000000045}"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\ProgID", "", "LockedFolder"

CreateKey "HKEY_CLASSES_ROOT\LockedFolder"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\DefaultIcon"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\Shell"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock\command"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\shellex"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers"
CreateKey "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers\{00021401-0000-0000-C000-000000000045}"

SetStringValue "HKEY_CLASSES_ROOT\LockedFolder\DefaultIcon", "", "" + App.Path + "\" + App.EXEName + ".exe ,0"
SetStringValue "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock\command", "", "" & EXEpath

CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\Lock"
CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\Lock\Command"
SetStringValue "HKEY_CLASSES_ROOT\Directory\Shell\Lock\Command", "", "" & EXEpath

If FirstRun = "No" Then

    Path = Command
        If Path = "" Then
        FrmHowTo.Show
        Else
        If Right(Path, 1) <> "}" Then
        FrmLock.Show
        Else
        FrmUnLock.Show
        End If
        End If

Else
CreateKey "HKEY_CURRENT_USER\Software\SFS"
SetStringValue "HKEY_CURRENT_USER\Software\SFS", "FirstRun", "No"
FrmSetPass.Show
End If
End Sub

