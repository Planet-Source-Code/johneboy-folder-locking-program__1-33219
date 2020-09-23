VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuke SFS"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "NukeSFS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "NukeIt"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\DefaultIcon"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\DefaultIcon"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\InProcServer32"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\InProcServer32"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentAddinsRegistered\{89BCB740-6119-101A-BCB7-00DD010655AF}"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentAddinsRegistered\{89BCB740-6119-101A-BCB7-00DD010655AF}"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentAddinsRegistered"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentAddinsRegistered"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentHandler"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\PersistentHandler"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\ProgID"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\ProgID"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\shellex\MayChangeDefaultMenu"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\shellex\MayChangeDefaultMenu"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\shellex"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}\shellex"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}"
List1.AddItem "HKEY_CLASSES_ROOT\CLSID\{00021401-0000-0000-C000-000000000045}"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder\DefaultIcon"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder\DefaultIcon"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock\command"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock\command"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder\Shell\Unlock"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder\Shell"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder\Shell"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers\{00021401-0000-0000-C000-000000000045}"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers\{00021401-0000-0000-C000-000000000045}"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder\shellex\ContextMenuHandlers"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder\shellex"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder\shellex"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\LockedFolder"
List1.AddItem "HKEY_CLASSES_ROOT\LockedFolder"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\Directory\Shell\Lock\Command"
List1.AddItem "HKEY_CLASSES_ROOT\Directory\Shell\Lock\Command"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CLASSES_ROOT\Directory\Shell\Lock"
List1.AddItem "HKEY_CLASSES_ROOT\Directory\Shell\Lock"
List1.AddItem "REMOVED"
DeleteKey "HKEY_CURRENT_USER\Software\SFS"
List1.AddItem "HKEY_CURRENT_USER\Software\SFS"
List1.AddItem "REMOVED"


MsgBox "SFS has been nuked"
Unload Me
End Sub

