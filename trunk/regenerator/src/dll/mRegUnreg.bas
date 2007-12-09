Attribute VB_Name = "mRegUnreg"
' Daisy 2.02 Regenerator DLL
' Copyright (C) 2003 Daisy Consortium
'
'    This file is part of Daisy 2.02 Regenerator.
'
'    Daisy 2.02 Regenerator is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Daisy 2.02 Regenerator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Daisy 2.02 Regenerator; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


Option Explicit
 
'All required Win32 SDK functions to register/unregister any ActiveX component

Private Declare Function LoadLibraryRegister Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibraryRegister Lib "KERNEL32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddressRegister Lib "KERNEL32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "KERNEL32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "KERNEL32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "KERNEL32" (ByVal dwExitCode As Long)
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Public Enum REGISTER_FUNCTIONS
    DllRegisterServer = 1
    DllUnRegisterServer = 2
End Enum
Public Enum STATUS
    [File Could Not Be Loaded Into Memory Space] = 1
    [Not A Valid ActiveX Component] = 2
    [ActiveX Component Registration Failed] = 3
    [ActiveX Component Registered Successfully] = 4
    [ActiveX Component UnRegistered Successfully] = 5
End Enum


Public Sub registerDll(strTheDll)

    Dim menum As STATUS

    menum = RegisterComponent(strTheDll, DllRegisterServer)

    If menum = [File Could Not Be Loaded Into Memory Space] Then

'        objOwner.addlog (strTheDll & " Could Not Be Loaded Into Memory Space")

    ElseIf menum = [Not A Valid ActiveX Component] Then

'        objOwner.addlog (strTheDll & "Not A Valid ActiveX Component")

    ElseIf menum = [ActiveX Component Registration Failed] Then

'        objOwner.addlog (strTheDll & "ActiveX Component Registration Failed")

    ElseIf menum = [ActiveX Component Registered Successfully] Then

        'objowner.addlog (strTheDll & " registered")

    End If

End Sub



Public Function RegisterComponent(ByVal FileName$, ByVal RegFunction As REGISTER_FUNCTIONS) As STATUS

    Dim lngLib&, lngProcAddress&, lpThreadID&, fSuccess&, dwExitCode&, hThread&

    If FileName = "" Then Exit Function

    lngLib = LoadLibraryRegister(FileName)

    If lngLib = 0 Then

        RegisterComponent = [File Could Not Be Loaded Into Memory Space]    'Couldn't load component
        Exit Function

    End If

    Select Case RegFunction

        Case REGISTER_FUNCTIONS.DllRegisterServer
            lngProcAddress = GetProcAddressRegister(lngLib, "DllRegisterServer")

        Case REGISTER_FUNCTIONS.DllUnRegisterServer
            lngProcAddress = GetProcAddressRegister(lngLib, "DllUnregisterServer")

        Case Else

    End Select

    If lngProcAddress = 0 Then

        RegisterComponent = [Not A Valid ActiveX Component]               'Not a Valid ActiveX Component

        If lngLib Then Call FreeLibraryRegister(lngLib)

        Exit Function

    Else

        hThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lngProcAddress, ByVal 0&, 0&, lpThreadID)

        If hThread Then

            fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)

            If Not fSuccess Then

                Call GetExitCodeThread(hThread, dwExitCode)
                Call ExitThread(dwExitCode)
                RegisterComponent = [ActiveX Component Registration Failed]        'Couldn't Register.

                If lngLib Then Call FreeLibraryRegister(lngLib)

                Exit Function

            Else

                If RegFunction = DllRegisterServer Then

                    RegisterComponent = [ActiveX Component Registered Successfully]         'Success. OK

                ElseIf RegFunction = DllUnRegisterServer Then

                    RegisterComponent = [ActiveX Component UnRegistered Successfully]         'Success. OK

                End If

            End If

            Call CloseHandle(hThread)

            If lngLib Then Call FreeLibraryRegister(lngLib)

        End If

    End If

End Function

