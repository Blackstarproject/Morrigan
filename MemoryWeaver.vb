Imports System.Runtime.InteropServices
Imports System.Text

Public Class MemoryWeaver

    Inherits Morrigan

    'VB.NET Memory and Process Hacking Class:
    'Features: Implementations for OpenProcess, NtReadVirtualMemory & NtWriteVirtualMemory plus easy to add to any VB.NET project
    'Functions:
    'GetOffsetByName (for reading csgo offsets from a string based on the name)
    'GetProcessByName (obtain a Process object from the process name)
    'GetHandle (opens a handle to an external process with PROCESS_ALL_ACCESS flag)
    'GetModuleBase (gets a base dll of a certain process module)
    'ReadProcessMemory (int, float, bool, vec2, vec3, viewmatrix)
    'WriteProcessMemory (int, float, bool)

#Region " Matrix Structures "

    Public Structure W2SMatrix
        Dim f00 As Single
        Dim f01 As Single
        Dim f02 As Single
        Dim f03 As Single
        Dim f10 As Single
        Dim f11 As Single
        Dim f12 As Single
        Dim f13 As Single
        Dim f20 As Single
        Dim f21 As Single
        Dim f22 As Single
        Dim f23 As Single
        Dim f30 As Single
        Dim f31 As Single
        Dim f32 As Single
        Dim f33 As Single
    End Structure

    Public Structure fVec2
        Dim x As Single
        Dim y As Single
    End Structure

    Public Structure fVec3
        Dim x As Single
        Dim y As Single
        Dim z As Single
    End Structure

#End Region

#Region "WindowsAPI"

    Public Const PROCESS_ALL_ACCESS = &H1F0FF
    Private Declare Function OpenProcess Lib "kernel32" (dwDesiredAccess As Integer, bInheritHandle As Integer, dwProcessId As Integer) As IntPtr
    Private Declare Function WriteMemory Lib "ntdll" Alias "NtWriteVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As Integer, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function WriteMemoryF Lib "ntdll" Alias "NtWriteVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As Single, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function WriteMemoryB Lib "ntdll" Alias "NtWriteVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As Boolean, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function ReadMemory Lib "ntdll" Alias "NtReadVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As Integer, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function ReadMemoryF Lib "ntdll" Alias "NtReadVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As Single, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function ReadMemoryB Lib "ntdll" Alias "NtReadVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As Boolean, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function ReadMemoryViewMatrix Lib "ntdll" Alias "NtReadVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As W2SMatrix, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function ReadMemoryString Lib "ntdll" Alias "NtReadVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, <Out()> u32Buffer As Byte(), nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function ReadMemoryFVec2 Lib "ntdll" Alias "NtReadVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As fVec2, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Private Declare Function ReadMemoryFVec3 Lib "ntdll" Alias "NtReadVirtualMemory" (hProcess As IntPtr, lpBaseAddress As Integer, ByRef lpBuffer As fVec3, nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Boolean
    Public Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (hobject As IntPtr) As Boolean

#End Region

#Region " Get Functions "

    Public Shared Function GetOffsetByName(data As String, offset As String) As Integer
        Try
            Dim pos As Integer = data.IndexOf(offset) + offset.Length + 5
            Dim s1 As String = data.Substring(pos, 20)
            Dim s2() As String = s1.Split(";")
            Dim x As Integer = Convert.ToInt32(s2(0), 16)
            Return x
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Shared Function GetProcessByName(processname) As Process
        Dim p As Process() = Process.GetProcessesByName(processname)
        If p.Length > 0 Then
            Return p.FirstOrDefault
        End If
        Return Nothing
    End Function

    Public Shared Function GetHandle(p As Process) As IntPtr
        Try
            Return OpenProcess(PROCESS_ALL_ACCESS, 0, p.Id)
        Catch ex As Exception
            Return IntPtr.Zero
        End Try
    End Function

    Public Shared Function GetModuleBase(p As Process, modulename As String) As Integer
        Try
            Dim base As Integer = 0
            For Each m As ProcessModule In p.Modules
                If m.ModuleName = modulename Then
                    base = m.BaseAddress
                End If
            Next
            Return base
        Catch ex As Exception
            Return 0
        End Try
    End Function

#End Region

#Region " Read | Write Infrastructure "

    Public Shared Function RPMInt(hProcess As IntPtr, address As Integer) As Integer
        Dim buffer As Integer
        ReadMemory(hProcess, address, buffer, 4, 0)
        Return buffer
    End Function

    Public Shared Function RPMFloat(hProcess As IntPtr, address As Integer) As Single
        Dim buffer As Single
        ReadMemoryF(hProcess, address, buffer, 4, 0)
        Return buffer
    End Function

    Public Shared Function RPMBool(hProcess As IntPtr, address As Integer) As Boolean
        Dim buffer As Boolean
        ReadMemoryB(hProcess, address, buffer, 1, 0)
        Return buffer
    End Function

    Public Shared Function RPMViewMatrix(hProcess As IntPtr, address As Integer) As W2SMatrix
        Dim buffer As W2SMatrix
        ReadMemoryViewMatrix(hProcess, address, buffer, 64, 0)
        Return buffer
    End Function

    Public Shared Function RPMString(hProcess As IntPtr, address As Integer, stringSize As Integer) As String
        Dim buffer(stringSize) As Byte
        ReadMemoryString(hProcess, address, buffer, stringSize, 0)
        Return Encoding.UTF8.GetString(buffer)
    End Function

    Public Shared Function RPMFVec2(hProcess As IntPtr, address As Integer) As fVec2
        Dim buffer As fVec2
        ReadMemoryFVec2(hProcess, address, buffer, 8, 0)
        Return buffer
    End Function

    Public Shared Function RPMFVec3(hProcess As IntPtr, address As Integer) As fVec3
        Dim buffer As fVec3
        ReadMemoryFVec3(hProcess, address, buffer, 12, 0)
        Return buffer
    End Function

    Public Shared Function WPMInt(hProcess As IntPtr, address As Integer, value As Integer) As Boolean
        If WriteMemory(hProcess, address, value, 4, 0) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function WPMFloat(hProcess As IntPtr, address As Integer, value As Single) As Boolean
        If WriteMemoryF(hProcess, address, value, 4, 0) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function WPMBool(hProcess As IntPtr, address As Integer, value As Boolean) As Boolean
        If WriteMemoryB(hProcess, address, value, 1, 0) Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

End Class
