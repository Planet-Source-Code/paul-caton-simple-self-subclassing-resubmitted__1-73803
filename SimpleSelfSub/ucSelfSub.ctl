VERSION 5.00
Begin VB.UserControl ucSelfSub 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   ScaleHeight     =   2025
   ScaleWidth      =   3240
End
Attribute VB_Name = "ucSelfSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'____________________________________
' Simple, self-sublassed usercontrol
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Option Explicit

Private Const WM_MOUSEMOVE As Long = &H200

Private Type MOUSE_XY
    x As Integer
    y As Integer
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'SelfSub declarations_____________________________________________________________________________________________________________________________________________________________
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, ByRef RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, ByRef RetVal As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)
Private z_hWnd  As Long
Private z_scMem As Long
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        sc_Subclass
    End If
End Sub

Private Sub UserControl_Terminate()
    sc_Terminate
End Sub

'SelfSub code______________________________________________________________
Private Function sc_Subclass(Optional ByRef RefData As Long = 0) As Boolean
    Dim v, n, nAddr As Long
    Dim b, m        As Byte
    
    GetMem4 ObjPtr(Me), nAddr                       'get address of the usercontrol's vtable
    nAddr = nAddr + &H7A4                           'bump to the user part of the usercontrol's vtable
    GetMem4 nAddr, n                                'read the address of the first entry point
    GetMem1 n, m                                    'read the jump opcode at the first entry point [&H33 for psuedo code, &HE9 for native code]
    For v = 1 To 512                                'scan a reasonable number of vtable entries
        nAddr = nAddr + 4                           'next entry address
        GetMem4 nAddr, n                            'read the address of the entry point
        If IsBadCodePtr(n) Then GoTo vTableEnd      'is the entry point address valid code?
        GetMem1 n, b                                'read the jump opcode at the entry point
        If b <> m Then GoTo vTableEnd               'does the jump opcode match that of the first vtable entry?
    Next v
    Exit Function                                   'last vtable entry not found... increase the For limit?
vTableEnd:
    GetMem4 nAddr - 4, nAddr                        'back one entry to the last private method
    z_scMem = VirtualAlloc(0, 44, &H1000&, &H40&)   'allocate executable memory
    PutMem8 z_scMem + 0, -854782363258311.4703@     'copy the subclass thunk to memory
    PutMem8 z_scMem + 8, 205082594635713.8405@
    PutMem8 z_scMem + 16, 850253272047553.4847@
    PutMem8 z_scMem + 24, -518126163307069.4644@
    PutMem4 z_scMem + 32, nAddr                     'call address
    PutMem8 z_scMem + 36, -802991802926118.8865@
    z_hWnd = UserControl.hWnd
    sc_Subclass = CBool(SetWindowSubclass(z_hWnd, z_scMem, ObjPtr(Me), RefData))
End Function

Private Sub sc_Terminate()
    If z_scMem Then
        RemoveWindowSubclass z_hWnd, z_scMem, ObjPtr(Me)
        VirtualFree z_scMem, 0, &H8000&
    End If
End Sub
'¯¯¯¯¯¯

'Subclass callback. Must be the last private routine in the source_________________________________________________________________________________
Private Function SubCallback(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal RefData As Long) As Long
    If uMsg = WM_MOUSEMOVE Then
        Dim xy As MOUSE_XY
        
        CopyMemory xy, lParam, 4
        UserControl.Cls
        UserControl.Print xy.x & "," & xy.y
    End If
    
    SubCallback = DefSubclassProc(lng_hWnd, uMsg, wParam, lParam)
End Function
'¯¯¯¯¯¯¯¯¯¯¯
