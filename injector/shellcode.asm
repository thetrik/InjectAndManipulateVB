; FASM
; by The trick
; 2018

include "win32wx.inc"

entry start

WM_CREATEOBJECT = WM_USER + 1
WM_DESTROYME = WM_USER + 2

struc GUID def
 {
   match d1-d2-d3-d4-d5, def
    \{
      .Data1 dd 0x\#d1
      .Data2 dw 0x\#d2
      .Data3 dw 0x\#d3
      .Data4 db 0x\#d4 shr 8,0x\#d4 and 0FFh
      .Data5 db 0x\#d5 shr 40,0x\#d5 shr 32 and 0FFh,0x\#d5 shr 24 and 0FFh,0x\#d5 shr 16 and 0FFh,0x\#d5 shr 8 and 0FFh,0x\#d5 and 0FFh
    \}
 }

struct tAPITable
    pfnRegisterClassExW       dd ?
    pfnUnregisterClassW       dd ?
    pfnCreateEventW	      dd ?
    pfnSetWindowsHookExW      dd ?
    pfnPostThreadMessageW     dd ?
    pfnWaitForSingleObject    dd ?
    pfnCloseHandle	      dd ?
    pfnUnhookWindowsHookEx    dd ?
    pfnCallNextHookEx	      dd ?
    pfnCreateWindowExW	      dd ?
    pfnDestroyWindow	      dd ?
    pfnSetEvent 	      dd ?
    pfnPostMessageW	      dd ?
    pfnLoadLibraryW	      dd ?
    pfnFreeLibrary	      dd ?
    pfnGetProcAddress	      dd ?
    pfnSetWindowLongW	      dd ?
    pfnGetWindowLongW	      dd ?
    pfnHeapAlloc	      dd ?
    pfnHeapReAlloc	      dd ?
    pfnHeapFree 	      dd ?
    pfnGetProcessHeap	      dd ?
    pfnGlobalFree	      dd ?
    pfnGlobalSize	      dd ?
    pfnGlobalLock	      dd ?
    pfnGlobalUnlock	      dd ?
    pfnSetTimer 	      dd ?
    pfnKillTimer	      dd ?

    pfnCoInitialize	      dd ?
    pfnCoUninitialize	      dd ?
    pfnCoMarshalInterface     dd ?
    pfnCreateStreamOnHGlobal  dd ?
    pfnGetHGlobalFromStream   dd ?

ends

struct tObjectDesc
    hLibrary		   dd ?
    pObject		   dd ?
    pFactory		   dd ?
    hMem		   dd ?
    dwStmDataSize	   dd ?
    pStmData		   dd ?
ends

struct tProcessData
    tAPIs		   tAPITable
    dwClassAtom 	   dd ?
    dwNumOfWindows	   dd ?
    dwTimerID		   dd ?
ends

struct tThreadData
    dwDestThreadID	   dd ?
    hEvent		   dd ?
    hHook		   dd ?
    hWnd		   dd ?
    pszDllName		   dd ?
    clsid		   dd 4 dup (?)
    iid 		   dd 4 dup (?)
    hr			   dd ?
ends

struct tShellCodeData
    pProcessData	   dd ?
    pThreadData 	   dd ?
ends

interface IClassFactory,\
	   QueryInterface,\
	   AddRef,\
	   Release,\
	   CreateInstance,\
	   LockServer

interface IUnknown,\
	   QueryInterface,\
	   AddRef,\
	   Release

interface IMoniker,\
	   QueryInterface,\
	   AddRef,\
	   Release,\
	   GetClassID,\
	   IsDirty,\
	   Load,\
	   Save,\
	   GetSizeMax,\
	   BindToObject,\
	   BindToStorage,\
	   Reduce,\
	   ComposeWith,\
	   Enum,\
	   IsEqual,\
	   Hash,\
	   IsRunning,\
	   GetTimeOfLastChange,\
	   Inverse,\
	   CommonPrefixWith,\
	   RelativePathTo,\
	   GetDisplayName,\
	   ParseDisplayName,\
	   IsSystemMoniker

section '.idata' import data readable writeable

library kernel32,'KERNEL32.DLL',\
	user32,'USER32.DLL', \
	ole32,'OLE32.DLL'

import user32,\
       RegisterClassExW, 'RegisterClassExW', \
       UnregisterClassW, 'UnregisterClassW', \
       SetWindowsHookExW, 'SetWindowsHookExW', \
       UnhookWindowsHookEx, 'UnhookWindowsHookEx', \
       PostThreadMessageW, 'PostThreadMessageW', \
       GetMessageW, 'GetMessageW', \
       TranslateMessage, 'TranslateMessage', \
       DispatchMessageW, 'DispatchMessageW' , \
       CallNextHookEx, 'CallNextHookEx', \
       CreateWindowExW, 'CreateWindowExW', \
       DestroyWindow, 'DestroyWindow', \
       PostMessageW, 'PostMessageW', \
       SendMessageW, 'SendMessageW', \
       SetWindowLongW, 'SetWindowLongW', \
       GetWindowLongW, 'GetWindowLongW', \
       SetTimer, 'SetTimer', \
       KillTimer, 'KillTimer'

import kernel32,\
       CreateEventW, 'CreateEventW', \
       CloseHandle, 'CloseHandle', \
       CreateThread, 'CreateThread', \
       WaitForSingleObject, 'WaitForSingleObject', \
       SetEvent, 'SetEvent', \
       LoadLibraryW, 'LoadLibraryW', \
       GetProcAddress, 'GetProcAddress', \
       HeapAlloc, 'HeapAlloc', \
       HeapReAlloc, 'HeapReAlloc', \
       HeapFree, 'HeapFree', \
       GetProcessHeap, 'GetProcessHeap', \
       FreeLibrary, 'FreeLibrary', \
       GlobalSize, 'GlobalSize', \
       GlobalLock, 'GlobalLock', \
       GlobalUnlock, 'GlobalUnlock', \
       GlobalFree, 'GlobalFree'


import ole32, \
       CoInitialize, 'CoInitialize', \
       CoUnmarshalInterface, 'CoUnmarshalInterface', \
       CreateStreamOnHGlobal, 'CreateStreamOnHGlobal'

section '.text' code readable executable writable


g_pszDllName du "C:\Temp\Project1.dll", 0
g_clsid GUID 87F10035-26D5-48B9-85C2-B66D1B120A2C

proc start
    locals
      msg MSG
      pObject dd ?
      pStream dd ?
    endl

    mov esi, [g_tShellData.pProcessData]
    mov edi, [g_tShellData.pThreadData]

    mov eax, [RegisterClassExW]
    mov [esi + tProcessData.tAPIs.pfnRegisterClassExW], eax
    mov eax, [UnregisterClassW]
    mov [esi + tProcessData.tAPIs.pfnUnregisterClassW], eax
    mov eax, [CreateEventW]
    mov [esi + tProcessData.tAPIs.pfnCreateEventW], eax
    mov eax, [CloseHandle]
    mov [esi + tProcessData.tAPIs.pfnCloseHandle], eax
    mov eax, [SetWindowsHookExW]
    mov [esi + tProcessData.tAPIs.pfnSetWindowsHookExW], eax
    mov eax, [UnhookWindowsHookEx]
    mov [esi + tProcessData.tAPIs.pfnUnhookWindowsHookEx], eax
    mov eax, [PostThreadMessageW]
    mov [esi + tProcessData.tAPIs.pfnPostThreadMessageW], eax
    mov eax, [WaitForSingleObject]
    mov [esi + tProcessData.tAPIs.pfnWaitForSingleObject], eax
    mov eax, [CallNextHookEx]
    mov [esi + tProcessData.tAPIs.pfnCallNextHookEx], eax
    mov eax, [CreateWindowExW]
    mov [esi + tProcessData.tAPIs.pfnCreateWindowExW], eax
    mov eax, [DestroyWindow]
    mov [esi + tProcessData.tAPIs.pfnDestroyWindow], eax
    mov eax, [SetEvent]
    mov [esi + tProcessData.tAPIs.pfnSetEvent], eax
    mov eax, [PostMessageW]
    mov [esi + tProcessData.tAPIs.pfnPostMessageW], eax
    mov eax, [LoadLibraryW]
    mov [esi + tProcessData.tAPIs.pfnLoadLibraryW], eax
    mov eax, [GetProcAddress]
    mov [esi + tProcessData.tAPIs.pfnGetProcAddress], eax
    mov eax, [GetWindowLongW]
    mov [esi + tProcessData.tAPIs.pfnGetWindowLongW], eax
    mov eax, [SetWindowLongW]
    mov [esi + tProcessData.tAPIs.pfnSetWindowLongW], eax
    mov eax, [HeapAlloc]
    mov [esi + tProcessData.tAPIs.pfnHeapAlloc], eax
    mov eax, [HeapReAlloc]
    mov [esi + tProcessData.tAPIs.pfnHeapReAlloc], eax
    mov eax, [HeapFree]
    mov [esi + tProcessData.tAPIs.pfnHeapFree], eax
    mov eax, [GetProcessHeap]
    mov [esi + tProcessData.tAPIs.pfnGetProcessHeap], eax
    mov eax, [FreeLibrary]
    mov [esi + tProcessData.tAPIs.pfnFreeLibrary], eax
    mov eax, [GlobalSize]
    mov [esi + tProcessData.tAPIs.pfnGlobalSize], eax
    mov eax, [GlobalLock]
    mov [esi + tProcessData.tAPIs.pfnGlobalLock], eax
    mov eax, [GlobalUnlock]
    mov [esi + tProcessData.tAPIs.pfnGlobalUnlock], eax
    mov eax, [GlobalFree]
    mov [esi + tProcessData.tAPIs.pfnGlobalFree], eax
    mov eax, [SetTimer]
    mov [esi + tProcessData.tAPIs.pfnSetTimer], eax
    mov eax, [KillTimer]
    mov [esi + tProcessData.tAPIs.pfnKillTimer], eax

    mov eax, [fs : 0x18]
    mov eax, [eax + 0x24]
    mov [edi + tThreadData.dwDestThreadID], eax
    mov [edi + tThreadData.pszDllName], g_pszDllName

    push edi

    lea edi, [edi + tThreadData.clsid]
    lea esi, [g_clsid]
    mov ecx, 4
    rep movsd

    pop edi
    push edi

    lea edi, [edi + tThreadData.iid]
    lea esi, [IID_IDispatch]
    mov ecx, 4
    rep movsd

    pop edi

    invoke CloseHandle, <invoke CreateThread, 0, 0, EntryPoint, addr g_tShellData, 0, 0>

  .msg_loop:

      .if [edi + tThreadData.hWnd]


	invoke SendMessageW, [edi + tThreadData.hWnd], WM_CREATEOBJECT, 0, 0

	invoke CreateStreamOnHGlobal, [eax + tObjectDesc.hMem], 0, addr pStream
	invoke CoUnmarshalInterface, [pStream], IID_IDispatch, addr pObject

	comcall [pObject], IUnknown, Release

	invoke DestroyWindow, [edi + tThreadData.hWnd]

      .endif

      invoke GetMessageW, addr msg, 0, 0, 0

      .if ~eax
	jmp .exit_proc
      .endif

      invoke TranslateMessage, addr msg
      invoke DispatchMessageW, addr msg

    jmp .msg_loop

  .exit_proc:

    ret

endp

g_tProcessData tProcessData ?
g_tThreadData tThreadData ?
g_tShellData tShellCodeData g_tProcessData, g_tThreadData

proc EntryPoint uses edi esi ebx
    locals
	tClass WNDCLASSEX
    endl

    call .get_ip
  .get_ip:
    pop ebx

    lea edi, [tClass]
    xor eax, eax
    mov ecx, (sizeof.WNDCLASSEX + 4) / 4
    rep stosd

    lea esi, [ebx - (.get_ip - g_tShellData)]
    mov edi, [esi + tShellCodeData.pProcessData]
    mov esi, [esi + tShellCodeData.pThreadData]

    push ebx

    invoke edi + tProcessData.tAPIs.pfnLoadLibraryW, "ole32"
    mov ebx, eax
    call @f
    db "CoInitialize", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCoInitialize], eax
    call @f
    db "CoUninitialize", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCoUninitialize], eax
    call @f
    db "CoMarshalInterface", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCoMarshalInterface], eax
    call @f
    db "CreateStreamOnHGlobal", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnCreateStreamOnHGlobal], eax
    call @f
    db "GetHGlobalFromStream", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, ebx
    mov [edi + tProcessData.tAPIs.pfnGetHGlobalFromStream], eax

    pop ebx

    .if ~[edi + tProcessData.dwClassAtom]

      mov [tClass.cbSize], sizeof.WNDCLASSEX
      lea eax, [ebx + WndProc - .get_ip]
      mov [tClass.lpfnWndProc], eax
      lea eax, [ebx + CLASS_NAME - .get_ip]
      mov [tClass.lpszClassName], eax
      mov eax, [fs:0x30]
      mov eax, [eax + 8] ; Image base
      mov [tClass.hInstance], eax
      mov [tClass.cbWndExtra], 8

      invoke edi + tProcessData.tAPIs.pfnRegisterClassExW, addr tClass

      .if ~ax

	stdcall GetLastErrorHresult
	mov [esi + tThreadData.hr], eax
	jmp .clean_up

      .endif

      movzx eax, ax
      mov [edi + tProcessData.dwClassAtom], eax

    .endif

    invoke edi + tProcessData.tAPIs.pfnCreateEventW, 0, 0, 0, 0

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [esi + tThreadData.hEvent], eax
    lea edx, [ebx + HookProc - .get_ip]

    invoke edi + tProcessData.tAPIs.pfnSetWindowsHookExW, WH_GETMESSAGE, edx, 0, [esi + tThreadData.dwDestThreadID]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [esi + tThreadData.hHook], eax

    invoke edi + tProcessData.tAPIs.pfnPostThreadMessageW, [esi + tThreadData.dwDestThreadID], WM_NULL, 0, esi

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    invoke edi + tProcessData.tAPIs.pfnWaitForSingleObject, [esi + tThreadData.hEvent]

    .if eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

  .clean_up:

    .if [esi + tThreadData.hHook]
      invoke edi + tProcessData.tAPIs.pfnUnhookWindowsHookEx, [esi + tThreadData.hHook]
      mov [esi + tThreadData.hHook], 0
    .endif

    .if [esi + tThreadData.hEvent]
      invoke edi + tProcessData.tAPIs.pfnCloseHandle, [esi + tThreadData.hEvent]
      mov [esi + tThreadData.hEvent], 0
    .endif

    ret
endp

proc HookProc uses esi edi ebx, lCode, wParam, lParam
    locals
      dwRet dd ?
    endl

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    .if [lCode] = HC_ACTION

	invoke edi + tProcessData.tAPIs.pfnCallNextHookEx, 0, [lCode], [wParam], [lParam]
	mov [dwRet], eax

	invoke edi + tProcessData.tAPIs.pfnUnhookWindowsHookEx, [esi + tThreadData.hHook]
	mov [esi + tThreadData.hHook], 0

	invoke edi + tProcessData.tAPIs.pfnCoInitialize, 0

	.if signed eax < 0
	  mov [esi + tThreadData.hr], eax
	  jmp .release_thread
	.endif

	lea ecx, [ebx + CLASS_NAME - g_tShellData]
	mov edx, [fs:0x30]
	mov edx, [edx + 8] ; Image base

	invoke edi + tProcessData.tAPIs.pfnCreateWindowExW, 0, ecx, 0, 0, 0, 0, 0, 0, -3, 0, edx, 0

	mov [esi + tThreadData.hWnd], eax

	.if ~eax

	  stdcall GetLastErrorHresult
	  mov [esi + tThreadData.hr], eax

	  invoke edi + tProcessData.tAPIs.pfnCoUninitialize

	.else
	  inc [edi + tProcessData.dwNumOfWindows]
	.endif

      .release_thread:

	invoke edi + tProcessData.tAPIs.pfnSetEvent, [esi + tThreadData.hEvent]

	mov eax, [dwRet]

    .else
	invoke edi + tProcessData.tAPIs.pfnCallNextHookEx, 0, [lCode], [wParam], [lParam]
    .endif

    ret

endp

proc GetLastErrorHresult

    mov eax, [fs:0x30]
    mov eax, [eax + 0x34]
    and eax, 0xffff
    or eax, 0x80070000

    ret

endp

proc WndProc uses edi esi ebx, hWnd, uMsg, wParam, lParam

    mov eax, [uMsg]

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    .if eax = WM_NCCREATE

      mov eax, 1

    .elseif eax = WM_NCDESTROY

      invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [hWnd], 0

      .if eax

	push esi
	push ebx

	mov ebx, eax

	invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [hWnd], 4
	mov esi, eax

	push esi

	.repeat

	  mov eax, [esi + tObjectDesc.pObject]
	  comcall eax, IUnknown, Release

	  mov eax, [esi + tObjectDesc.pFactory]
	  comcall eax, IClassFactory, LockServer, FALSE

	  mov eax, [esi + tObjectDesc.pFactory]
	  comcall eax, IClassFactory, Release

	  call @f
	  db "DllCanUnloadNow", 0
	  @@:
	  invoke edi + tProcessData.tAPIs.pfnGetProcAddress, [esi + tObjectDesc.hLibrary]

	  stdcall eax

	  .if eax = 0
	    invoke edi + tProcessData.tAPIs.pfnFreeLibrary, [esi + tObjectDesc.hLibrary]
	  .endif

	  add esi, sizeof.tObjectDesc
	  dec ebx

	.until ebx = 0

	pop esi

	invoke edi + tProcessData.tAPIs.pfnHeapFree, <invoke edi + tProcessData.tAPIs.pfnGetProcessHeap>, 0, esi

	pop ebx
	pop esi

      .endif

      mov [esi + tThreadData.hWnd], 0

      invoke edi + tProcessData.tAPIs.pfnCoUninitialize

      dec [edi + tProcessData.dwNumOfWindows]

      .if ZERO?

	lea eax, [ebx + TimerProc - g_tShellData]
	invoke edi + tProcessData.tAPIs.pfnSetTimer, 0, 0, 1, eax
	mov [edi + tProcessData.dwTimerID], eax

      .endif

    .elseif eax = WM_CREATEOBJECT

      stdcall CreateObject

    .elseif eax = WM_DESTROYME

      invoke edi + tProcessData.tAPIs.pfnDestroyWindow, [hWnd]

    .else

      xor eax, eax

    .endif

    ret
endp

proc TimerProc uses esi edi ebx, hWnd, uMsg, idEvent, dwTime

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    mov eax, [fs:0x30]
    mov eax, [eax + 8]
    lea ecx, [ebx + CLASS_NAME - g_tShellData]

    invoke edi + tProcessData.tAPIs.pfnUnregisterClassW, ecx, eax

    mov [edi + tProcessData.dwClassAtom], 0

    invoke edi + tProcessData.tAPIs.pfnKillTimer, 0, [edi + tProcessData.dwTimerID]

    mov [edi + tProcessData.dwTimerID], 0

    ret

endp

proc CreateObject uses esi edi ebx
    locals
      hLib dd 0
      pFactory dd 0
      pObject dd 0
      pStream dd 0
      hMem dd 0
      pStmData dd 0
      dwStmDataSize dd 0
      pList dd ?
      pRet dd 0
    endl

    call .get_ip
  .get_ip:
    pop ebx
    sub ebx, .get_ip - g_tShellData

    mov edi, [ebx + tShellCodeData.pProcessData]
    mov esi, [ebx + tShellCodeData.pThreadData]

    invoke edi + tProcessData.tAPIs.pfnLoadLibraryW, [esi + tThreadData.pszDllName]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [hLib], eax
    call @f
    db "DllGetClassObject", 0
    @@:
    invoke edi + tProcessData.tAPIs.pfnGetProcAddress, eax

    .if ~eax
      jmp .clean_up
    .endif

    lea ecx, [ebx + IID_IClassFactory - g_tShellData]

    stdcall eax, addr esi + tThreadData.clsid, ecx, addr pFactory

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    mov eax, [pFactory]
    comcall eax, IClassFactory, LockServer, TRUE

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    ;lea ecx, [ebx + IID_IDispatch - g_tShellData]

    mov eax, [pFactory]
    comcall eax, IClassFactory, CreateInstance, 0, addr esi + tThreadData.iid, addr pObject

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    invoke edi + tProcessData.tAPIs.pfnCreateStreamOnHGlobal, 0, 0, addr pStream

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    invoke edi + tProcessData.tAPIs.pfnGetHGlobalFromStream, [pStream], addr hMem

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    ;lea ecx, [ebx + IID_IDispatch - g_tShellData]

    invoke edi + tProcessData.tAPIs.pfnCoMarshalInterface, [pStream], addr esi + tThreadData.iid, [pObject], 0, 0, 0

    .if signed eax < 0
      mov [esi + tThreadData.hr], eax
      jmp .clean_up
    .endif

    invoke edi + tProcessData.tAPIs.pfnGlobalSize, [hMem]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [dwStmDataSize], eax

    invoke edi + tProcessData.tAPIs.pfnGlobalLock, [hMem]

    .if ~eax

      stdcall GetLastErrorHresult
      mov [esi + tThreadData.hr], eax
      jmp .clean_up

    .endif

    mov [pStmData], eax

    invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [esi + tThreadData.hWnd], 4

    .if eax

      push ebx
      push esi

      mov ebx, eax

      invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [esi + tThreadData.hWnd], 0
      inc eax
      imul eax, sizeof.tObjectDesc
      mov esi, eax

      invoke edi + tProcessData.tAPIs.pfnHeapReAlloc, <invoke edi + tProcessData.tAPIs.pfnGetProcessHeap>, 0, ebx, esi
      mov [pList], eax

      sub esi, sizeof.tObjectDesc
      mov eax, esi

      pop esi
      pop ebx

    .else
      invoke edi + tProcessData.tAPIs.pfnHeapAlloc, <invoke edi + tProcessData.tAPIs.pfnGetProcessHeap>, 0, sizeof.tObjectDesc
      mov [pList], eax
    .endif

    .if ~eax
      mov [esi + tThreadData.hr], 0x80070007
      jmp .clean_up
    .endif

    mov edx, [hLib]
    mov [eax + tObjectDesc.hLibrary], edx
    mov edx, [pObject]
    mov [eax + tObjectDesc.pObject], edx
    mov edx, [pFactory]
    mov [eax + tObjectDesc.pFactory], edx
    mov edx, [hMem]
    mov [eax + tObjectDesc.hMem], edx
    mov edx, [pStmData]
    mov [eax + tObjectDesc.pStmData], edx
    mov edx, [dwStmDataSize]
    mov [eax + tObjectDesc.dwStmDataSize], edx

    mov [pRet], eax

    invoke edi + tProcessData.tAPIs.pfnSetWindowLongW, [esi + tThreadData.hWnd], 4, [pList]

    invoke edi + tProcessData.tAPIs.pfnGetWindowLongW, [esi + tThreadData.hWnd], 0
    inc eax
    invoke edi + tProcessData.tAPIs.pfnSetWindowLongW, [esi + tThreadData.hWnd], 0, eax

  .clean_up:

    .if [pStream]
      comcall [pStream], IUnknown, Release
    .endif

    .if ~[pRet]

      .if [hMem]
	invoke edi + tProcessData.tAPIs.pfnGlobalFree, [hMem]
      .endif

      .if [pObject]
	comcall [pObject], IUnknown, Release
      .endif

      .if [pFactory]
	comcall [pFactory], IClassFactory, LockServer, FALSE
	comcall [pFactory], IClassFactory, Release
      .endif

      .if [hLib]
	invoke edi + tProcessData.tAPIs.pfnFreeLibrary, [hLib]
      .endif

    .endif

    mov eax, [pRet]

    ret

endp

IID_IClassFactory      GUID 00000001-0000-0000-C000-000000000046
IID_IDispatch	       GUID 00020400-0000-0000-C000-000000000046
CLASS_NAME: du "VBInjectClass", 0