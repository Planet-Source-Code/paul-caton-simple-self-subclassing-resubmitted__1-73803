
;_________________________________________________________________________________________
;
; SubUser.asm - self-subclassing thunk. Assemble with FASM.
;_________________________________________________________________________________________

;bp stack frame after 'add ebp, 28'
dwRefData	equ [ebp + 28]
uIdSubclass	equ [ebp + 24]
lParam		equ [ebp + 20]
wParam		equ [ebp + 16]
uMsg		equ [ebp + 12]
hWnd		equ [ebp +  8]

use32
	xor	eax, eax		;clear eax
	xor	ecx, ecx		;clear ecx
	mov	cl, 7			;number of parameters
	pushad				;preserve all registers
	mov	ebp, esp		;reference the stack
	add	ebp, 28 		;adjust to match stack frame offset definitions and play nice with vb's error system on entry into the callback

	pushd	ebp			;address of the return value, return value popped into eax after popad
	pushd	dwRefData		;RefData
	pushd	lParam			;lParam
	pushd	wParam			;wParam
	pushd	uMsg			;uMsg
	pushd	hWnd			;hWnd

	pushd	uIdSubclass		;object instance
	mov	eax, 0xAAAAAAAA 	;callback address, patched at runtime
	call	eax			;call callback
	popad				;restore registers, pops callback return value into eax
	ret	24			;return
	nop				;padding to 44 bytes
	nop				;padding to 44 bytes