;FastXor.asm for NASM

%Define str [ebp+8]
%Define len [ebp+12]
%Define pwstr [ebp+16]
%Define pwlen [ebp+20]

[BITS 32]

push ebp ; preserve registers
mov ebp,esp
push ebx
push esi
push edi
;-----------
mov eax,len ; test strings for zero length
test eax,eax
jz exit
mov eax,pwlen
test eax,eax
jz exit
mov ecx,len ; ecx=len
shl ecx,1 ; ecx=2*ecx because strings are saved as unicode in VB,
          ; (every character is followed by hex 00)
mov edx,str ; => edx points to the first character
add edx,ecx ; => edx points to the last character,
neg ecx ; => edx+ecx points to the first character
mov ebx,pwlen ; ebx=pwlen
shl ebx,1 ; ebx=2*ebx  because strings are saved as unicode in VB
mov eax,pwstr ; => eax points to the first character
add eax,ebx ; => eax points to the last character,
mov pwstr,eax ; we have to save pwstr to memory because eax is used in the routine
neg ebx ; => eax+ebx points to the first character
mov pwlen,ebx ; save pwlen because it is used again in the loop
LP:
mov eax,[edx+ecx] ; eax points to the next character
add ebx,pwstr ; ebx points to the next password character
xor al,[ebx] ; xor character
sub ebx,pwstr ; ebx = offset of current password char
add ebx,2 ; let ebx point to the next password char
jnz PWNZ ; if we looped through the password,
mov ebx,pwlen ; jump to the first password character again
PWNZ:
mov [edx+ecx],eax ; store encrypted character in string
add ecx,2 ; let ecx point to the next character
jnz LP ; loop if there are still characters to encrypt
exit:
xor eax,eax ; eax=0, CallWindowProc returns this value
;-----------
pop edi ; Recover register values
pop esi
pop ebx
mov esp,ebp
pop ebp
ret 16
