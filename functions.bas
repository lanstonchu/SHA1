Attribute VB_Name = "functions"
' All these funtions use their input only on ByVal basis
' long-A (limited to 64 bit) is the length of message (limited to 2^64 Byte)
' but according to the code, A should be limited to 4*8=32 bits
Dim i_t As Long

Option Explicit

Type FourBytes
'4*8 = 32 bits in total
    A As Byte
    B As Byte
    C As Byte
    D As Byte
End Type

Type OneLong
'64 bits
    L As Long
End Type


Function U32ShiftLeft3(ByVal A As Long) As Long
' e.g. 1111edf...z -> 1edf...z000
' e.g. 1011edf...z -> 0edf...z000

Dim n As Integer
n = 3

'abcdefg...wxyz -> 0000efg...wxyz -> 0efg...wxyz000
U32ShiftLeft3 = U32Chop(1, n, A)

'if first 4 bits of original A is &B1111, then keep the first bit of U32ShiftLeft3 as binary 1
'e.g. if abcd=1111 then (0efg...wxyz000 -> 1efg...wxyz000)
 If A And (2 ^ (32 - n - 1)) Then U32ShiftLeft3 = U32ShiftLeft3 Or (-(2 ^ 31))
     
End Function

Function U32Add(ByVal A As Long, ByVal B As Long) As Long

'When A and B are one positive and negative, they can be simply added up
If (A Xor B) < 0 Then
U32Add = A + B
Else

'When A and B have the same signs, add them in binary form, and then drop the extra first digits to keep 32 bits
U32Add = (A Xor &H80000000) + B Xor &H80000000
End If

End Function

Function U32RotateLeft(n As Integer, ByVal A As Long) As Long
'n = 1, ..., 31
'e.g. abcdefg...yz -> fg...yzabcde (for n = 5)

'e.g. abcdefg...yz -> g...yzabcde (for n = 5)
U32RotateLeft = U32Chop(0, n, A) Or U32Chop(1, n, A)
 
'e.g. gh...yzabcde -> fgh...yzabcde
If A And (2 ^ (32 - n - 1)) Then U32RotateLeft = U32RotateLeft - (2 ^ 31) 'again, use negative number to avoid overflow

End Function

Function U32Chop(typ As Integer, n As Integer, ByVal A As Long) As Long
'typ = 0 (chop for head) or 1 (chop for tail)

Dim Head_One As Long
Dim Tail_One As Long
Dim Chop4Head As Long
Dim Chop4Tail As Long

Select Case typ

    Case 0 'chop for head (i.e. chop tail)
    
    'use negative value to store to avoid overflow; x = x - 2^32 to store, if x > (2^31)-1
    Head_One = (2 ^ n - 1) * (2 ^ (32 - n)) - (2 ^ 32) 'e.g. 5*one at head when n = 5
    
    'e.g. (N_abcdefghi....yz -> N_abcde -> abcde)
    Chop4Head = Int((A And Head_One) / (2 ^ (32 - n))) And (2 ^ n - 1)
    
    U32Chop = Chop4Head
    
    Case 1 'chop for tail  (i.e. chop head)
    
    Tail_One = 2 ^ (32 - n - 1) - 1 'e.g. 26*one at tail when n = 5 (not 27 to avoid overflow)
    
    'e.g. (abcdefgh...yz -> 0gh...yz00000)
    Chop4Tail = (A And Tail_One) * (2 ^ n)

    U32Chop = Chop4Tail

End Select

End Function


Function Combine_BCD(Sec_i As Integer, B As Long, C As Long, D As Long)

Select Case Sec_i

    Case 1
    Combine_BCD = ((B And C) Or ((Not B) And D))
    
    Case 2
    Combine_BCD = (B Xor C Xor D)
    
    Case 3
    Combine_BCD = ((B And C) Or (B And D) Or (C And D))
    
    Case 4
    Combine_BCD = (B Xor C Xor D)

End Select

End Function

