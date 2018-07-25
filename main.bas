Attribute VB_Name = "main"

Option Explicit

Sub main()

MsgBox SHA1HASH("Test. It is the Message.")

End Sub

Public Function SHA1HASH(Message) As String
' originally length < 2^64 bit
' but in this program length < 2^32 byte = 2^35 bit

Dim I As Integer

Dim msg_b() As Byte
ReDim msg_b(0 To Len(Message) - 1) As Byte

Dim Key(1 To 4) As Long
Dim H() As Long
ReDim H(1 To 5) As Long

'Convert the string into bytes as msg_b()
For I = 0 To Len(Message) - 1
msg_b(I) = Asc(Mid(Message, I + 1, 1))
Next I

'set Key
Key(1) = &H5A827999: Key(2) = &H6ED9EBA1: Key(3) = &H8F1BBCDC: Key(4) = &HCA62C1D6

'&H means hexadecimal format (e.g. 0x); set initial H1-H5
H(1) = &H67452301: H(2) = &HEFCDAB89: H(3) = &H98BADCFE: H(4) = &H10325476: H(5) = &HC3D2E1F0

'Call xSHA1 to update H1-H5 recursively
H = xSHA1(msg_b, Key, H)

'Call DecToHex5 to convert H1-H5 from Dec to Hex; Combine H1-H5 to obtain SHA1_Code_temp
SHA1HASH = DecToHex5(H)

'Trim x, and also upper case -> lower case
SHA1HASH = Replace(LCase(SHA1HASH), " ", "")

End Function

Function xSHA1(Msg() As Byte, Key() As Long, H() As Long) As Variant
'It updates the H1-H5 from old to new

Dim U As Long
Dim FB As FourBytes, OL As OneLong
Dim I As Integer
Dim Sec_i As Integer
Dim W(80) As Long
Dim A As Long, B As Long, C As Long, D As Long, E As Long
Dim T As Long


' U (limited to 64 bit) = length of Msg (limited to 2^64 Bytes)
U = UBound(Msg) + 1
OL.L = U32ShiftLeft3(U)

'chop 4*7 = 28 binary zeros, i.e. U32ShiftRight29(U)
A = Int(U / &H20000000)

'LSet means "paste within the entire block in hard drive" in some sense; little endian
'LSet 0x1E, 0x25, 0x3C, 0x41 within FB (totally 4*8 = 32 bits) can become 0x413C251E within OL (totally 64 bits)
' vice versa for "LSet OL = FB"
LSet FB = OL ' break OL into 4 pieces of FB.i

' lengthen Msg from n to n + N, where n + N = 63 (mod 64) by padding zeros in Msg(n+1 to n+N)
ReDim Preserve Msg(0 To (U + 8 And -64) + 63) '-64 = &B11000000
' only &B00111111, &B01111111, &B10111111, &B11111111 are possibles; most likely &B00111111 (e.g. &D63) for short message


Msg(U) = 128 '&D128 = &B10000000
U = UBound(Msg) 'i.e. U = &D63 for short message

'Assign the end of Msg by the message length stored in FB
Msg(U - 4) = A
Msg(U - 3) = FB.D
Msg(U - 2) = FB.C
Msg(U - 1) = FB.B
Msg(U) = FB.A

' Assign W(0) to W(15), totally 16*4 = 64 bytes = 512 bit
For I = 0 To 15
    FB.D = Msg(I * 4)
    FB.C = Msg(I * 4 + 1)
    FB.B = Msg(I * 4 + 2)
    FB.A = Msg(I * 4 + 3)
    LSet OL = FB
    W(I) = OL.L 'i.e. W(i) = Msg(i * 4 + 0) & Msg(i * 4 + 1) & Msg(i * 4 + 2) & Msg(i * 4 + 3)
Next I

' Assign W(16) to W(79), totally 64*4 = 256 bytes = 2048 bit
For I = 16 To 79
    W(I) = U32RotateLeft(1, W(I - 3) Xor W(I - 8) Xor W(I - 14) Xor W(I - 16))
Next I

A = H(1): B = H(2): C = H(3): D = H(4): E = H(5)

For I = 0 To 79
Sec_i = Int((I + 0) / 20) + 1 'Section no. = 1, 2, 3, 4

    T = U32Add(U32Add(U32Add(U32Add(U32RotateLeft(5, A), E), W(I)), Key(Sec_i)), Combine_BCD(Sec_i, B, C, D))
    E = D: D = C: C = U32RotateLeft(30, B): B = A: A = T
Next I

H(1) = U32Add(H(1), A): H(2) = U32Add(H(2), B): H(3) = U32Add(H(3), C): H(4) = U32Add(H(4), D): H(5) = U32Add(H(5), E)

xSHA1 = H

End Function

Function DecToHex5(H() As Long) As String

Dim H_temp As String, L As Long
Dim I As Integer

DecToHex5 = "00000000 00000000 00000000 00000000 00000000"

'DecToHex5 = H1 & H2 & H3 & H4 & H5

For I = 1 To 5

'Convert H(i) from Dec to Hex
H_temp = Hex(H(I))
L = Len(H_temp)
Mid(DecToHex5, 9 * I - L, L) = H_temp

Next I

End Function

