Attribute VB_Name = "HighAccuracy"


Public Function Jia(ByVal Num1 As String, ByVal Num2 As String) As String
'使用方法: 结果 = HighSum(加数1,加数2)
'加数1,2中只能存在整数(不能存在小数点和负号——存在负号就是减法了),如果需要用到小数先化为整数(乘以10的N次方)
'加数1,2必须使用字符串型
'返回值为字符串型

Dim sAnswer As String
Dim lCurIndex As Long
Dim aSingleSum As Integer
Dim bMoreThanTen As Boolean
Dim Length As Long
Dim tNumber1 As Integer, tNumber2 As Integer
Dim Len1 As Long, Len2 As Long, Len3 As Long
Dim bLongerIsNum1 As Boolean
Dim lMostLength As Long
'Initialize
lCurIndex = 0
Num1 = Trim(Num1)
Num2 = Trim(Num2)
Len1 = Len(Num1)
Len2 = Len(Num2)
bLongerIsNum1 = (Len1 > Len2)
'预留足够长的缓冲空间
If bLongerIsNum1 Then
    sAnswer = Space(Len(Num1) + 1)
    lMostLength = Len1 + 1
Else
    sAnswer = Space(Len(Num2) + 1)
    lMostLength = Len2 + 1
End If
Len3 = Len(sAnswer)
'Loop
Do Until (lCurIndex >= lMostLength)
    If Len1 > lCurIndex Then tNumber1 = CInt(Mid$(Num1, Len1 - lCurIndex, 1))
    If Len2 > lCurIndex Then tNumber2 = CInt(Mid$(Num2, Len2 - lCurIndex, 1))
    aSingleSum = CInt(tNumber1 + tNumber2)
    If bMoreThanTen = True Then
        aSingleSum = aSingleSum + 1
        bMoreThanTen = False
    End If
    If aSingleSum >= 10 Then '进位
        bMoreThanTen = True
        aSingleSum = aSingleSum - 10
    End If
    Length = 1
    Mid$(sAnswer, Len3 - lCurIndex, Length) = aSingleSum
    '调整指针变量
    lCurIndex = lCurIndex + 1
    tNumber1 = 0
    tNumber2 = 0
Loop
    Len3 = Len(Trim(sAnswer))
    Dim Len4 As Long
    Dim sLeftNum As String
    If bLongerIsNum1 Then
        Len4 = Len3 - Len1
        If Len4 < 0 Then sLeftNum = Left(Num1, Abs(Len4))
    Else
        Len4 = Len3 - Len2
        If Len4 < 0 Then sLeftNum = Left(Num2, Abs(Len4))
    End If
    Mid$(sAnswer, 1, Abs(Len4)) = sLeftNum
    sAnswer = Trim(sAnswer)
    'Remove Zero
    LeftestNum = Left(sAnswer, 1)
    Do Until LeftestNum <> 0
        sAnswer = Right(sAnswer, Len(sAnswer) - 1)
        LeftestNum = Left(sAnswer, 1)
    Loop
Jia = sAnswer
End Function

Public Function Jian(ByVal A As String, ByVal B As String) As String
Dim Length1, Length2, MaxLength As Integer
Dim Num1, Num2 As String
Dim S1(), S2(), Result() As Integer
Dim Pos As Integer
Dim OK As Boolean



'先比较两个数的大小
Length1 = Len(A)
Length2 = Len(B)

If Length1 > Length2 Then
    Num1 = A: Num2 = Replace(Space(Length1 - Length2), Space(1), "0") & B
    MaxLength = Length1
ElseIf Length2 > Length1 Then
    Num1 = B: Num2 = Replace(Space(Length2 - Length1), Space(1), "0") & A
    MaxLength = Length2
Else
    For i = 1 To Length1
        Tmp1 = Val(Mid(A, i, 1))
        Tmp2 = Val(Mid(B, i, 1))
        If Tmp1 <> Tmp2 Then Exit For
    Next
    If i > Length1 Then Jian = "0": Exit Function
    If Tmp1 > Tmp2 Then
        Num1 = A: Num2 = B
        MaxLength = Length1
    Else
        Num1 = B: Num2 = A
        MaxLength = Length2
    End If
End If


'数组化减数和被减数
ReDim S1(MaxLength), S2(MaxLength), Result(MaxLength) As Integer

For i = MaxLength To 1 Step -1
    S1(i) = Val(Mid(Num1, i, 1))
    S2(i) = Val(Mid(Num2, i, 1))
Next

For i = MaxLength To 1 Step -1
    Result(i) = S1(i) - S2(i)
    If Result(i) < 0 Then
        Result(i) = Result(i) + 10
        Pos = i
        Do
            Pos = Pos - 1
            If S1(Pos) = 0 Then
                S1(Pos) = 9
            Else
                S1(Pos) = S1(Pos) - 1
            End If
        Loop Until S1(Pos) <> 9           '依次借位减1
    End If
Next


'此时结果已经保存在Result数组中，最后还需要去除前面的0

Pos = 1
tmp = ""
OK = False

Do
    If Result(Pos) <> 0 Or OK = True Then tmp = tmp & Trim(Str(Result(Pos))): OK = True
    Pos = Pos + 1
Loop Until Pos > MaxLength


Jian = tmp
End Function








Public Function Cheng(ByVal m As String, ByVal n As String) As String


'高精度乘法
'调用方式
'Dim S1 As String, S2 As String
'S1 = "2385290385102580215818501924820348902395780995725252356236"
'S2 = "1234124923785720589204529017401750734892357947623895893465"
'Print ChengFa(S1, S2)

Dim A() As Integer, B() As Integer, s() As Integer
ReDim A(Len(m)) As Integer, B(Len(n)) As Integer, s(Len(m) + Len(n)) As Integer
For i = 1 To Len(m)
    A(i) = Val(Mid(StrReverse(m), i, 1))
Next
For i = 1 To Len(n)
    B(i) = Val(Mid(StrReverse(n), i, 1))
Next

For i = 1 To Len(m)
    For j = 1 To Len(n)
        s(i + j - 1) = s(i + j - 1) + A(i) * B(j)
    Next
Next
For i = 1 To Len(m) + Len(n)
    If s(i) > 9 Then
        s(i + 1) = s(i + 1) + s(i) \ 10
        s(i) = s(i) Mod 10
    End If
Next
For i = Len(m) + Len(n) To 1 Step -1
    ChengFa = ChengFa & IIf(s(i) = 0, " ", s(i))
Next
Cheng = Replace(LTrim((ChengFa)), " ", 0)
End Function


Public Function Chu(ByVal A As Long, ByVal B As Long, Optional ByVal Accuracy As Integer = 100) As String
Dim arr() As String
Dim s, Pos

Pos = Len(CStr(A \ B))
ReDim arr(1 To Accuracy)

For m = 1 To Accuracy
    arr(m) = A \ B
    A = (A Mod B) * 10
Next

s = Join(arr, "")
Chu = Left(s, Pos) & "." & Right(s, Len(s) - Pos)

End Function




Public Function JieCheng(ByVal num As Long) As String
Dim NumLen As Long, Last As Long, x As Long
Dim i As Long, m As Long, n As Long, nl As Long, s0 As String
Dim Result() As Long, s() As String
NumLen = 1

ReDim Result(1 To NumLen)
nl = 9 - Len(CStr(num))   '根据两数相乘最大得数长度，不会超过两数长度总和的原理，
                    '让数组中每个元素长度与阶乘数长度之和不能超过9，以防止溢出。
If nl < 1 Then nl = 1   '最小长度是1位，若真到了这么大的数，恐怕没人会去试了^-^
n = 10 ^ nl         '缓存用于分隔大数的被除数，数组中每个元素的长度是 nl，该数就是10的 nl 次方
Result(1) = 1
x = 1
Do While x <= num
   Last = 0
   For i = 1 To NumLen
        m = Result(i) * x + Last    '数组中每个元素进行与待乘数相乘后，再加上上次进位数
        Result(i) = m Mod n         '分隔大数
        Last = m / n                '保存进位数并等待累计进下一个数组元素
   Next
   If Last > 0 Then
        m = Len(CStr(Last)) / nl + 1    '对超过数组元素上限的进位数要增加数组大小，并按长度nl分隔
        ReDim Preserve Result(1 To NumLen + m)
        For i = 1 To m
            Result(NumLen + i) = Last Mod n
            Last = Last / n
        Next
        NumLen = UBound(Result)
   End If
   x = x + 1
Loop
ReDim s(1 To NumLen)
s0 = String(nl, "0")    '对长度不足nl的数组元素要在前面补0，不然结果就在错特错了
For i = 1 To NumLen
s(i) = Format(Result(NumLen + 1 - i), s0)   '格式化补 0 每个数组元素
Next
s(1) = Val(s(1))
If s(1) = 0 Then s(1) = ""                  '最高位要去掉0，虽对得数没影响，但位数会错。
JieCheng = Join(s, "")

End Function
