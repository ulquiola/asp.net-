<%
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)
 
Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount
    
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    
    lMessageLength = Len(sMessage)
    
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
    Dim lByte
    Dim lCount
    
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

Function gfnFormatString(str)
	'格式化字符串，主要用于显示有TextArea提交的字符串
		Dim targetStr
		targetStr = replace(str, ">", "&gt;")
		targetStr = replace(str, "<", "&lt;")
		targetStr = Replace(targetStr,vbCrLf,"<BR>")
		targetStr = Replace(targetStr,vbCr,"<BR>")
		targetStr = Replace(targetStr,vbLf,"<BR>")
		'targetStr = Replace(targetStr," ","&nbsp;")
		gfnFormatString = targetStr
		Exit Function
	End Function

Function Crypt(texti,salasana)
       On Error Resume Next
              For T = 1 To Len(salasana)
                     sana = Ascw(Mid(salasana, T, 1))
                     X1 = X1 + sana
              Next
       X1 = Int((X1 * 0.1) / 6)
       salasana = X1
       G = 0
            For TT = 1 To Len(texti)
                     sana = Ascw(Mid(texti, TT, 1))
                     G = G + 1
				If G = 6 Then G = 0
				X1 = 0
				If G = 0 Then X1 = sana - (salasana - 2)
				If G = 1 Then X1 = sana + (salasana - 5)
				If G = 2 Then X1 = sana - (salasana - 4)
				If G = 3 Then X1 = sana + (salasana - 2)
				If G = 4 Then X1 = sana - (salasana - 3)
				If G = 5 Then X1 = sana + (salasana - 5)
				X1 = X1 + G
				Crypted = Crypted & Chrw(X1)
			Next
Crypt = Crypted
End Function


Function DeCrypt(texti,salasana)

       On Error Resume Next

              For T = 1 To Len(salasana)
                     sana = Ascw(Mid(salasana, T, 1))
                     X1 = X1 + sana
              Next

       X1 = Int((X1 * 0.1) / 6)
       salasana = X1
       G = 0

          For TT = 1 To Len(texti)
              sana = Ascw(Mid(texti, TT, 1))
              G = G + 1

               If G = 6 Then G = 0
                X1 = 0
				If G = 0 Then X1 = sana + (salasana - 2)
				If G = 1 Then X1 = sana - (salasana - 5)
				If G = 2 Then X1 = sana + (salasana - 4)
				If G = 3 Then X1 = sana - (salasana - 2)
				If G = 4 Then X1 = sana + (salasana - 3)
				If G = 5 Then X1 = sana - (salasana - 5)
			    X1 = X1 - G
               DeCrypted = DeCrypted & Chrw(X1)
          Next
        DeCrypt = DeCrypted
 End Function
response.expires = 0
response.expiresabsolute = Now() - 1
response.addHeader "pragma","no-cache"
response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"      
%>

