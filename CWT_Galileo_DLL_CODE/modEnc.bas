Attribute VB_Name = "modEnc"
Public Const gintKey = 5280

Public Function encrypt(txt As String, pw As Integer)
    Dim key, enc As String
    Dim i, f As Integer
    
    
    f = Mid(pw, 1, 1)                    'first number of the pw decides the hash
    For i = 1 To (Int(Len(txt) / 4) + 1) 'length of the text divided by 4 because 4 is the length of the key
        key = key & (pw)    ' creating the key by adding pw to key each time
        pw = pw + f        ' equation for messing with pw so it changes
        If Len(key) >= Len(txt) Then Exit For  'if the key is as long as the txt, stop
    Next i
    For i = 1 To Len(txt)
        enc = enc & Chr(Asc(Mid(txt, i, 1)) + Mid(key, i, 1))  'adds the first key # to the ascii value of the first char in the txt and so on...
    Next i
        encrypt = enc
        
End Function

Public Function decrypt(txt As String, pw As Integer)
    Dim key, enc As String
    Dim i, f As Integer
    f = Mid(pw, 1, 1)
    For i = 1 To (Int(Len(txt) / 4) + 1)
        key = key & (pw)
        pw = pw + f
        If Len(key) >= Len(txt) Then Exit For
    Next i
    
     For i = 1 To Len(txt)
        enc = enc & Chr(Asc(Mid(txt, i, 1)) - Mid(key, i, 1))
    Next i
    
    decrypt = enc
End Function

