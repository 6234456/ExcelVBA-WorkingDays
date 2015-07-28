Function tageImMonat(ByVal m As Integer, j As Integer) As Integer
    
    Dim res As Integer
    Dim i
    
    
    For i = 28 To 31
        If Month(DateSerial(j, m, i + 1)) = m + 1 Then
            res = i
            Exit For
        End If
    Next i
    
    tageImMonat = res

End Function

Private Function imMonat(ByVal m As Integer, ByVal j As Integer, Optional land As String) As Integer
    imMonat = vonbisLand(DateSerial(j, m, 1), DateSerial(j, m, tageImMonat(m, j)), land)
End Function

Private Function bisLetztemSonntag(ByVal t As Date, Optional ByVal land As String = "b") As Integer

    lastSunday = letzterSonntag(t)
    
    bisLetztemSonntag = vonbisLand(DateSerial(Year(lastSunday), Month(lastSunday), 1), lastSunday, land)

End Function
Function vonbis(ByVal s As Date, ByVal e As Date, mixin As Variant) As Integer
   
    Dim res As Integer
    res = 0
    
    Dim tmp As Date
    tmp = s
    
    Do While tmp <= e
        
        If Not isWochenende(tmp) Then
            If Not isFeiertag(tmp, mixin) Then
                res = res + 1
            End If
        End If
        
        tmp = DateAdd("d", 1, tmp)
    Loop
    
    vonbis = res

End Function

Function letzterSonntag(ByVal t As Date) As Date
    letzterSonntag = DateAdd("d", -1 * Weekday(t, vbMonday), t)
End Function

Function letzterMontag(ByVal t As Date) As Date
    letzterMontag = DateAdd("d", -6, letzterSonntag(t))
End Function

Function vonbisLand(ByVal s As Date, ByVal e As Date, Optional ByVal land As String = "b") As Integer
    Dim res As Integer
    res = 0

    If Year(s) = Year(e) Then
        res = vonbis(s, e, feiertagMixin(Year(s), land))
    Else
        res = vonbis(s, e, mixins(rng(Year(s), Year(e)), land))
    End If
    
    vonbisLand = res

End Function


Function isFeiertag(ByVal targDate As Date, mixin As Variant) As Boolean
    
    Dim i
    Dim res As Boolean
    
    
    For i = LBound(mixin) To UBound(mixin)
        If mixin(i) = targDate Then
            res = True
            Exit For
        End If
    Next i
    
    isFeiertag = res
    
End Function

Function isWochenende(ByVal targDate As Date) As Boolean
    
    isWochenende = (Weekday(targDate, vbMonday) > 5)

End Function

Function arrConcate(ByVal arr1, ByVal arr2) As Variant

    Dim l1, h1, l2, h2, i As Integer
    
    l1 = LBound(arr1)
    l2 = LBound(arr2)
    
    h1 = UBound(arr1)
    h2 = UBound(arr2)
    
    Dim arr()
    ReDim arr(l1 To h1 + h2 - l2 + 1)
    
    For i = l1 To h1 + h2 - l2 + 1
        If i <= h1 Then
            arr(i) = arr1(i)
        Else
            arr(i) = arr2(l2 + i - h1 - 1)
        End If
    Next i
    
    arrConcate = arr

End Function


Function mixins(yearArr, Optional ByVal land As String = "b") As Variant

    Dim arr() As Variant
    
    Dim h, l, i As Integer
    
    h = UBound(yearArr)
    l = LBound(yearArr)
    
    arr = feiertagMixin(yearArr(l), land)
    
    For i = l + 1 To h
        arr = arrConcate(arr, feiertagMixin(yearArr(i), land))
    Next i
    
    mixins = arr

End Function


Function feiertagMixin(ByVal j As Integer, Optional ByVal land As String = "b") As Variant

    Dim arr()
    
     If land = "b" Or land = "B" Then
        arr = Array( _
                    DateSerial(j, 1, 1), _
                    DateSerial(j, 1, 6), _
                    DateSerial(j, 5, 1), _
                    DateSerial(j, 8, 15), _
                    DateSerial(j, 10, 3), _
                    DateSerial(j, 11, 1), _
                    DateSerial(j, 12, 25), _
                    DateSerial(j, 12, 26) _
                    )
    Else
        arr = Array(DateSerial(j, 1, 1), DateSerial(j, 5, 1), DateSerial(j, 10, 3), DateSerial(j, 10, 31), DateSerial(j, 12, 25), DateSerial(j, 12, 26))
    End If
    
    ost = osterbezogen(j)
    
    l = LBound(arr)
    h = UBound(arr)
    le = UBound(ost) - LBound(ost) + 1
    
    ReDim Preserve arr(l To h + le)
    
    For i = h + 1 To h + le
        arr(i) = ost(i - h - 1)
    Next i
    
    feiertagMixin = arr

End Function

' from http://en.wikipedia.org/wiki/Computus
Private Function easter(X)                                  ' X = year to compute
    Dim K, m, s, A, D, R, OG, SZ, OE
 
    K = X \ 100                                     ' Secular number
    m = 15 + (3 * K + 3) \ 4 - (8 * K + 13) \ 25    ' Secular Moon shift
    s = 2 - (3 * K + 3) \ 4                         ' Secular sun shift
    A = X Mod 19                                    ' Moon parameter
    D = (19 * A + m) Mod 30                         ' Seed for 1st full Moon in spring
    R = D \ 29 + (D \ 28 - D \ 29) * (A \ 11)       ' Calendarian correction quantity
    OG = 21 + D - R                                 ' Easter limit
    SZ = 7 - (X + X \ 4 + s) Mod 7                  ' 1st sunday in March
    OE = 7 - (OG - SZ) Mod 7                        ' Distance Easter sunday from Easter limit in days
 
    easter = DateSerial(X, 3, OG + OE)              ' Result: Easter sunday as number of days in March
End Function


Function osterbezogen(X)
    Dim Karfreitag, Ostersonntag, Ostermontag, ChrHim, Pfingstmontag, Fronleichnam As Date
    
     Ostersonntag = easter(X)
    
    'Karfreitag Freitag vor Ostersonntag
     Karfreitag = DateAdd("d", -2, Ostersonntag)
    
    'Montag nach Ostersonntag
     Ostermontag = DateAdd("d", 1, Ostersonntag)
    
    'Christi Himmelfahrt 39. Tag nach Ostersonntag
     ChrHim = DateAdd("d", 39, Ostersonntag)
    
    'Pfingstmontag 50. Tag nach Ostersonntag
     Pfingstmontag = DateAdd("d", 50, Ostersonntag)
    
    'Fronleichnam 60. Tag nach Ostersonntag
     Fronleichnam = DateAdd("d", 60, Ostersonntag)
    
     osterbezogen = Array(Karfreitag, Ostersonntag, Ostermontag, ChrHim, Pfingstmontag, Fronleichnam)

End Function

Private Function rng(ByVal s As Integer, ByVal e As Integer, Optional ByVal interval As Integer = 1)
    Dim arr()
    
    If e <= s Then
        Err.Raise (8888)
    Else
        ReDim arr(0 To e - s)
        Dim i As Integer
        
        For i = 0 To e - s Step interval
            arr(i) = s + i
        Next i
    End If
    
    rng = arr
    
End Function


