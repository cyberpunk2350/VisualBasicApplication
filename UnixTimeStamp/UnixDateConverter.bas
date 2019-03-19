'http://www.vbforums.com/showthread.php?513727-RESOLVED-Convert-Unix-Time-to-Date

Private Const Unix1970 As Long = 25569 'CDbl(DateSerial(1970, 1, 1))

Public Function Date2Unix(ByVal vDate As Date) As Long
Date2Unix = DateDiff("s", Unix1970, vDate)
End Function

Public Function Unix2Date(vUnixDate As Long) As Date
Unix2Date = DateAdd("s", vUnixDate, Unix1970)
End Function

