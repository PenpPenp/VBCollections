Module Module1
    Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" _
        (ByRef lpFrequency As Long) As Boolean
    Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" _
        (ByRef lpPerformanceCount As Long) As Boolean

    Sub Main()
        Dim frequency As Decimal
        Dim startTime As Decimal
        Dim endTime As Decimal
        Dim res As Double

        Dim c As Long = 0

        ' get the frequency counter
        ' return zero if hardware doesn't support high-res performance counters
        If QueryPerformanceFrequency(frequency) = 0 Then
            'MsgBox "This computer doesn't support high-res timers", vbCritical
            Console.WriteLine("his computer doesn't support high-res timers")
            Exit Sub
        End If

        ' start timing
        QueryPerformanceCounter(startTime)

        For i As Integer = 0 To 9999999
            c += i
        Next i

        ' end timing  
        QueryPerformanceCounter(endTime)

        ' note that both dividend and divisor are scaled
        ' by 10,000, so you don't need to scale the result
        res = ((endTime - startTime) / frequency) * 1000
        Console.WriteLine("Elapsed Time (" + Format(res, "0.00000") + " ms)")

        Console.ReadLine()
    End Sub

End Module

'Output:
'Elapsed Time (18.78680 ms)