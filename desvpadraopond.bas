Attribute VB_Name = "Módulo1"
Function mediap(ParamArray v() As Variant) As Double
    
    Dim i As Integer
    Dim soman As Double
    Dim nums As Range, pesos As Range
    
    Set nums = v(0)
    Set pesos = v(1)
    
    mediap = -1
    
    If UBound(v()) <> 1 Or UBound(nums()) <> UBound(pesos()) Then Exit Function
    
        For i = 0 To UBound(nums())
                If IsNumeric(nums(i)) And IsNumeric(pesos(i)) Then
                    soman = soman + pesos(i).Value * nums(i).Value
                    somap = somap + pesos(i).Value
                End If
        Next
            
    mediap = soman / somap

End Function

Function desvpp(ParamArray v() As Variant)

Dim medp As Double

medp = mediap(v(0), v(1))

    Dim i As Integer
    Dim soman As Double
    Dim nums As Range, pesos As Range
    
    Set nums = v(0)
    Set pesos = v(1)
    
    desvpp = -1
    
    If UBound(v()) <> 1 Or UBound(nums()) <> UBound(pesos()) Then Exit Function
    
        For i = 0 To UBound(nums())
                If IsNumeric(nums(i)) And IsNumeric(pesos(i)) Then
                    soman = soman + pesos(i).Value * (medp - nums(i)) ^ 2
                    somap = somap + pesos(i).Value
                End If
        Next
        
        desvpp = Math.Sqr(soman / (somap - 1))

End Function

