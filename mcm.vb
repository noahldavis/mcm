
Sub Moodys_Method()
    Dim Num_bonds As Integer
    Dim kount As Integer
    Dim j As Integer
    Dim PD, rho, PD_conditional(), tmp As Double
    Dim Log_prod_PD_cond(), jsum As Double
    Dim Prob_vec(), Prob_zero_rho() As Double
    Dim Prob_array(), tmpj As Double
    
    '   Clear output
    Range("Screen_out").Resize(1000, 3).ClearContents
    
    '   Read input
    Num_bonds = Range("Num_bonds").Value
    PD = Range("PD").Value
    rho = Range("rho").Value
    
    ReDim Prob_vec(Num_bonds)
    ReDim Prob_zero_rho(Num_bonds)
    ReDim PD_conditional(Num_bonds)
    ReDim Log_prod_PD_cond(Num_bonds)
    ReDim Prob_array(Num_bonds, Num_bonds)
    
    '   As it happens, the current "My method" doesn't work.
    '   My view is that the probabilities are path-dependent
    '   as I found analytically.  I should be able to write a
    '   path-dependent procedure.  It would have interesting
    '   attributes.  First, it means that the Moody's assumption
    '   that all paths have equal weight is wrong.  Second, it
    '   then defeats the idea of correlation that doesn't care
    '   about the order of consideration of the bonds.  Third,
    '   the algorithm will never get too far since the number of
    '   paths is 2^Num_bonds.  That means I can demonstrate for, say,
    '   10 bonds that it works and makes sense (and differs from
    '   Moody's), but I'd never have an algorithm for a practical
    '   number of bonds.
    
    '   Try the Moody's "better" algorithm here.
    PD_conditional(1) = PD
    Prob_array(1, 1) = PD_conditional(1)
    Prob_array(0, 1) = 1# - Prob_array(1, 1)
    For kount = 2 To Num_bonds
        PD_conditional(kount) = rho + (1# - rho) * PD_conditional(kount - 1)
        Prob_array(kount, kount) = PD_conditional(kount) * Prob_array(kount - 1, kount - 1)
        '   Moody's method
        For j = kount - 1 To 0 Step -1
            Prob_array(j, kount) = Prob_array(j, kount - 1) - Prob_array(j + 1, kount)
        Next j
'        '   My method
'        For j = kount - 1 To 0 Step -1
'            Prob_array(j, kount) = Prob_array(j, kount - 1) * _
'                (1# - PD_conditional(j + 1) * ((1# - rho) ^ (kount - j - 1)))
'        Next j
    Next kount
    
    '   Also try what I call the Moody's "literal" algorithm.  This is the
    '   brute force inner summation.  Use PD_conditional() we computed above
    '   and now also form the cumulative sum of logarithms for the literal algorithm.
    Log_prod_PD_cond(0) = 0#
    For j = 1 To Num_bonds
        Log_prod_PD_cond(j) = Log_prod_PD_cond(j - 1) + Log(PD_conditional(j))
    Next j
    
    '   Calculate and write the probabilities of getting
    '   k defaults of the total number of bonds.
    For kount = 0 To Num_bonds
        tmp = JoeLogCombin(Num_bonds, kount)
        Prob_zero_rho(kount) = Exp(tmp + kount * Log(PD) + (Num_bonds - kount) * Log(1# - PD))
        
        '   Moody's "better" algorithm
        Prob_vec(kount) = Prob_array(kount, Num_bonds) * Exp(tmp)
        
        '   Moody's "literal" algorithm
        jsum = 0#
        For j = 0 To Num_bonds - kount Step 2
            tmpj = JoeLogCombin(Num_bonds - kount, j)
            jsum = jsum + Exp(tmpj + Log_prod_PD_cond(j + kount))
        Next j
        For j = 1 To Num_bonds - kount Step 2
            tmpj = JoeLogCombin(Num_bonds - kount, j)
            jsum = jsum - Exp(tmpj + Log_prod_PD_cond(j + kount))
        Next j
        '   Prob_vec(kount) = jsum * Exp(tmp)
        
        Range("Screen_out").Offset(kount, 0).Value = kount
        Range("Screen_out").Offset(kount, 1).Value = Prob_vec(kount)
        Range("Screen_out").Offset(kount, 2).Value = Prob_zero_rho(kount)
    Next kount

End Sub

Sub Shift_PD_Method()
    Dim Num_bonds As Integer
    Dim kount As Long
    Dim k As Integer
    Dim j, b(), Begin_PD_calc As Integer
    Dim Num_paths As Long
'   Dim Path_string As String
    Dim PD, rho, wm_rho, log_wm_rho, Shift_PD(), tmp, tmp_time As Double
    Dim Prob_vec(), Prob_zero_rho(), Shift_temp() As Double
    Dim wm_rho_power(), Sum_temp() As Double
    
    '   Clear output
    Range("Screen_out").Resize(1000, 3).ClearContents
    
    '   Read input
    Num_bonds = Range("Num_bonds").Value
    PD = Range("PD").Value
    rho = Range("rho").Value
    
    ReDim Prob_vec(Num_bonds)
    ReDim Prob_zero_rho(Num_bonds)
    ReDim Shift_PD(Num_bonds)
    ReDim Shift_temp(Num_bonds)
    ReDim wm_rho_power(Num_bonds)
    ReDim Sum_temp(Num_bonds, Num_bonds)
    ReDim b(Num_bonds)
    
    tmp_time = Timer
    
    '   Useful pre-calculations
    wm_rho = 1# - rho
    log_wm_rho = Log(wm_rho)
    For j = 1 To Num_bonds
        wm_rho_power(j) = Exp((j - 1) * log_wm_rho)
    Next j
    
    '   Initialize Prob_vec, b(j), Num_defaults, Begin_PD_calc, and Sum_temp.
    For j = 0 To Num_bonds
        Prob_vec(j) = 0#
        b(j) = 0
        For k = 0 To Num_bonds
            Sum_temp(j, k) = 0#
        Next k
    Next j
    Num_defaults = 0
    Begin_PD_calc = 2
    
    '   Set Shift_temp().  This is a portion of Shift_PD()
    '   that we need not calculate within the Num_paths loop.
    Shift_temp(1) = PD
    For k = 2 To Num_bonds
        Shift_temp(k) = PD * wm_rho_power(k)
    Next k
    
    '   This is my path-dependent procedure.  It should have interesting
    '   attributes.  First, it means that the Moody's assumption
    '   that all paths have equal weight is wrong.  Second, it
    '   then defeats the idea of correlation that doesn't care
    '   about the order of consideration of the bonds.
    
    Num_paths = Application.Power(2, Num_bonds)
    Shift_PD(1) = PD
    For kount = 0 To Num_paths - 1
        '   The lines we comment out below represent our first
        '   method of determing the b(j) vector.  Each component
        '   of b(j) is zero or one depending on whether the bond j
        '   is in default.  Naturally, this b(j) has Num_bonds
        '   components.  We find a faster algorithm (that is less
        '   intuitive to read) and it sits at the bottom of
        '   this for-loop.
'        '   Get the path string, its associated string array,
'        '   and the number of defaults of the path.
'        Num_defaults = 0
'        Path_string = Base_Two_String(kount, Num_bonds)
'        For j = 1 To Num_bonds
'            b(j) = CInt(Mid$(Path_string, j, 1))
'            '   b(j) = 1
'            Num_defaults = Num_defaults + b(j)
'        Next j
        '   Get the Shifted PD at each bond.
        For k = Begin_PD_calc To Num_bonds
            '   tmp = 0#
            For j = Begin_PD_calc - 1 To k - 1
            '   For j = 1 To k - 1
                '   tmp = tmp + b(k - j) * wm_rho_power(j)
                Sum_temp(j, k) = Sum_temp(j - 1, k) + b(j) * wm_rho_power(k - j)
            Next j
            '   Shift_PD(k) = Shift_temp(k) + rho * tmp
            Shift_PD(k) = Shift_temp(k) + rho * Sum_temp(k - 1, k)
        Next k
        
        '   Get the probability of this path.
        tmp = 0#
        For k = 1 To Num_bonds
            tmp = tmp + b(k) * Log(Shift_PD(k)) + (1 - b(k)) * Log(1# - Shift_PD(k))
        Next k
        '   Get the probability for each possible number of defaults
        '   by summing over all paths.
        Prob_vec(Num_defaults) = Prob_vec(Num_defaults) + Exp(tmp)
        
        '   There's a faster way to generate the b(j) than our first method
        '   of converting kount to Base 2.  We do this at the end of the loop
        '   so that the value we set here applies for the next kount.  We
        '   adjust Num_defaults with the counting.  We also set Begin_PD_calc
        '   to tell the algorithm at what point the b(j) change from one value
        '   of kount to the next.  Having this Begin_PD_calc prevents
        '   unnecessarily re-calculating Shift_PD(k) values for k < Begin_PD_calc.
        For j = Num_bonds To 0 Step -1
            If b(j) = 0 Then
                b(j) = 1
                Num_defaults = Num_defaults + 1
                Begin_PD_calc = j + 1
                Exit For
            Else
                b(j) = 0
                Num_defaults = Num_defaults - 1
            End If
        Next j
        
    Next kount
    
    '   Calculate and write the probabilities of getting
    '   k defaults of the total number of bonds.
    For k = 0 To Num_bonds
        tmp = JoeLogCombin(Num_bonds, k)
        Prob_zero_rho(k) = Exp(tmp + k * Log(PD) + (Num_bonds - k) * Log(1# - PD))
        Range("Screen_out").Offset(k, 0).Value = k
        Range("Screen_out").Offset(k, 1).Value = Prob_vec(k)
        Range("Screen_out").Offset(k, 2).Value = Prob_zero_rho(k)
    Next k
    Range("Screen_out").Offset(0, -1).Value = Timer - tmp_time

End Sub

'   For an input number of bonds, we will chart all possible binary paths (indicated default / no default).
'   The number of paths is 2 raised to power of the number of bonds.  The path numbers range from zero to
'   2^Num_bonds-1.  This function produces a string of 0/1 values giving the path number in Base 2.
'   The length of the output string is Num_bonds.
Function Base_Two_String(ByVal Long_int As Long, ByVal Num_bonds As Integer) As String
    Dim kount, Add_zeroes As Integer
    Dim tmp_string As String
    
    tmp_string = ""
    If Long_int < 0 Or Long_int >= Application.Power(2, Num_bonds) Then
        Base_Two_String = tmp_string
        Exit Function
    End If
    
    Do While Long_int > 0
        If 2 * CLng(Long_int / 2) = Long_int Then
            tmp_string = "0" + tmp_string
        Else
            tmp_string = "1" + tmp_string
            Long_int = Long_int - 1
        End If
        Long_int = Long_int / 2
    Loop
    
    '   Add zeroes at beginning to make the string length equal to Num_bonds.
    Add_zeroes = Num_bonds - Len(tmp_string)
    For kount = 1 To Add_zeroes
        tmp_string = "0" + tmp_string
    Next kount
    
    Base_Two_String = tmp_string
    
End Function

Sub Mixture_Method()
    Dim Num_bonds As Integer
    Dim kount As Integer
'   Dim j As Integer
    Dim PD, rho, tmp, weight, p1, p2 As Double
    Dim p1_minus_p2 As Double
    Dim Prob_vec(), Prob_zero_rho() As Double
    
    '   Clear output
    Range("Screen_out_mix").Resize(1000, 3).ClearContents
    
    '   Read input
    Num_bonds = Range("Num_bonds_mix").Value
    PD = Range("PD_mix").Value
    rho = Range("rho_mix").Value
    weight = Range("Weight").Value
    
    ReDim Prob_vec(Num_bonds)
    ReDim Prob_zero_rho(Num_bonds)
    
    '   Determine p1 and p2.
    p1_minus_p2 = Sqr(rho * PD * (1# - PD) / (weight * (1# - weight)))
    p2 = PD - weight * p1_minus_p2
    p1 = p2 + p1_minus_p2
    
    '   Calculate and write the probabilities of getting
    '   k defaults of the total number of bonds.
    For kount = 0 To Num_bonds
        tmp = JoeLogCombin(Num_bonds, kount)
        Prob_zero_rho(kount) = Exp(tmp + kount * Log(PD) + (Num_bonds - kount) * Log(1# - PD))
        Prob_vec(kount) = Exp(tmp + Log(weight) + kount * Log(p1) + (Num_bonds - kount) * Log(1# - p1)) + _
            Exp(tmp + Log(1# - weight) + kount * Log(p2) + (Num_bonds - kount) * Log(1# - p2))
        Range("Screen_out_mix").Offset(kount, 0).Value = kount
        Range("Screen_out_mix").Offset(kount, 1).Value = Prob_vec(kount)
        Range("Screen_out_mix").Offset(kount, 2).Value = Prob_zero_rho(kount)
    Next kount

End Sub

'   Return the log of the combinatorial "N choose k".
Function JoeLogCombin(N As Integer, k As Integer) As Double
    Dim N_dbl, k_dbl As Double
    
    N_dbl = CDbl(N)
    k_dbl = CDbl(k)
    
    JoeLogCombin = 0#
    
    '   Faulty input guard
If k < 0 OrElse k > N Then
        Exit Function
    End If
    
If k = 1 OrElse k = N - 1 Then
        JoeLogCombin = Log(N_dbl)
    ElseIf k > 1 AndAlso k < N - 1 Then
        JoeLogCombin = GammLn(N_dbl + 1#) - GammLn(k_dbl + 1#) - GammLn(N_dbl - k_dbl + 1#)
    End If
    
End Function
