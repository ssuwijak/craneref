Dim radius, maxWeight, opt
Dim msg1, msg2

opt = InputBox("What counterweight configuration?" & vbCrLf & vbCrLf & "(1) 10,800# CW" & vbCrLf & "(2) 18,400# CW" & vbCrLf & vbCrLf & "type 1 or 2")

If (opt<>1 AND opt<>2) Then
	msg1 = "Sorry, you entered wrong value"
	Wscript.quit()
End If

maxWeight = InputBox("What is the max weight?", "Crane Load Chart")


If IsNumeric(maxWeight) Then
    maxWeight = CInt(maxWeight)
    radius = InputBox("What is your radius for the 40 Ton?", "Crane Load Chart")

    If IsNumeric(radius) Then
        radius = CInt(radius)
        Select Case radius
            Case 10
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 137100, maxWeight) & _
                DisplayBoom("50'", 89000, maxWeight) & _
                DisplayBoom("60'", 62900, maxWeight) & _
                DisplayBoom("70'", 49800, maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 137100, maxWeight) & _
                DisplayBoom("50'", 89000, maxWeight) & _
                DisplayBoom("60'", 77000, maxWeight) & _
                DisplayBoom("70'", 74000, maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 137100, maxWeight) & _
                DisplayBoom("50'", 89000, maxWeight) & _
                DisplayBoom("60'", 77000, maxWeight) & _
                DisplayBoom("70'", 74400, maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 137100, maxWeight) & _
                DisplayBoom("47'", 110700, maxWeight) & _
                DisplayBoom("61.3'", 106900, maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
                
				msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 140500, maxWeight) & _
                DisplayBoom("50'", 89000, maxWeight) & _
                DisplayBoom("60'", 62900, maxWeight) & _
                DisplayBoom("70'", 49800, maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 140500, maxWeight) & _
                DisplayBoom("50'", 89000, maxWeight) & _
                DisplayBoom("60'", 77000, maxWeight) & _
                DisplayBoom("70'", 74000, maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 140500, maxWeight) & _
                DisplayBoom("50'", 89000, maxWeight) & _
                DisplayBoom("60'", 77000, maxWeight) & _
                DisplayBoom("70'", 74400, maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 140500, maxWeight) & _
                DisplayBoom("47'", 110700, maxWeight) & _
                DisplayBoom("61.3'", 106900, maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)      
				
            Case 15
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 88500, maxWeight) & _
                DisplayBoom("50'", 80000, maxWeight) & _
                DisplayBoom("60'", 53000, maxWeight) & _
                DisplayBoom("70'", 42200, maxWeight) & _
                DisplayBoom("80'", 42400, maxWeight) & _
                DisplayBoom("90'", 36700, maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 88500, maxWeight) & _
                DisplayBoom("50'", 80000, maxWeight) & _
                DisplayBoom("60'", 66000, maxWeight) & _
                DisplayBoom("70'", 68500, maxWeight) & _
                DisplayBoom("80'", 51900, maxWeight) & _
                DisplayBoom("90'", 48700, maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 88500, maxWeight) & _
                DisplayBoom("50'", 80000, maxWeight) & _
                DisplayBoom("60'", 66000, maxWeight) & _
                DisplayBoom("70'", 68700, maxWeight) & _
                DisplayBoom("80'", 62800, maxWeight) & _
                DisplayBoom("90'", 61400, maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 88500, maxWeight) & _
                DisplayBoom("47'", 110700, maxWeight) & _
                DisplayBoom("61.3'", 90900, maxWeight) & _
                DisplayBoom("76'", 77400, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 120500, maxWeight) & _
                DisplayBoom("50'", 80000, maxWeight) & _
                DisplayBoom("60'", 53000, maxWeight) & _
                DisplayBoom("70'", 42200, maxWeight) & _
                DisplayBoom("80'", 42400, maxWeight) & _
                DisplayBoom("90'", 36700, maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 94400, maxWeight) & _
                DisplayBoom("50'", 80000, maxWeight) & _
                DisplayBoom("60'", 66000, maxWeight) & _
                DisplayBoom("70'", 68500, maxWeight) & _
                DisplayBoom("80'", 51900, maxWeight) & _
                DisplayBoom("90'", 48700, maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 94400, maxWeight) & _
                DisplayBoom("50'", 80000, maxWeight) & _
                DisplayBoom("60'", 66000, maxWeight) & _
                DisplayBoom("70'", 68700, maxWeight) & _
                DisplayBoom("80'", 62800, maxWeight) & _
                DisplayBoom("90'", 61400, maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 94400, maxWeight) & _
                DisplayBoom("47'", 95800, maxWeight) & _
                DisplayBoom("61.3'", 96800, maxWeight) & _
                DisplayBoom("76'",77400 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)     		

            Case 20
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 63600, maxWeight) & _
                DisplayBoom("50'", 65500, maxWeight) & _
                DisplayBoom("60'", 45700, maxWeight) & _
                DisplayBoom("70'", 36400, maxWeight) & _
                DisplayBoom("80'", 37900, maxWeight) & _
                DisplayBoom("90'", 36700, maxWeight) & _
                DisplayBoom("100'", 30700, maxWeight) & _
                DisplayBoom("110'", 29100, maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 63600, maxWeight) & _
                DisplayBoom("50'", 65500, maxWeight) & _
                DisplayBoom("60'", 57800, maxWeight) & _
                DisplayBoom("70'", 60800, maxWeight) & _
                DisplayBoom("80'", 44900, maxWeight) & _
                DisplayBoom("90'", 42800, maxWeight) & _
                DisplayBoom("100'", 42500, maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 63600, maxWeight) & _
                DisplayBoom("50'", 65500, maxWeight) & _
                DisplayBoom("60'", 57800, maxWeight) & _
                DisplayBoom("70'", 60900, maxWeight) & _
                DisplayBoom("80'", 62800, maxWeight) & _
                DisplayBoom("90'", 57100, maxWeight) & _
                DisplayBoom("98.73'", 47900, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 63600, maxWeight) & _
                DisplayBoom("47'", 89900, maxWeight) & _
                DisplayBoom("61.3'", 66100, maxWeight) & _
                DisplayBoom("76'", 66000, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 94400, maxWeight) & _
                DisplayBoom("50'", 69900, maxWeight) & _
                DisplayBoom("60'", 45700, maxWeight) & _
                DisplayBoom("70'", 36400, maxWeight) & _
                DisplayBoom("80'", 37900, maxWeight) & _
                DisplayBoom("90'", 36700, maxWeight) & _
                DisplayBoom("100'",30700 , maxWeight) & _
                DisplayBoom("110'", 29100, maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 68000, maxWeight) & _
                DisplayBoom("50'", 69900, maxWeight) & _
                DisplayBoom("60'", 57800, maxWeight) & _
                DisplayBoom("70'", 60800, maxWeight) & _
                DisplayBoom("80'", 44900, maxWeight) & _
                DisplayBoom("90'", 42800, maxWeight) & _
                DisplayBoom("100'",42500 , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 68000, maxWeight) & _
                DisplayBoom("50'", 69900, maxWeight) & _
                DisplayBoom("60'", 57800, maxWeight) & _
                DisplayBoom("70'", 60900, maxWeight) & _
                DisplayBoom("80'", 62800, maxWeight) & _
                DisplayBoom("90'", 57100, maxWeight) & _
                DisplayBoom("98.73'", 47900, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 68000, maxWeight) & _
                DisplayBoom("47'", 69400, maxWeight) & _
                DisplayBoom("61.3'", 70600, maxWeight) & _
                DisplayBoom("76'",70400 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)     
				
            Case 25
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 48500, maxWeight) & _
                DisplayBoom("50'", 50400, maxWeight) & _
                DisplayBoom("60'", 40200, maxWeight) & _
                DisplayBoom("70'", 31900, maxWeight) & _
                DisplayBoom("80'", 33800, maxWeight) & _
                DisplayBoom("90'", 34500, maxWeight) & _
                DisplayBoom("100'", 30700, maxWeight) & _
                DisplayBoom("110'", 29100, maxWeight) & _
                DisplayBoom("120'", 25500, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 48500, maxWeight) & _
                DisplayBoom("50'", 50400, maxWeight) & _
                DisplayBoom("60'", 51400, maxWeight) & _
                DisplayBoom("70'", 52000, maxWeight) & _
                DisplayBoom("80'", 39500, maxWeight) & _
                DisplayBoom("90'", 38100, maxWeight) & _
                DisplayBoom("100'", 38300, maxWeight) & _
                DisplayBoom("113.08'", 32800, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 48500, maxWeight) & _
                DisplayBoom("50'", 50400, maxWeight) & _
                DisplayBoom("60'", 51400, maxWeight) & _
                DisplayBoom("70'", 51900, maxWeight) & _
                DisplayBoom("80'", 51700, maxWeight) & _
                DisplayBoom("90'", 49900, maxWeight) & _
                DisplayBoom("98.73'", 42500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 48500, maxWeight) & _
                DisplayBoom("47'", 64900, maxWeight) & _
                DisplayBoom("61.3'", 51200, maxWeight) & _
                DisplayBoom("76'", 51100, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 68000, maxWeight) & _
                DisplayBoom("50'", 54000, maxWeight) & _
                DisplayBoom("60'", 40200, maxWeight) & _
                DisplayBoom("70'", 31900, maxWeight) & _
                DisplayBoom("80'", 33800, maxWeight) & _
                DisplayBoom("90'", 34500, maxWeight) & _
                DisplayBoom("100'",30700 , maxWeight) & _
                DisplayBoom("110'", 29100, maxWeight) & _
                DisplayBoom("120'", 25500, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 52000, maxWeight) & _
                DisplayBoom("50'", 54000, maxWeight) & _
                DisplayBoom("60'", 51400, maxWeight) & _
                DisplayBoom("70'", 54800, maxWeight) & _
                DisplayBoom("80'", 39500, maxWeight) & _
                DisplayBoom("90'", 38100, maxWeight) & _
                DisplayBoom("100'",38300 , maxWeight) & _
                DisplayBoom("113.08'",32800 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 52000, maxWeight) & _
                DisplayBoom("50'", 54000, maxWeight) & _
                DisplayBoom("60'", 51400, maxWeight) & _
                DisplayBoom("70'", 54900, maxWeight) & _
                DisplayBoom("80'", 55300, maxWeight) & _
                DisplayBoom("90'", 49900, maxWeight) & _
                DisplayBoom("98.73'", 42500, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 52000, maxWeight) & _
                DisplayBoom("47'", 53400, maxWeight) & _
                DisplayBoom("61.3'", 54700, maxWeight) & _
                DisplayBoom("76'",54600 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)    
				
            Case 30
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 38300, maxWeight) & _
                DisplayBoom("50'", 40300, maxWeight) & _
                DisplayBoom("60'", 35900, maxWeight) & _
                DisplayBoom("70'", 28400, maxWeight) & _
                DisplayBoom("80'", 30300, maxWeight) & _
                DisplayBoom("90'", 30700, maxWeight) & _
                DisplayBoom("100'", 28400, maxWeight) & _
                DisplayBoom("110'", 29100, maxWeight) & _
                DisplayBoom("120'", 25500, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 38300, maxWeight) & _
                DisplayBoom("50'", 40300, maxWeight) & _
                DisplayBoom("60'", 41400, maxWeight) & _
                DisplayBoom("70'", 41900, maxWeight) & _
                DisplayBoom("80'", 35100, maxWeight) & _
                DisplayBoom("90'", 34100, maxWeight) & _
                DisplayBoom("100'", 34800, maxWeight) & _
                DisplayBoom("113.08'", 32100, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 38300, maxWeight) & _
                DisplayBoom("50'", 40300, maxWeight) & _
                DisplayBoom("60'", 41400, maxWeight) & _
                DisplayBoom("70'", 41900, maxWeight) & _
                DisplayBoom("80'", 41700, maxWeight) & _
                DisplayBoom("90'", 41600, maxWeight) & _
                DisplayBoom("98.73'", 38100, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 38300, maxWeight) & _
                DisplayBoom("47'", 49900, maxWeight) & _
                DisplayBoom("61.3'", 41000, maxWeight) & _
                DisplayBoom("76'", 41000, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 52000, maxWeight) & _
                DisplayBoom("50'", 43300, maxWeight) & _
                DisplayBoom("60'", 35900, maxWeight) & _
                DisplayBoom("70'", 28400, maxWeight) & _
                DisplayBoom("80'", 30300, maxWeight) & _
                DisplayBoom("90'", 30700, maxWeight) & _
                DisplayBoom("100'",28400 , maxWeight) & _
                DisplayBoom("110'", 29100, maxWeight) & _
                DisplayBoom("120'", 25500, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", 41300, maxWeight) & _
                DisplayBoom("50'", 43300, maxWeight) & _
                DisplayBoom("60'", 44400, maxWeight) & _
                DisplayBoom("70'", 44900, maxWeight) & _
                DisplayBoom("80'", 35100, maxWeight) & _
                DisplayBoom("90'", 34100, maxWeight) & _
                DisplayBoom("100'",34800 , maxWeight) & _
                DisplayBoom("113.08'",32100 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", 41300, maxWeight) & _
                DisplayBoom("50'", 43300, maxWeight) & _
                DisplayBoom("60'", 44400, maxWeight) & _
                DisplayBoom("70'", 44900, maxWeight) & _
                DisplayBoom("80'", 44700, maxWeight) & _
                DisplayBoom("90'", 44100, maxWeight) & _
                DisplayBoom("98.73'", 38100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", 41300, maxWeight) & _
                DisplayBoom("47'", 42700, maxWeight) & _
                DisplayBoom("61.3'", 44000, maxWeight) & _
                DisplayBoom("76'",44000 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 16200, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)    		

            Case 35
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 32700, maxWeight) & _
                DisplayBoom("60'", 32400, maxWeight) & _
                DisplayBoom("70'", 25600, maxWeight) & _
                DisplayBoom("80'", 27500, maxWeight) & _
                DisplayBoom("90'", 27700, maxWeight) & _
                DisplayBoom("100'", 25700, maxWeight) & _
                DisplayBoom("110'", 26600, maxWeight) & _
                DisplayBoom("120'", 2500, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 32700, maxWeight) & _
                DisplayBoom("60'", 33900, maxWeight) & _
                DisplayBoom("70'", 34400, maxWeight) & _
                DisplayBoom("80'", 31600, maxWeight) & _
                DisplayBoom("90'", 30800, maxWeight) & _
                DisplayBoom("100'", 31500, maxWeight) & _
                DisplayBoom("113.08'", 29200, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 32700, maxWeight) & _
                DisplayBoom("60'", 33900, maxWeight) & _
                DisplayBoom("70'", 34300, maxWeight) & _
                DisplayBoom("80'", 34100, maxWeight) & _
                DisplayBoom("90'", 33900, maxWeight) & _
                DisplayBoom("98.73'", 33800, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", 39700, maxWeight) & _
                DisplayBoom("61.3'", 33500, maxWeight) & _
                DisplayBoom("76'", 33300, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", 41300, maxWeight) & _
                DisplayBoom("50'", 35600, maxWeight) & _
                DisplayBoom("60'", 32400, maxWeight) & _
                DisplayBoom("70'", 25600, maxWeight) & _
                DisplayBoom("80'", 27500, maxWeight) & _
                DisplayBoom("90'", 27700, maxWeight) & _
                DisplayBoom("100'",25700 , maxWeight) & _
                DisplayBoom("110'", 26600, maxWeight) & _
                DisplayBoom("120'", 25500, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 35600, maxWeight) & _
                DisplayBoom("60'", 36700, maxWeight) & _
                DisplayBoom("70'", 37300, maxWeight) & _
                DisplayBoom("80'", 31600, maxWeight) & _
                DisplayBoom("90'", 30800, maxWeight) & _
                DisplayBoom("100'",31500 , maxWeight) & _
                DisplayBoom("113.08'",29200 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 35600, maxWeight) & _
                DisplayBoom("60'", 36700, maxWeight) & _
                DisplayBoom("70'", 37200, maxWeight) & _
                DisplayBoom("80'", 37100, maxWeight) & _
                DisplayBoom("90'", 37000, maxWeight) & _
                DisplayBoom("98.73'", 34300, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", 35000, maxWeight) & _
                DisplayBoom("61.3'", 36400, maxWeight) & _
                DisplayBoom("76'",36400 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10900, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 15600, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 10300, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)    	
				
            Case 40
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 25800, maxWeight) & _
                DisplayBoom("60'", 27400, maxWeight) & _
                DisplayBoom("70'", 23200, maxWeight) & _
                DisplayBoom("80'", 25200, maxWeight) & _
                DisplayBoom("90'", 25200, maxWeight) & _
                DisplayBoom("100'", 23400, maxWeight) & _
                DisplayBoom("110'", 24400, maxWeight) & _
                DisplayBoom("120'", 25000, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 25800, maxWeight) & _
                DisplayBoom("60'", 27100, maxWeight) & _
                DisplayBoom("70'", 27600, maxWeight) & _
                DisplayBoom("80'", 27900, maxWeight) & _
                DisplayBoom("90'", 27800, maxWeight) & _
                DisplayBoom("100'", 27600, maxWeight) & _
                DisplayBoom("113.08'", 26700, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 25800, maxWeight) & _
                DisplayBoom("60'", 27100, maxWeight) & _
                DisplayBoom("70'", 27500, maxWeight) & _
                DisplayBoom("80'", 27300, maxWeight) & _
                DisplayBoom("90'", 27100, maxWeight) & _
                DisplayBoom("98.73'", 26900, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", 32000, maxWeight) & _
                DisplayBoom("61.3'", 26700, maxWeight) & _
                DisplayBoom("76'", 26500, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 29800, maxWeight) & _
                DisplayBoom("60'", 29500, maxWeight) & _
                DisplayBoom("70'", 23200, maxWeight) & _
                DisplayBoom("80'", 25200, maxWeight) & _
                DisplayBoom("90'", 25200, maxWeight) & _
                DisplayBoom("100'",23400 , maxWeight) & _
                DisplayBoom("110'", 24400, maxWeight) & _
                DisplayBoom("120'", 25000, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 29800, maxWeight) & _
                DisplayBoom("60'", 31000, maxWeight) & _
                DisplayBoom("70'", 31500, maxWeight) & _
                DisplayBoom("80'", 28700, maxWeight) & _
                DisplayBoom("90'", 28100, maxWeight) & _
                DisplayBoom("100'",29000 , maxWeight) & _
                DisplayBoom("113.08'",26700 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", 29800, maxWeight) & _
                DisplayBoom("60'", 31000, maxWeight) & _
                DisplayBoom("70'", 31500, maxWeight) & _
                DisplayBoom("80'", 31300, maxWeight) & _
                DisplayBoom("90'", 31200, maxWeight) & _
                DisplayBoom("98.73'", 31100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", 18900, maxWeight) & _
                DisplayBoom("61.3'", 30600, maxWeight) & _
                DisplayBoom("76'",30600 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10900, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 14900, maxWeight) & _
                DisplayBoom("15'", 12500, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 9900, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)    			

            Case 45
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 22400, maxWeight) & _
                DisplayBoom("70'", 21300, maxWeight) & _
                DisplayBoom("80'", 23200, maxWeight) & _
                DisplayBoom("90'", 23100, maxWeight) & _
                DisplayBoom("100'", 21400, maxWeight) & _
                DisplayBoom("110'", 22500, maxWeight) & _
                DisplayBoom("120'", 23200, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 22100, maxWeight) & _
                DisplayBoom("70'", 22700, maxWeight) & _
                DisplayBoom("80'", 23000, maxWeight) & _
                DisplayBoom("90'", 23000, maxWeight) & _
                DisplayBoom("100'", 22800, maxWeight) & _
                DisplayBoom("113.08'", 22500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 22100, maxWeight) & _
                DisplayBoom("70'", 22700, maxWeight) & _
                DisplayBoom("80'", 22500, maxWeight) & _
                DisplayBoom("90'", 22300, maxWeight) & _
                DisplayBoom("98.73'", 22200, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", 18900, maxWeight) & _
                DisplayBoom("61.3'", 21800, maxWeight) & _
                DisplayBoom("76'", 21700, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 26300, maxWeight) & _
                DisplayBoom("70'", 21300, maxWeight) & _
                DisplayBoom("80'", 23200, maxWeight) & _
                DisplayBoom("90'", 23100, maxWeight) & _
                DisplayBoom("100'",21400 , maxWeight) & _
                DisplayBoom("110'", 22500, maxWeight) & _
                DisplayBoom("120'", 23200, maxWeight) & _
                DisplayBoom("127'", 22600, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 26000, maxWeight) & _
                DisplayBoom("70'", 26600, maxWeight) & _
                DisplayBoom("80'", 26200, maxWeight) & _
                DisplayBoom("90'", 25800, maxWeight) & _
                DisplayBoom("100'",26600 , maxWeight) & _
                DisplayBoom("113.08'",24500 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 26000, maxWeight) & _
                DisplayBoom("70'", 26500, maxWeight) & _
                DisplayBoom("80'", 26300, maxWeight) & _
                DisplayBoom("90'", 26100, maxWeight) & _
                DisplayBoom("98.73'", 26000, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", 25700, maxWeight) & _
                DisplayBoom("76'",25600 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10900, maxWeight) & _
                DisplayBoom("15'", 10400, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 14300, maxWeight) & _
                DisplayBoom("15'", 12000, maxWeight) & _
                DisplayBoom("30'", 10100, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 7300, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 9400, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)    	
				
           Case 50
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 18600, maxWeight) & _
                DisplayBoom("70'", 19600, maxWeight) & _
                DisplayBoom("80'", 19900, maxWeight) & _
                DisplayBoom("90'", 20100, maxWeight) & _
                DisplayBoom("100'", 19700, maxWeight) & _
                DisplayBoom("110'", 19800, maxWeight) & _
                DisplayBoom("120'", 19600, maxWeight) & _
                DisplayBoom("127'", 19500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 18300, maxWeight) & _
                DisplayBoom("70'", 18900, maxWeight) & _
                DisplayBoom("80'", 19300, maxWeight) & _
                DisplayBoom("90'", 19300, maxWeight) & _
                DisplayBoom("100'", 19100, maxWeight) & _
                DisplayBoom("113.08'", 18800, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 18300, maxWeight) & _
                DisplayBoom("70'", 18900, maxWeight) & _
                DisplayBoom("80'", 18700, maxWeight) & _
                DisplayBoom("90'", 18600, maxWeight) & _
                DisplayBoom("98.73'", 18400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", 18000, maxWeight) & _
                DisplayBoom("76'", 18000, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 22100, maxWeight) & _
                DisplayBoom("70'", 19600, maxWeight) & _
                DisplayBoom("80'", 21500, maxWeight) & _
                DisplayBoom("90'", 21200, maxWeight) & _
                DisplayBoom("100'",19700 , maxWeight) & _
                DisplayBoom("110'", 20900, maxWeight) & _
                DisplayBoom("120'", 21600, maxWeight) & _
                DisplayBoom("127'", 21600, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 21800, maxWeight) & _
                DisplayBoom("70'", 22400, maxWeight) & _
                DisplayBoom("80'", 22800, maxWeight) & _
                DisplayBoom("90'", 22700, maxWeight) & _
                DisplayBoom("100'",22500 , maxWeight) & _
                DisplayBoom("113.08'",22300 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", 21800, maxWeight) & _
                DisplayBoom("70'", 22400, maxWeight) & _
                DisplayBoom("80'", 22200, maxWeight) & _
                DisplayBoom("90'", 22100, maxWeight) & _
                DisplayBoom("98.73'", 21900, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", 21500, maxWeight) & _
                DisplayBoom("76'",21500 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10900, maxWeight) & _
                DisplayBoom("15'", 10300, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 13600, maxWeight) & _
                DisplayBoom("15'", 11500, maxWeight) & _
                DisplayBoom("30'", 9800, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 7300, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 9000, maxWeight) & _
                DisplayBoom("15'", 7400, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)				
				
           Case 55
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 16600, maxWeight) & _
                DisplayBoom("80'", 16900, maxWeight) & _
                DisplayBoom("90'", 17000, maxWeight) & _
                DisplayBoom("100'", 17100, maxWeight) & _
                DisplayBoom("110'", 16800, maxWeight) & _
                DisplayBoom("120'", 16600, maxWeight) & _
                DisplayBoom("127'", 16400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 15900, maxWeight) & _
                DisplayBoom("80'", 16300, maxWeight) & _
                DisplayBoom("90'", 16200, maxWeight) & _
                DisplayBoom("100'", 16000, maxWeight) & _
                DisplayBoom("113.08'", 15800, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 15900, maxWeight) & _
                DisplayBoom("80'", 15700, maxWeight) & _
                DisplayBoom("90'", 15500, maxWeight) & _
                DisplayBoom("98.73'", 15400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", 15000, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 18200, maxWeight) & _
                DisplayBoom("80'", 20000, maxWeight) & _
                DisplayBoom("90'", 19600, maxWeight) & _
                DisplayBoom("100'",18200 , maxWeight) & _
                DisplayBoom("110'", 19400, maxWeight) & _
                DisplayBoom("120'", 19800, maxWeight) & _
                DisplayBoom("127'", 19400, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 19100, maxWeight) & _
                DisplayBoom("80'", 19500, maxWeight) & _
                DisplayBoom("90'", 19500, maxWeight) & _
                DisplayBoom("100'",19300 , maxWeight) & _
                DisplayBoom("113.08'",19000 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 19000, maxWeight) & _
                DisplayBoom("80'", 18900, maxWeight) & _
                DisplayBoom("90'", 18800, maxWeight) & _
                DisplayBoom("98.73'", 18700, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'",18200 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10900, maxWeight) & _
                DisplayBoom("15'", 10100, maxWeight) & _
                DisplayBoom("30'", 9100, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 13000, maxWeight) & _
                DisplayBoom("15'", 11100, maxWeight) & _
                DisplayBoom("30'", 9500, maxWeight) & _
                DisplayBoom("45'", 8600, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 7200, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 8500, maxWeight) & _
                DisplayBoom("15'", 7000, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)				

           Case 60
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 14100, maxWeight) & _
                DisplayBoom("80'", 14400, maxWeight) & _
                DisplayBoom("90'", 14600, maxWeight) & _
                DisplayBoom("100'", 14600, maxWeight) & _
                DisplayBoom("110'", 14400, maxWeight) & _
                DisplayBoom("120'", 14100, maxWeight) & _
                DisplayBoom("127'", 14000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 13500, maxWeight) & _
                DisplayBoom("80'", 13900, maxWeight) & _
                DisplayBoom("90'", 13800, maxWeight) & _
                DisplayBoom("100'", 13600, maxWeight) & _
                DisplayBoom("113.08'", 13400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 13400, maxWeight) & _
                DisplayBoom("80'", 13300, maxWeight) & _
                DisplayBoom("90'", 13100, maxWeight) & _
                DisplayBoom("98.73'", 13000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", 12600, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 17000, maxWeight) & _
                DisplayBoom("80'", 17400, maxWeight) & _
                DisplayBoom("90'", 17600, maxWeight) & _
                DisplayBoom("100'",16900 , maxWeight) & _
                DisplayBoom("110'", 17400, maxWeight) & _
                DisplayBoom("120'", 17200, maxWeight) & _
                DisplayBoom("127'", 17000, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 16400, maxWeight) & _
                DisplayBoom("80'", 16800, maxWeight) & _
                DisplayBoom("90'", 16800, maxWeight) & _
                DisplayBoom("100'",16600 , maxWeight) & _
                DisplayBoom("113.08'",16400 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", 16400, maxWeight) & _
                DisplayBoom("80'", 16300, maxWeight) & _
                DisplayBoom("90'", 16200, maxWeight) & _
                DisplayBoom("98.73'", 16100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'",15600 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10900, maxWeight) & _
                DisplayBoom("15'", 9900, maxWeight) & _
                DisplayBoom("30'", 8900, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 12500, maxWeight) & _
                DisplayBoom("15'", 10700, maxWeight) & _
                DisplayBoom("30'", 9300, maxWeight) & _
                DisplayBoom("45'", 8400, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 7100, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 8100, maxWeight) & _
                DisplayBoom("15'", 6700, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
				
          Case 65
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 12500, maxWeight) & _
                DisplayBoom("90'", 12700, maxWeight) & _
                DisplayBoom("100'", 12700, maxWeight) & _
                DisplayBoom("110'", 12500, maxWeight) & _
                DisplayBoom("120'", 12300, maxWeight) & _
                DisplayBoom("127'", 12100, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 11900, maxWeight) & _
                DisplayBoom("90'", 11900, maxWeight) & _
                DisplayBoom("100'", 11700, maxWeight) & _
                DisplayBoom("113.08'", 11500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 11400, maxWeight) & _
                DisplayBoom("90'", 11200, maxWeight) & _
                DisplayBoom("98.73'", 11100, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", 10600, maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 15200, maxWeight) & _
                DisplayBoom("90'", 15400, maxWeight) & _
                DisplayBoom("100'",15500 , maxWeight) & _
                DisplayBoom("110'", 15200, maxWeight) & _
                DisplayBoom("120'", 15000, maxWeight) & _
                DisplayBoom("127'", 14900, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 14600, maxWeight) & _
                DisplayBoom("90'", 14700, maxWeight) & _
                DisplayBoom("100'",14500 , maxWeight) & _
                DisplayBoom("113.08'",14300 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 14100, maxWeight) & _
                DisplayBoom("90'", 14000, maxWeight) & _
                DisplayBoom("98.73'", 13900, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'",13400 , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10700, maxWeight) & _
                DisplayBoom("15'", 9700, maxWeight) & _
                DisplayBoom("30'", 8700, maxWeight) & _
                DisplayBoom("45'", 8100, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 11900, maxWeight) & _
                DisplayBoom("15'", 10300, maxWeight) & _
                DisplayBoom("30'", 9000, maxWeight) & _
                DisplayBoom("45'", 8300, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6900, maxWeight) & _
                DisplayBoom("15'", 6100, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 7700, maxWeight) & _
                DisplayBoom("15'", 6400, maxWeight) & _
                DisplayBoom("30'", 5400, maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
				
          Case 70
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 10800, maxWeight) & _
                DisplayBoom("90'", 11100, maxWeight) & _
                DisplayBoom("100'", 11100, maxWeight) & _
                DisplayBoom("110'", 10800, maxWeight) & _
                DisplayBoom("120'", 10600, maxWeight) & _
                DisplayBoom("127'", 10500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 10200, maxWeight) & _
                DisplayBoom("90'", 10300, maxWeight) & _
                DisplayBoom("100'", 10100, maxWeight) & _
                DisplayBoom("113.08'", 9900, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 9700, maxWeight) & _
                DisplayBoom("90'", 9600, maxWeight) & _
                DisplayBoom("98.73'", 9500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 13300, maxWeight) & _
                DisplayBoom("90'", 13600, maxWeight) & _
                DisplayBoom("100'",13600 , maxWeight) & _
                DisplayBoom("110'", 13400, maxWeight) & _
                DisplayBoom("120'", 13200, maxWeight) & _
                DisplayBoom("127'", 13100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 12800, maxWeight) & _
                DisplayBoom("90'", 12800, maxWeight) & _
                DisplayBoom("100'",12800 , maxWeight) & _
                DisplayBoom("113.08'",12600 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", 12300, maxWeight) & _
                DisplayBoom("90'", 12300, maxWeight) & _
                DisplayBoom("98.73'", 12200, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10500, maxWeight) & _
                DisplayBoom("15'", 9500, maxWeight) & _
                DisplayBoom("30'", 8600, maxWeight) & _
                DisplayBoom("45'", 8000, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 11400, maxWeight) & _
                DisplayBoom("15'", 10000, maxWeight) & _
                DisplayBoom("30'", 8800, maxWeight) & _
                DisplayBoom("45'", 8100, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6800, maxWeight) & _
                DisplayBoom("15'", 6000, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 7300, maxWeight) & _
                DisplayBoom("15'", 6200, maxWeight) & _
                DisplayBoom("30'", 5200, maxWeight) & _
                DisplayBoom("45'", , maxWeight)						

          Case 75
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 9600, maxWeight) & _
                DisplayBoom("100'", 9700, maxWeight) & _
                DisplayBoom("110'", 9500, maxWeight) & _
                DisplayBoom("120'", 9300, maxWeight) & _
                DisplayBoom("127'", 9100, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 8900, maxWeight) & _
                DisplayBoom("100'", 8700, maxWeight) & _
                DisplayBoom("113.08'", 8500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 8200, maxWeight) & _
                DisplayBoom("98.73'", 8200, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 12100, maxWeight) & _
                DisplayBoom("100'",12200 , maxWeight) & _
                DisplayBoom("110'", 11900, maxWeight) & _
                DisplayBoom("120'", 11700, maxWeight) & _
                DisplayBoom("127'", 11600, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 11300, maxWeight) & _
                DisplayBoom("100'",11200 , maxWeight) & _
                DisplayBoom("113.08'",11000 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 10700, maxWeight) & _
                DisplayBoom("98.73'", 10700, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 10200, maxWeight) & _
                DisplayBoom("15'", 9300, maxWeight) & _
                DisplayBoom("30'", 8500, maxWeight) & _
                DisplayBoom("45'", 7900, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 11000, maxWeight) & _
                DisplayBoom("15'", 9700, maxWeight) & _
                DisplayBoom("30'", 8600, maxWeight) & _
                DisplayBoom("45'", 8000, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6600, maxWeight) & _
                DisplayBoom("15'", 5800, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 7000, maxWeight) & _
                DisplayBoom("15'", 5900, maxWeight) & _
                DisplayBoom("30'", 5100, maxWeight) & _
                DisplayBoom("45'", , maxWeight)						

          Case 80
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 8400, maxWeight) & _
                DisplayBoom("100'", 8500, maxWeight) & _
                DisplayBoom("110'", 8300, maxWeight) & _
                DisplayBoom("120'", 8100, maxWeight) & _
                DisplayBoom("127'", 8000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 7700, maxWeight) & _
                DisplayBoom("100'", 7600, maxWeight) & _
                DisplayBoom("113.08'", 7400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 7000, maxWeight) & _
                DisplayBoom("98.73'", 7000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 10700, maxWeight) & _
                DisplayBoom("100'",10800 , maxWeight) & _
                DisplayBoom("110'", 10600, maxWeight) & _
                DisplayBoom("120'", 10400, maxWeight) & _
                DisplayBoom("127'", 10300, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 10000, maxWeight) & _
                DisplayBoom("100'",9900 , maxWeight) & _
                DisplayBoom("113.08'",9700 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", 9400, maxWeight) & _
                DisplayBoom("98.73'", 9300, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 9600, maxWeight) & _
                DisplayBoom("15'", 9000, maxWeight) & _
                DisplayBoom("30'", 8300, maxWeight) & _
                DisplayBoom("45'", 7700, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 10500, maxWeight) & _
                DisplayBoom("15'", 9400, maxWeight) & _
                DisplayBoom("30'", 8500, maxWeight) & _
                DisplayBoom("45'", 7900, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6500, maxWeight) & _
                DisplayBoom("15'", 5600, maxWeight) & _
                DisplayBoom("30'", 4900, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 6700, maxWeight) & _
                DisplayBoom("15'", 5700, maxWeight) & _
                DisplayBoom("30'", 4900, maxWeight) & _
                DisplayBoom("45'", 4400, maxWeight)						

          Case 85
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", 7500, maxWeight) & _
                DisplayBoom("110'", 7300, maxWeight) & _
                DisplayBoom("120'", 7100, maxWeight) & _
                DisplayBoom("127'", 7000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", 6500, maxWeight) & _
                DisplayBoom("113.08'", 6400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", 6000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'",9600 , maxWeight) & _
                DisplayBoom("110'", 9400, maxWeight) & _
                DisplayBoom("120'", 9200, maxWeight) & _
                DisplayBoom("127'", 9100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'",8700 , maxWeight) & _
                DisplayBoom("113.08'",8500 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", 8100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 9100, maxWeight) & _
                DisplayBoom("15'", 8500, maxWeight) & _
                DisplayBoom("30'", 8000, maxWeight) & _
                DisplayBoom("45'", 7700, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 9400, maxWeight) & _
                DisplayBoom("15'", 9100, maxWeight) & _
                DisplayBoom("30'", 8300, maxWeight) & _
                DisplayBoom("45'", 7800, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6300, maxWeight) & _
                DisplayBoom("15'", 5500, maxWeight) & _
                DisplayBoom("30'", 4800, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 6400, maxWeight) & _
                DisplayBoom("15'", 5500, maxWeight) & _
                DisplayBoom("30'", 4800, maxWeight) & _
                DisplayBoom("45'", 4300, maxWeight)						

          Case 90
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", 6600, maxWeight) & _
                DisplayBoom("110'", 6400, maxWeight) & _
                DisplayBoom("120'", 6200, maxWeight) & _
                DisplayBoom("127'", 6100, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", 5600, maxWeight) & _
                DisplayBoom("113.08'", 5500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", 5100, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'",8600 , maxWeight) & _
                DisplayBoom("110'", 8400, maxWeight) & _
                DisplayBoom("120'", 8200, maxWeight) & _
                DisplayBoom("127'", 8100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'",7600 , maxWeight) & _
                DisplayBoom("113.08'",7500 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", 7100, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 8600, maxWeight) & _
                DisplayBoom("15'", 8100, maxWeight) & _
                DisplayBoom("30'", 7700, maxWeight) & _
                DisplayBoom("45'", 7400, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 8400, maxWeight) & _
                DisplayBoom("15'", 8800, maxWeight) & _
                DisplayBoom("30'", 8200, maxWeight) & _
                DisplayBoom("45'", 7700, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6100, maxWeight) & _
                DisplayBoom("15'", 5300, maxWeight) & _
                DisplayBoom("30'", 4700, maxWeight) & _
                DisplayBoom("45'", 4200, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 6100, maxWeight) & _
                DisplayBoom("15'", 5300, maxWeight) & _
                DisplayBoom("30'", 4700, maxWeight) & _
                DisplayBoom("45'", 4300, maxWeight)		
				
         Case 95
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", 5600, maxWeight) & _
                DisplayBoom("120'", 5400, maxWeight) & _
                DisplayBoom("127'", 5300, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", 4700, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", 7500, maxWeight) & _
                DisplayBoom("120'", 7300, maxWeight) & _
                DisplayBoom("127'", 7200, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'",6600 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 7700, maxWeight) & _
                DisplayBoom("15'", 7800, maxWeight) & _
                DisplayBoom("30'", 7400, maxWeight) & _
                DisplayBoom("45'", 7200, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 7500, maxWeight) & _
                DisplayBoom("15'", 7800, maxWeight) & _
                DisplayBoom("30'", 8000, maxWeight) & _
                DisplayBoom("45'", 7600, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 5900, maxWeight) & _
                DisplayBoom("15'", 5200, maxWeight) & _
                DisplayBoom("30'", 4600, maxWeight) & _
                DisplayBoom("45'", 4200, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 5800, maxWeight) & _
                DisplayBoom("15'", 5100, maxWeight) & _
                DisplayBoom("30'", 4600, maxWeight) & _
                DisplayBoom("45'", 4200, maxWeight)		
				
		Case 100
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", 4900, maxWeight) & _
                DisplayBoom("120'", 4700, maxWeight) & _
                DisplayBoom("127'", 4600, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", 4000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", 6700, maxWeight) & _
                DisplayBoom("120'", 6500, maxWeight) & _
                DisplayBoom("127'", 6400, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'",5800 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6900, maxWeight) & _
                DisplayBoom("15'", 7300, maxWeight) & _
                DisplayBoom("30'", 7100, maxWeight) & _
                DisplayBoom("45'", 6900, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 6700, maxWeight) & _
                DisplayBoom("15'", 7000, maxWeight) & _
                DisplayBoom("30'", 7300, maxWeight) & _
                DisplayBoom("45'", 7500, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 5700, maxWeight) & _
                DisplayBoom("15'", 5100, maxWeight) & _
                DisplayBoom("30'", 4500, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 5600, maxWeight) & _
                DisplayBoom("15'", 5000, maxWeight) & _
                DisplayBoom("30'", 4500, maxWeight) & _
                DisplayBoom("45'", 4200, maxWeight)		
				
				
		Case 105
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", 4100, maxWeight) & _
                DisplayBoom("127'", 4000, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", 3400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", 5800, maxWeight) & _
                DisplayBoom("127'", 5700, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'",5100 , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 6200, maxWeight) & _
                DisplayBoom("15'", 6500, maxWeight) & _
                DisplayBoom("30'", 6800, maxWeight) & _
                DisplayBoom("45'", 6700, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 5900, maxWeight) & _
                DisplayBoom("15'", 6200, maxWeight) & _
                DisplayBoom("30'", 6500, maxWeight) & _
                DisplayBoom("45'", 6600, maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 5400, maxWeight) & _
                DisplayBoom("15'", 4900, maxWeight) & _
                DisplayBoom("30'", 4400, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 5400, maxWeight) & _
                DisplayBoom("15'", 4800, maxWeight) & _
                DisplayBoom("30'", 4400, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight)		
				
		Case 110
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", 3500, maxWeight) & _
                DisplayBoom("127'", 3400, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", 5100, maxWeight) & _
                DisplayBoom("127'", 5000, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 5500, maxWeight) & _
                DisplayBoom("15'", 5800, maxWeight) & _
                DisplayBoom("30'", 6100, maxWeight) & _
                DisplayBoom("45'", 6300, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 5300, maxWeight) & _
                DisplayBoom("15'", 5500, maxWeight) & _
                DisplayBoom("30'", 5700, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 5100, maxWeight) & _
                DisplayBoom("15'", 4700, maxWeight) & _
                DisplayBoom("30'", 4300, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 5200, maxWeight) & _
                DisplayBoom("15'", 4700, maxWeight) & _
                DisplayBoom("30'", 4300, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight)		
								
		Case 115
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", 2900, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", 4500, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 4900, maxWeight) & _
                DisplayBoom("15'", 5200, maxWeight) & _
                DisplayBoom("30'", 5500, maxWeight) & _
                DisplayBoom("45'", 5600, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 4700, maxWeight) & _
                DisplayBoom("15'", 4900, maxWeight) & _
                DisplayBoom("30'", 5100, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 4900, maxWeight) & _
                DisplayBoom("15'", 4500, maxWeight) & _
                DisplayBoom("30'", 4200, maxWeight) & _
                DisplayBoom("45'", 4000, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 5000, maxWeight) & _
                DisplayBoom("15'", 4600, maxWeight) & _
                DisplayBoom("30'", 4200, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight)		
								
				
		Case 120
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", 2500, maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", 2700, maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 4400, maxWeight) & _
                DisplayBoom("15'", 4700, maxWeight) & _
                DisplayBoom("30'", 4900, maxWeight) & _
                DisplayBoom("45'", 5000, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 4200, maxWeight) & _
                DisplayBoom("15'", 4400, maxWeight) & _
                DisplayBoom("30'", 4400, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 4700, maxWeight) & _
                DisplayBoom("15'", 4300, maxWeight) & _
                DisplayBoom("30'", 4000, maxWeight) & _
                DisplayBoom("45'", 3900, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 4900, maxWeight) & _
                DisplayBoom("15'", 4500, maxWeight) & _
                DisplayBoom("30'", 4200, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight)		
								

		Case 125
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 3900, maxWeight) & _
                DisplayBoom("15'", 4200, maxWeight) & _
                DisplayBoom("30'", 4400, maxWeight) & _
                DisplayBoom("45'", 4400, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 3700, maxWeight) & _
                DisplayBoom("15'", 3800, maxWeight) & _
                DisplayBoom("30'", 2100, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 4500, maxWeight) & _
                DisplayBoom("15'", 4100, maxWeight) & _
                DisplayBoom("30'", 3900, maxWeight) & _
                DisplayBoom("45'", 3800, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 4500, maxWeight) & _
                DisplayBoom("15'", 4400, maxWeight) & _
                DisplayBoom("30'", 4100, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight)		
								
		Case 130
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 3500, maxWeight) & _
                DisplayBoom("15'", 3700, maxWeight) & _
                DisplayBoom("30'", 3900, maxWeight) & _
                DisplayBoom("45'", 3900, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 4000, maxWeight) & _
                DisplayBoom("15'", 4000, maxWeight) & _
                DisplayBoom("30'", 3700, maxWeight) & _
                DisplayBoom("45'", 3600, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 4100, maxWeight) & _
                DisplayBoom("15'", 4300, maxWeight) & _
                DisplayBoom("30'", 4100, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight)		
								
		Case 135
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 3100, maxWeight) & _
                DisplayBoom("15'", 3300, maxWeight) & _
                DisplayBoom("30'", 3400, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 3600, maxWeight) & _
                DisplayBoom("15'", 3800, maxWeight) & _
                DisplayBoom("30'", 3600, maxWeight) & _
                DisplayBoom("45'", 3500, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 3700, maxWeight) & _
                DisplayBoom("15'", 3900, maxWeight) & _
                DisplayBoom("30'", 4100, maxWeight) & _
                DisplayBoom("45'", 4100, maxWeight)		
																
		Case 140
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 2700, maxWeight) & _
                DisplayBoom("15'", 2900, maxWeight) & _
                DisplayBoom("30'", 3000, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 3200, maxWeight) & _
                DisplayBoom("15'", 3600, maxWeight) & _
                DisplayBoom("30'", 3500, maxWeight) & _
                DisplayBoom("45'", 3400, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 3300, maxWeight) & _
                DisplayBoom("15'", 3500, maxWeight) & _
                DisplayBoom("30'", 3700, maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
		Case 145
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 2400, maxWeight) & _
                DisplayBoom("15'", 2500, maxWeight) & _
                DisplayBoom("30'", 2600, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 2900, maxWeight) & _
                DisplayBoom("15'", 3200, maxWeight) & _
                DisplayBoom("30'", 3400, maxWeight) & _
                DisplayBoom("45'", 3400, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 2900, maxWeight) & _
                DisplayBoom("15'", 3100, maxWeight) & _
                DisplayBoom("30'", 3200, maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
		Case 150
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 2100, maxWeight) & _
                DisplayBoom("15'", 2200, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 2600, maxWeight) & _
                DisplayBoom("15'", 2900, maxWeight) & _
                DisplayBoom("30'", 3100, maxWeight) & _
                DisplayBoom("45'", 3300, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 2600, maxWeight) & _
                DisplayBoom("15'", 2800, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
		Case 155
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 1800, maxWeight) & _
                DisplayBoom("15'", 1800, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 2300, maxWeight) & _
                DisplayBoom("15'", 2500, maxWeight) & _
                DisplayBoom("30'", 2800, maxWeight) & _
                DisplayBoom("45'", 2800, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", 2300, maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
		Case 160
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 2000, maxWeight) & _
                DisplayBoom("15'", 2200, maxWeight) & _
                DisplayBoom("30'", 2400, maxWeight) & _
                DisplayBoom("45'", 2400, maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
		Case 165
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 1700, maxWeight) & _
                DisplayBoom("15'", 1900, maxWeight) & _
                DisplayBoom("30'", 2100, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
		Case 170
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 1500, maxWeight) & _
                DisplayBoom("15'", 1700, maxWeight) & _
                DisplayBoom("30'", 1700, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
		Case 175
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 1300, maxWeight) & _
                DisplayBoom("15'", 1400, maxWeight) & _
                DisplayBoom("30'", 1400, maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
						
		Case 180
                msg1 = "HTC-8675 II - 10,800 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 10,800 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) 
				
                msg2 = "HTC-8675 II - 18,400 CW Main Boom EM1" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("110'", , maxWeight) & _
                DisplayBoom("120'", , maxWeight) & _
                DisplayBoom("127'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM2" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("100'", , maxWeight) & _
                DisplayBoom("113.08'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM3" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("50'", , maxWeight) & _
                DisplayBoom("60'", , maxWeight) & _
                DisplayBoom("70'", , maxWeight) & _
                DisplayBoom("80'", , maxWeight) & _
                DisplayBoom("90'", , maxWeight) & _
                DisplayBoom("98.73'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW Main Boom EM4" & vbCrLf & _
                DisplayBoom("41'", , maxWeight) & _
                DisplayBoom("47'", , maxWeight) & _
                DisplayBoom("61.3'", , maxWeight) & _
                DisplayBoom("76'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 38' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 38' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 127' Main 64' Jib EM1" & vbCrLf & _
                DisplayBoom("2'", 1100, maxWeight) & _
                DisplayBoom("15'", 1100, maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight) & _
                "HTC-8675 II - 18,400 CW 98.73' Main 64' Jib EM3" & vbCrLf & _
                DisplayBoom("2'", , maxWeight) & _
                DisplayBoom("15'", , maxWeight) & _
                DisplayBoom("30'", , maxWeight) & _
                DisplayBoom("45'", , maxWeight)		
						
						
						
            Case Else
                MsgBox = "Invalid radius value. Please enter a value in increments of 5 starting at 10."
        End Select
    Else
        MsgBox = "Please enter a valid numeric radius."
    End If
Else
    MsgBox = "Please enter a valid numeric maximum weight."
End If

If opt=1 Then
	MsgBox msg1
Else
	MsgBox msg2
End If

Function DisplayBoom(boomLength, boomWeight, maxWeight)
    If IsNumeric(boomWeight) Then
        If boomWeight >= maxWeight Then
            Dim percentage
            percentage = (maxWeight / boomWeight) * 100
            DisplayBoom = boomLength & " = " & FormatNumber(boomWeight, 0) & "#" & " (" & FormatNumber(percentage, 2) & "%)" & vbCrLf
        Else
            DisplayBoom = ""
        End If
    Else
        DisplayBoom = ""
    End If
End Function
