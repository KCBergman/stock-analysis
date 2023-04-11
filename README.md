# stock-analysis
This repository contains my stock analysis work for module 2 of my data analytics bootcamp.

The file KendalBergman_Challenge2_VBA_Script_Final.vb is the file containing my VBA script, which ran on the course-provided excel file.
I have included 3 screenshots showing how my script ran on each year tab of the course-provided excel file. 

In order to complete this assignment I worked with a tutor, Kourt Bailey, as well as looked at the Ask The Class channel within slack. In particular, the following code(s) from my classmate Kirsten Rain were helpful in completing my assignment:

  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         StockName = Cells(i, 1).Value
         Total_Volume = Total_Volume + Cells(i, 7).Value
         ClosePrice = Cells(i, 6).Value
         PriceDiff = ClosePrice – OpenPrice
  ‘Write the stuff to the summary table like you started
  End If
  OpenPrice = Cells(i+1,3).value
  Summary_Table_Row = Summary_Table_Row+1
  Total_Volume = 0
  Else
  ‘if same keep totaling
  Total_Volume = Total_Volume +Cells(I,7).value
  End if
  Next i
  
  For j = 2 To Summary_Row
    If Cells(j, 11).Value > Greatest_Increase Then
        Greatest_Increase = Cells(j, 11).Value
        TickerInc = Cells(j, 9).Value
    End If
    If Cells(j, 11).Value < Greatest_Decrease Then
        Greatest_Decrease = Cells(j, 11).Value
        TickerDec = Cells(j, 9).Value
    End If
    If Cells(j, 12) > Greatest_Volume Then
        Greatest_Volume = Cells(j, 12).Value
        TickerVol = Cells(j, 9).Value
    End If
Next j

I also worked with TA Erik Nam and was able to find out how to use a For each loop to run my code across all worksheets in a workbook. 
I also used the internet to find a code for determining the last row in a sheet and setting that as part of a for loop. I found that information here: https://www.wallstreetmojo.com/vba-last-row/ 
