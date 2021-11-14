# Stock-Analysis
Challange_Mod2_VBA
# Stock Analysis using Excel VBA

## Overview of Project: Explain the purpose of this analysis.
### Background 
Steve recently graduated with a finance degree and wanted to use his knowledge to assist/ advise his parents with their investment decisions. Unfortunately, they had invested all their money into one stock, based on their opinions of the company and reusable energy sources. Not company performance and the stock’s return. 
### Purpose 
To reduce the number of errors and mistakes, as well as create a tool that Steve can utilize to analyses many stock types and hopefully diversify his parents’ weak investment portfolio, we created Steve an analysis macro in VBA. Once the stock data is uploaded Steve can now, by stock type, see the daily volume and return.  
## Results 
### Comparison of stock performance between 2017 and 2018
Overall, the stocks performed much better in 2017 than 2018. Of the 12 stocks analyzed in 2017 all but TERP had a positive return. And four of the stocks (DQ, ENPH, FSLR and SEDG) had a positive return greater than 100%! In 2018 of these 12 same stocks 10 had negative returns. Meaning only two stocks had a positive return for a second year (ENPH and RUN). This would allow us to conclude that ENPH and RUN are strong stocks. But specifically, ENPH is an extremely stable stock to invest in, with exact returns of 129.5% in 2017 and 81.9% in 2018. Total daily volumes for all stocks between both years stayed fairly consistent, with the average difference in total volumes from 2017 to 2018 being 11,616,592. For context the lowest total daily volume was DQ’s stock in 2017 at 35,796,200, and most others exceed 100,000,000.   
### Comparison of execution times 
Thankfully refactoring was successful and our code is more efficient which can easily be seen by comparing the execution times. With refactored code the 2017 data took only 6.6406 seconds ![VBA_Challange_2017.png](C:\Users\Sydney Kieswetter\Desktop\UofT_Class_Work\2_VBA\Final_Submission\Resources) to run and 2018’s took 7.4218![VBA_Challange_2018.png](C:\Users\Sydney Kieswetter\Desktop\UofT_Class_Work\2_VBA\Final_Submission\Resources). This is in comparison to 57669.31 seconds for 2017 using the original code, and 57691.82 seconds for 2018. You could even see a visual different with how fast the text and the table refreshed when running the refactored macro.  
## Summary of Refactoring 
As shown in the execution time comparison above refactoring code is great to be able to increase it’s efficiently and therefore process more data as well as be able to adapt to data. For example Steve can easily use different stock types in the future now. However, a detailed plan is definitely needed to manipulate the original code, and a lot of debugging and review of the code may be needed to ensure your loops especially are working.  
