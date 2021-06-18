# Stock Analysis
## Overview
#### The purpose of this project was to determine if DAQO was an appropriate stock purchase for Steve and his parents as compared to other renewable energy companies. A handful of companies were looked at in the years 2017 and 2018 in order to decide what Steve should do with his stocks. The project used VBA code in Excel to automate the process, so Steve could simply click a button, depending on the year, and determine what the best decision would be in the long run.
---
## Results
#### For Steve and his parents, the best choice of stocks was not DAQO. While DAQO had almost a 200% growth in 2017, with more recent data from 2018, the stock stumbled 63%. When looking at stocks, the progression of the company is important to keep in mind, so proving DAQO was not doing well in 2018 shows that Steve should move his funds. The most logical stocks for Steve to purchase would be ENPH and RUN. ENPH grew 129.5% in 2017 and 81.9% in 2018. RUN grew 5.5% in 2017 and 84.0% in 2018. These are the only two companies that grew in both fiscal years. This data can be seen in the below tables.
![2017 Data](https://user-images.githubusercontent.com/85752084/122598377-fce47800-d029-11eb-9554-05799efb7a2a.PNG) ![2018 Data](https://user-images.githubusercontent.com/85752084/122598391-0077ff00-d02a-11eb-9733-af53b5f2454c.PNG)
#### Most of the companies grew in 2017, aside from TERP. DAQO (DQ) grew the most in 2017 with a growth of 199.4%. This might lead Steve's parents to want to invest in this company, but the 2018 data shows that this is not the wisest decision. As aforementioned, only ENPH and RUN continued to grow into 2018. DQ took a particularly hard hit with a decrease of 62.6%.
---
#### The original code was refactored in this project. The purpose of refactoring is to make the code work faster and continue to work with the addition of new code that might have messed up the original code. This code was refactored by adding arrays to store all the values in the loops, rather than single variables, which would then be called on after the calculation loop had finished to print all values almost at once.
![Pre Refraction 2017](https://user-images.githubusercontent.com/85752084/122598394-0241c280-d02a-11eb-9a33-4a5c64984b5d.PNG) ![VBA Challenge 2017](https://user-images.githubusercontent.com/85752084/122598405-053cb300-d02a-11eb-9a0d-a194f765735d.PNG)

#### Before refactoring the code, to run the 2017 data, it took 0.742 seconds. After refactoring the code, the 2017 data took 0.594 seconds. This is a decrease of 24.9%. In this code, the data table is relatively small, so 24.9% does not affect the time much, but in a data set that is much larger, this could take the time it takes to run the code from a couple of seconds to less than a second, which is substantial.
![Pre Refraction 2018](https://user-images.githubusercontent.com/85752084/122598401-040b8600-d02a-11eb-8c66-8f80d1c0a673.PNG)
![VBA Challenge 2018](https://user-images.githubusercontent.com/85752084/122598411-066de000-d02a-11eb-9587-329d5b3710a6.PNG)
#### Before refactoring the code, to run the 2018 code, it took 0.875 seconds. After refactoring the code, the 2018 data took 0.586 seconds. This is a decrease of 49.3%. This difference is enough to make a code that runs for 2 seconds take 1 second. A 50% refactoring rate is very substantial when looking at a very large data set.
#### In addition, during the refactoring process, an error handler was added. If an incorrect year was inputted, the following code would run at the beginning:
```
yearValue = InputBox("What year would you like to run the analysis on?")
    On Error GoTo errorProt
```
#### This would take the code to the following lines:
```
'if this all worked, skip error message
    GoTo skipMessage
    
'if there is an invalid year
errorProt:
    MsgBox "This is an invalid year. Please try again."
    
'normal exit
skipMessage:
End Sub
```
#### This sends the executor to the error message if something went wrong, but skips it if everything went as expected. Another line could be added that looks like this:
```
On Error GoTo 0
```
#### This would show the exact error message, which would be very useful while developing the code, but since the code already works, the only error Steve will get is if he inputs an invalid year, so this can be left out. This is important because it allows the code an exit and message to determine what the problem was while writing the code. 
## Summary
#### There are several advantages to refactoring code. The code can run faster, which is important in files with very large data sets. Also, the code can run more smoothly with an addition of new code. Some code works but is not pleasurable to look at and especially edit. Refactoring allows code to be read and added to, which is important in the coding community. There are also disadvantages. The main one is that it actually has to be done. If the code works, it's not motivating to try to make it better. Another is the code can stop working after changes have been made, which can be frustrating. This makes it much more desirable to leave the code how it started, but it is much more difficult in the long run, as it might not be able to accept new code, it is slower, and it takes much more RAM.
#### All of these pros and cons were observed in refactoring the code for this project. The code for this project did run faster after refactoring, which is a major pro. Also, when the error handler was added, it was easy to input because the code was readable, so I knew where to put the exit lines. There were also disadvantages. The code did not run for about 2 hours because there were some errors, which would not have happened if I kept the original code, as it was already working. Because of this, refactoring the code took about 3 hours, which I would have saved if I had kept the original code. Refactoring might be an unpleasant part of coding, but it allows code to run better, which is ultimately the goal.
For Steve, it can be logically concluded that DAQO is not the company to invest in. Instead, he should put his money in ENPH first and RUN second. These are both acceptable stocks, but ENPH has done better over the two years, so it should hold priority. Because of the growth in 2017 for DAQO, $10,000 would be worth $11,197.56 after 2018. ENPH, on the other hand, would be worth $41,746.05. RUN would be worth $19,412.00. Clearly, ENPH is the best choice, but RUN is also doing well. It should be noted that Steve has not lost money yet. In fact, he has gained quite a bit, but he should switch stocks now.
