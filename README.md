# stock_analysis

## Overview
The purpose of this stock analysis, initially, was to create an way for Steve to see sorted data in a way that would allow him to make good investment decisions without having to manually parse through thousands of rows. Steve then wanted something that could sort through significantly more data so we refactored the code to run more efficiently through larger data sets. 

## Results
Using our workbook Steve has a clear vision of his portfolio's performance during 2017 and 2018. The majority of his portfolio did better in 2017 than 2018 with the notable exceptions of ENPH and RUN. Each of which provided returns of over 80% in 2018. 
![2017](https://user-images.githubusercontent.com/25140609/189162320-8a85522d-ce0c-4ffa-98d8-8a1f5010a9c0.png)
![2018](https://user-images.githubusercontent.com/25140609/189162386-d192d94a-62d3-44d2-89fa-aaef66df991b.png)

After refactoring the code to run more efficiently we received significantly different speed test results in running our code: 

<img width="264" alt="VBA_Challenge_2018_Old" src="https://user-images.githubusercontent.com/25140609/189163780-e6fcad6b-8629-4264-a75c-a6eb31f38e13.png">
![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/25140609/189163974-567a86d0-b654-4deb-9ef2-38f55e28f047.png)

Our old code ran in .35 seconds, while our refactored code ran in .05 seconds. This may seem insignificant for our current use case considering that both results were produced in under a second, but that different will be much more noticeable with larger data sets. 

In order to streamline our code we made use of multiple arrays which allowed us to more efficiently sort and output our data using multiple arrays 

Our use of arrays made the data collected by the For loops faster to parse through for calculation and output:

  Dim tickerVolumes(11) As Long
  Dim tickerStartingPrices(11) As Single
  Dim tickerEndingPrices(11) As Single
    
    For i = 0 To 11
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
            
 ##Summary
 It is advantagous to have more efficient code for larger data sets as it reduces the time it takes to compute, but sometimes it is faster/simpler to write something a bit less efficient especially when the compute time different is negligable. In our case, both pieces of code produced results in a fraction of a second.
 
 The original code is a bit more flexible for someone wanting to reporpose it quickly and easily for a similar task, while the new code is much faster. 


   
