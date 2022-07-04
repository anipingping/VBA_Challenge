# VBA_Challenge
This repository is for the 2nd module challenge
## Overview of Project
The purpose of this project was to refactor a visual basic code that analyzed a dataset of stock market activity. This refactor procedure updated the VBA code, expanded the analysis of the dataset to include two years of activity for the entire stock market, created new arrays of data, and sped up the run time of the Excel macro that performs the analysis. 
## Results
In 2017, 11 of the 12 stock tickers captured in this dataset produced positive returns. The lowest positive return, for SunRun (RUN) electrical equipment company, was 5.5 percent. The highest rate of return produced in this data during 2017 was an eye-wateringly good 199.4 percent for Daqo New Energy (DQ). The only company to generate a negative return was TerraForm Power LLC (TERP), at -7.2 percent. These figures are provided in Table I below:
### Table I: 2017 Stock Analysis
![VBA_Challenge_2017](https://user-images.githubusercontent.com/106618404/177062633-b6eca231-1d34-4a2d-aedc-f0df92ee94f2.PNG)

In 2018, this same list of stocks performed much worse in comparision to 2017. Only two companies generated positive returns, Enphase Energy (ENPH) at 81.9 percent and SunRun (RUN) at 84 percent. Every other company generated negative returns. After a stellar 2017, Daqo New Energy (DQ) suffered a -62.6 return in 2018, eliminating a significant chunk of the gains made the previous year. These figures are provided in Table IIa and IIb below:
### Table IIa: Original 2018 Stock Analysis
![VBA_Challenge_2018_original script](https://user-images.githubusercontent.com/106618404/177062979-0e4489c7-982e-47c3-b90f-be047fb8ebe8.PNG)

### Table IIb: Refactored 2018 Stock Analysis
![VBA Challenge_2018](https://user-images.githubusercontent.com/106618404/177062862-1692297d-1e69-4e82-982a-9db3892ca2b2.PNG)

## Summary
### What are the advantages or disadvantages of refactoring code?
#### Advantages
One advantage of the refactored code in this analysis is the decreased run time of the script in Excel. The refactored code assesses a large set of data and provides results more quickly. Another advantage is that refactored code can be re-written and reassessed for more efficient coding language. The refactored code is more elegant.
#### Disadvantages
One disadvantage is that refactoring an existing code requires the coder to go into a chunk or block of existing code and make changes from within what already exists. It can be more difficult, and therefore more time consuming, to type within existing code and to make small changes to nouns (singular/plural), worksheet titles (2017 or 2018, etc.) Sometimes it's better/easier just to start over from scratch and avoid the hassle of trivial but critical syntax issues.

### How do these pros and cons apply to refactoring the original VBA script?
The refactored script called for the creation of 3 new arrays, tickerVolumes, tickerStartingPrices and tickerEndingPrices. Crucially, there is an 's' at the end of these array titles, which the original script did not have. Consequently, I had to keep an extremely sharp eye out for small changes to noun structure when finalizing the refactored script. This was very easy to miss initially and I had to very carefully read through what I had created (see screenshot below):

![code image capture](https://user-images.githubusercontent.com/106618404/177063367-61057021-02ab-4c42-9d82-6effbc99a25e.PNG)


