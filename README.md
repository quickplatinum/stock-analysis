# Stock-Analysis Challenge
University of Toronto SCS Data Analytics Boot Camp Module 2 Challenge: VBA of Wall Street
# Stock-Analysis with VBA

## Overview of Purpose Project
Performing analysis on histrocial stock data to uncover trends, this includes creating a VBA Macro that will help our friend Steve analyze stocks performance for a given year. Initially he was impressed by the ability to have VBA analyze a large data set for this his parent's stock "DQ". But he wanted us to go further by creating a Macro that will analyze multiple stocks for many given years and that will be easily preformed by clicking single button. Included in the VBA Macro is also the formatting of the results table and conditonal formatting that shows how the stock is has perfomred after a year of trading. 

### Purpose
The purpose of the project is automation and code refactoring, the initial code that was written was great but had limitations if more data was added. THe original code was also not as fast, in this written report a comparison between the refactored code performance and the original code peroformance will be presented. 

## Analysis and Results

### 2017 vs 2018 Sock Performance
![2017 Stock Performance](https://user-images.githubusercontent.com/88692025/133011318-494236a0-dd1a-4404-9100-fc1dbdd4b89b.png)
![2018 Stock Performance](https://user-images.githubusercontent.com/88692025/133011679-f345b85e-e60a-4011-9321-b4f1b0247099.png)
In 2017, the highest performing stock was "DQ" at 199.4% gain, with "SEDG" at 184.5% as a close second, in comparison both stocks did not perform as well in 2018. Both stocks lost value throughout that year. The general pattern is that stocks in 2017 performed better, there were more gains across the board. Furthermore, a higher total daily volume in 2017 compared to 2018. The above tables summarize the results in both years for comparsion. The tables were created through the refactored VBA Macro using 2017 and 2018 data and formatted within the same press of a button.

## Comparing Original Code Running Time vs Refactored Code Running Time
### The orginal code was considerably slower at analyzing the data for both years, and took well above 1 second to analyze the data. Screenshotted below is the running time for the 2017 and 2018 all sotck analyis.
![VBA_Challenge 2017 (Original) ](https://user-images.githubusercontent.com/88692025/133011973-5ab5d70d-710d-4925-9a96-a1b3e369b0be.png) ![VBA_Challenge 2018 (Original) ](https://user-images.githubusercontent.com/88692025/133012026-afaae617-6ba2-4e20-be07-fb7b1b9cfeb8.png)
### The refactored code was considerably faster at analyzing the data for both years, and took well below 0.5 seconds to analyze the data. Screenshotted below is the running time for the 2017 and 2018 all sotck analyis.
![VBA_Challenge 2017 (Refactored) ](https://user-images.githubusercontent.com/88692025/133012173-c66aa50d-c93e-40d0-8aa6-cbe40add7477.png) ![VBA_Challenge 2018 (Refactored) ](https://user-images.githubusercontent.com/88692025/133012178-ae83f884-1d70-44dd-b436-802e5eb51444.png)

# Summary
### Advantages of refactoring code:

* The code is easier for people to read and becomes user friendlier for future editors and workers on the code
* Doing so is a clear and obvious way to make the code more efficient and reducing run time by taking fewer steps
* Refactoring uses less memory

### Disadvantages of refactoring code:

* Time consuming: it takes a while to figure out pitfalls and repurpose code
* If it is already a large code, reformatting might cause more bugs and it will be rsiky when fixing a large code
* Takes longer to fix an existing code than to work on a new one, especially if the person had nothing to do with the original code.

How do these pros and cons apply to refactoring the original VBA script?
* The conclusion is that the refactored code has served it's purpose, the running time was cut down by almost 4x in both cases! this means that the original code was clunkier and the refactored code is more streamlined and effecient.
* For comparison's sake, attached are screenshots of the original and refactored code in VBA.
#### Original Code
![Original Code](https://user-images.githubusercontent.com/88692025/133012440-44862072-b419-4030-8855-7da9e72dc3dd.PNG)
#### Refactored Code
![Refactored Code](https://user-images.githubusercontent.com/88692025/133012449-2faa269d-7cf0-4942-bc9b-3b423bbe39cf.PNG)
* The difference within the loop creation and range limit creates a more streamlined code when running the refactrored version. Overall our mission to help Steve is a success and the second code is even better for his purpose of adding more companies. The time spent and reformat proccess outweigh the risks of reformatting this specific code.
