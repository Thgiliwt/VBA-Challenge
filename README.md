# VBA-Challenge
This is the 2nd homework from my coding bootcamp course
Within this homework, I am required to use VBA scripting to analyze stock market data. Below are the concepts only within the script and also screenshots from the excuated Excel.xlsm file showing the results of year 2014, 2015 and 2016. For detailed coding contents please check at 'The VBA of Wall Street.vbs' within this repository.

Before I started, I firstly had 'option explicit' at the top so I could know if there is any declartion went wrong when running the scripts. I have also create a module for my script.

Step 1 - Understanding about the tasks
The Excel workbook has 3 worksheets which are named '2014, 2015, 2016' respectively. They have the same layouts but differences are the number of rows in the ticker column. I will then generate a summary table where to show each ticker of stock name in one column, the yearly change and percent change as well as the total stock volume in another 3. For the bonus question, I will use some Excel functions to find out the maximum increased, decreased and total volume from all of the stocks in a second table.

Step 2 - Building up scripts
With reference to the uploaded script file, I will explain this within 4 parts
  
  Part 1 - setting texts in cells
  This is purely about setting up the summary table layouts etc. I used the 'split' function and a 'for' loop to make it easier. Then manually typed the reset texts in cell O2, O3 and O4.
  
  Part 2 - Determining values for Ticker(stock name), Yearly Change, Percent Change and Total Stock Volume & Formatting
  This is the main task of this homework so it involves lots of codes. As the first summary table, where we list up all the stocks and we should have their corresponding values under 'Yearly Change', 'Percent Change' and 'Total Stock Volume' columns, I decided to have a 'for' loop covers all of these, together with 'if' conditional statements.
  To locate a difference within a list of data, we could use the IF statement, if the value of current row is not equal to the value in the next row, or, if the value of previous row is not equal to the value in the current row, then there is a difference. To make it easier, I picked the first statement as it could spot it right to my current selected row. We would like to have this happened over and over again all the way to the very last row in the datasheet so we could have all of the values returned. For the 'Yearly Change' value, it will be the close value in the end of a stock, substracts the open value at the very start of a typical year. The 'Percent Change' equals to 'Yearly Change/openvalue' and the 'Total Stock Volume' is just the sum up in column G.
  We now have the main concept, so we need to setup some variables and range for VBA to know what we expect from it. By viewing the data, we will make Ticker as string, openvalues and closevalues as double (there are decimals) and stock volume as Variant (it is out of the Long range by checking). We will also have new variables as lastrow, rownumbers and a as long (it is over the integer limit).
  Within the loop, we need to have some important codes to make the loop loops. First of all we need to make the rownumber to rolling up with a step of 1 so our first summary table could be well organized. We also need to have the change of openvalue based on our a value so whenever there is a difference spotted then the 'Yearly Change' outcome will be the closevalue of the current row (row a) minors the previous openvalue. To make the openvalue rolling as well, we need to have an initial value of it, luckily, all openvalue is stable in cell C2 for all 3 worksheets. Along the way, we notice that the openvalue of the next Ticker, could be expressed as cells(a+1,"C"). Hence the basic idea is to ask the VBA to stored the next openvalue after the current calculation is done, so we could have the next closevalue, which is the cells(a,"F") play with it.
  The total volume calculation is a little bit tricky. We could apply with the same method as discussed above, but only in that way, the VBA only records the volume at the row where it detects the differences in column A. What we want is the total sum of a typical Ticker, which means that we would have the sum of the volume in the rows where there is no difference in column A. In that case, we will have the 'Else' statement comprised within the outter IF, to ask the VBA add up the volume for the same Ticker. This is not enough as now, the VBA will just adding everything up. We will have to ask it to only have the sum of the same Ticker. To do this, we need to ask it to return a value of '0' after the difference is spotted and added. Please note that this should happen right after the difference is spotted, so we have it wihin the outter IF statement.
  There are few things we need to pay attension to. First of all, we need to have formatting for the Yearly Change values. Postive with green and negative with red. To make it more obvious, I also set blue as a value of '0'. This is done by a inner IF. For the Percent Change, we will have another inner IF, as when the openvalue is 0, there will be an error where openvalue as the divident cannot be 0. and we need to have the percentage format of cells for this column.
  
  Part 3 - Determining greatest % change for both postive and negative, greatest total stock volume and their corresponding name
  This is the bouns part, and the best way to solve it is to set up an array so we could apply some basic Excel function to it rightaway. Things to notice, that in an array, the first item index number is 0. So it is very important that you have the +1 when you need to mix with the actual row number in a particualr column.
  
  Part 4 - Layouts
  This is again very simple. Just to make the cell looks easy to read.

By far, we have done the coding on one worksheet. Now we need to tell VBA to run the same script on every single worksheet to meet the requirement. In this section, I used a 'CALL' function where it could call the subroutines that you have created so you do not need to write it again. And it has tidy looking within the script. It is suggested to have the screenupdating temporily disabled just to make the system runs quicker. We need to use 'For Each' loop for this because we just repeating the application 3 times, or say, to 3 worksheets which could be solved with the same script.

Please see below screenshots of results:
<img width="1145" alt="2014" src="https://user-images.githubusercontent.com/83489530/118499759-93dbbd00-b76a-11eb-8140-1637fc152b85.PNG">
<img width="1147" alt="2015" src="https://user-images.githubusercontent.com/83489530/118499779-963e1700-b76a-11eb-8ea2-01cffb91d2f0.PNG">
<img width="1160" alt="2016" src="https://user-images.githubusercontent.com/83489530/118499794-98a07100-b76a-11eb-8653-75d686c0fcaf.PNG">


That is everything for this homework. I was amazed that I got the result without any error the first time I ran it on the final excel file. A huge encourage to myself. However, there are things I have been struggling with, such as:
  1. On the test runs, I firstly got the issue there it was not providing correct results in sheet P. It turned out that there is 0 openvalue invovled so I added an inner IF statement to toggle out the situation where the open value is 0;
  2. The excel file corrupted few times, it said that could be fixed if you trusted the source but it just eliminate all the coding. So I have learnt to copy my script into a notepad, regularly. (It not help if just makeing duplicated saves on xlsm, I tired to figure out the reason but it is just a deadend);
  3. There are times where I could not work out the correct codes for VBA to run, or to have the result that I would love to see. I have searched many resources and I also joined a forum where people share their knowledge on programming languages. It is very helpful because I feel the learning is not like what I have covered before, it is more about the trigger of your thoughts. It is all about sudden enlightment, and it will finally show-up when you have put enough practices.
