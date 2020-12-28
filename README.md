# stock-analysis

Overview of project:

  The purpose of this project was to create a VBA script that could loop through daily stock data (for multiple stocks) and extract data (ticker, total daily volume, and yearly return). The module showed us one way to do it using nested for loops. The job of the student was to recreate this functionality wihtout using nested (2 intertwined) for loops. In writing this code, for loops, if statements, VB formatting, arrays, and VB timer functions were used. In completing this project, students were taught how to use VBA efficiently.


Results:

  Much of the original code was left intact. The code that was left intact was the calculation and display of run time, table formatting, ticker array, and year input.
[Insert a picture of the entry form]

  The adjustment that I made to the code that was shown in the module is that I avoided using a nested for loop. In the module, the macro ran through the stock data once for each ticker (12 times total). In my refraction, I only looped through the data once. This improved run times from 1.8 seconds to 0.2 seconds per year.
  
original code:


improved code:



Summary:

  Refractoring code is a useful process, as long as the editor knows the entire scope of this project. I was able to take advantage of the ticker_index integer to avoid having to run a nested for loop. I was only able to do this because the data was sorted by date and ticker. Had the not been sorted, the new code would not have worked properly. If this project was expanded to query for the data set, the query must have an "order by ticker, date" statement at the end.
