# agamimGEOCODER

Eden: 
* This application will allow a user to upload an excel spreadsheet, select a column holding addresses, and recieve a new spreadsheet containing the geocoded addresses in a new column.
* I am using the Google Geocoding API
* I need your help with an async function
* I have a for loop initiating an async `geocoder.geocode(address)` function that has a callback.
* In the callback I am pushing the `res.latitude` to the array I created.
* After the for loop, I am trying to `console.log(array)` but it is empty (because the geocode function has not finished running)

#### How do I set up some Promise Async/Await kind of thing that will wait until all geocodes and array.push's in the for loop have finished?
Here is the current for loop:  
![](https://i.imgur.com/Ggk5CUc.png)

If you want to check out code, you can fork this repo, `npm i`, `npm run dev`, enter your Google API key (I'll send you mine to save time), and start reading from line 56  in `server.js`  
I commented most of the code too... (don't remove or add an excel sheet, I hardcoded it to the current sheet for now)
