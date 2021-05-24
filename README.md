# CryptoStreamerDNA
Near-Time-Data crypto-currency streaming functions for Excel, powered by .NET via ExcelDNA

## Spec Summary:

**=CryptoStream(Symbol, Metric):** core function which implements the CryptoStreamerDNA functionality

When called from an Excel cell as =CryptoStream([Symbol], [Metric]) (eg. =CryptoStream("BNBBTC", "last_price")), it creates an approximately real-time data feed which will automatically refresh periodically without needing to explicitly recalculate Excel. The streams are provided by the popular **Binance** API (using the public endpoints for Market Data).

The refresh rate is dynamically adjusted so as to be as high as it reasonably can without risking exceeding the API policies on Http request limits. Such limits are defined in terms of a maximum allowed 'request weight' limit per unit time (consecutively resetting after the unitary interval has passed). The RTD server which powers the Excel RTD feeds has a number of safety measures to ensure the API policies are never violated regardless of any end-user's Internet connection speed. These measures include:

- Automatic runtime reading and interpretation of API policy limits when the Excel session is first asked to connect
- A negative feedback loop which imposes an artificial dynamic delay between requests, in order to steer the expected request weight usage to be around 75% of the maximum allowed
- A damping mechanism which progressively waits more between calls if despite the negative feedbackloop, weight usage is becoming close to the limit
- A 'last resource' automatic cut-off -> cooling timer -> restart routine if, despite the measures above, the maximum allowed weight was (nearly) reached. 

Note: the cooling routine **should not happen frequently**, although if it does happen recurrently, it means the code isn't being robust / restrictive enough in the damping mechanism, or it's being too ambitious in aiming for 75% of the maximum allowed weight. Right now this would require a tweak to the source code, which although relatively simple, is not ideal - ideally, I will include a Ribbon input to specify a desired preiodicity between requests, which although will not be allowed to be lower than what the RTD server deems sustainable, may be arbitrarily high (hence allowing arbitrarily slower pace for the functionality).

**Streamer:** Excel Ribbon group which serves as the User Interface for CryptoStreamerDNA

This Excel ribbon group contains 2 sections: 

- a Doge button which controls On/Off switching of the streamer mechanism (this switch needs to be turned on **in addition** to CryptoStream() Excel feeds existing in the worksheet, in order for values to actually be streamed). The button also shows the overall status of the CryptoStreamerDNA, through various Doge status displays.

- A telemetry box which, whenever things are actually being streamed, will show the key stats regading the data in-flow taking place; this includes request weight limit information obtained from the API, as well as currently used weight during this time interval, and the average time between requests (which is dinamically adjusted in order to always remain sustainable)

**=CryptoSymbols():** This helper Excel function takes no arguments, and when called in Excel (preferably in a free column), will produce a dynamic column array containing all valid cryptocurrency symbols that can be selected from in the CryptoStream function.

**=CryptoMetrics():** This helper Excel function also takes no arguments, and when called in Excel (preferably in a free column), will also produce a dynamic column array, this time containing all valid cryptocurrency metrics that can be selected from (for any given symbol) in the CryptoStream function.

## Introduction
This hopefully useful Excel functionality for (approximately) Real-Time-Data streaming are fully written in VB.NET and plugged-in to Excel as an **xll** add-in. The functionality consists of a core Excel function (CryptoStream), 2 helper / documentation functions (CryptoSymbols, CryptoMetrics), and an Excel Ribbon group which serves as the main UI ("Streamer").

These functions and the ribbon rely entirely on ExcelDNA (by Govert van Drimmelen) in order for them to be visible from within Excel.

Internally, the JSON deserialization of API responses relies on the extremely popular .NET library [NewtonSoft.Json](https://www.newtonsoft.com/json) (by James Newton-King).

I've commented the code extensively because I hope some bits can serve as a 'sample contribution' of how to implement certain things with ExcelDNA. I am hoping this can be of use for Excel power users, VBA and Excel developers who may have varying levels of familiarity with .NET. 

The ability to create .NET-powered functions such as these and then exposing those functions to Excel worksheets is traditionally the type of thing that is made dramatically easier, more tracktable and more seamless using the excellent ExcelDNA open-source project. However, in many situations it is also true that the creation of rich / reactive UI elements is also ideally suited for ExcelDNA. I believe that is the case here, where the Ribbon section of the Excel UI is basically used to implement an interactive display / control panel for the CryptoStream functionality.

This code is open-source (MIT license) and these functions, whilst (hopefully) useful on their own, are again also meant as a small contribution to showcase the ExcelDNA toolset to experienced Excel users and programmers who may at times either feel limited by VBA, or tend to build extremely complex programs in VBA which would be better suited for .NET.

This project is the second of a series with currently 2 projects: TextUtilsDNA and this one (CryptoStreamerDNA). They are functionally unrelated, so from the perspective of usage they are independent from each other. However, insofar as these are also meant to serve as a learning tool for ExcelDNA, from the point of view of becoming familiar with the .NET / Visual Studio / ExcelDNA ecossystem, TextUtilsDNA is the best project to start with (both in terms of actual code and GitHub documentation), and then this one is a good follow-up. You can find TextUtilsDNA in the following GitHub repo:
[https://github.com/hugodiz/TextUtilsDNA]

Ultimately, depending on your project's size, performance and interoperability needs, VB.NET might be a much better choice than VBA. It's certainly a great stepping stone into .NET, for those interested. As of 2021, ExcelDNA is one of the best ways to bring the power of .NET (C#/VB.NET/F#) to Excel. If this is new to you please visit:  
[https://docs.excel-dna.net/what-and-why-an-introduction-to-net-and-excel-dna](https://docs.excel-dna.net/what-and-why-an-introduction-to-net-and-excel-dna)    
as a starting point.

These functions are ideally meant to be used with Excel 365, because the 2 helper functions levarage the power of dynamic arrays.  
However, the core function (CryptoStream) should give no problems in most Excel versions. Basically, you should be fine whenever one of these functions would return a scalar anyway (which CryptoStream does, anyway).

Without dynamic arrays in your Excel version, I believe that, for CryptoSymbols() and CryptoMetrics(), you will need to pre-select a range of the right size, then use the function normally, but trigger it with ctrl + shift + Enter instead of just Enter. Otherwise, Excel might just show you the upper-left corner of the result instead of the whole (array) result.  

