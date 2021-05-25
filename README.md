# CryptoStreamerDNA
Near-time-data crypto-currency streaming functions for Excel, powered by .NET via ExcelDNA

## Spec Summary:

**=CryptoStream(Symbol, Metric):** core function which implements the CryptoStreamerDNA functionality

When called from an Excel cell as =CryptoStream([Symbol], [Metric]) (eg. =CryptoStream("BNBBTC", "last_price")), it creates an approximately real-time data feed which will automatically refresh periodically without needing to explicitly recalculate Excel. The streams are provided by the popular **Binance** API (using the public endpoints for Market Data).

The refresh rate is dynamically adjusted so as to be as high as it reasonably can without risking exceeding the API policies on Http request limits. Such limits are defined in terms of a maximum allowed 'request weight' limit per unit time (consecutively resetting after the unitary interval has passed). The RTD server which powers the Excel RTD feeds has a number of safety measures to ensure the API policies are never violated regardless of any end-user's Internet connection speed. These measures include:

- Automatic runtime reading and interpretation of API policy limits when the Excel session is first asked to connect
- A negative feedback loop which imposes an artificial dynamic delay between requests, in order to steer the expected request weight usage to be around 75% of the maximum allowed
- A damping mechanism which progressively waits more between calls if despite the negative feedback loop, weight usage is becoming close to the limit
- A 'last resource' automatic cut-off -> cooling timer -> restart routine if, despite the measures above, the maximum allowed weight was (nearly) reached. 

Note: the cooling routine **should not happen frequently**, although if it does happen recurrently, it means the code isn't being robust / restrictive enough in the damping mechanism, or it's being too ambitious in aiming for 75% of the maximum allowed weight. Right now this would require a tweak to the source code, which although relatively simple, is not ideal - ideally, I will include a Ribbon input to specify a desired peeiodicity between requests, which although will not be allowed to be lower than what the RTD server deems sustainable, may be arbitrarily high - hence allowing arbitrarily slower paces for the functionality.

**Streamer:** Excel Ribbon group which serves as the User Interface for CryptoStreamerDNA

This Excel ribbon group contains 2 sections: 

- a Doge button which controls On/Off switching of the streamer mechanism (this switch needs to be turned on **in addition** to CryptoStream() Excel feeds existing in the worksheet, in order for values to actually be streamed). The button also shows the overall status of the CryptoStreamerDNA, through various Doge status displays.

- A telemetry box which, whenever things are actually being streamed, will show the key stats regarding the data transfers taking place; this includes request weight limit information obtained from the API, as well as currently-used weight during this time interval, and the average time between requests - which is dinamically adjusted in order to always remain sustainable.

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

At the end of the day, if you can, you really should be using Excel 365, it's worth it.

ExcelDNA will automatically produce 32 and 64-bit versions of the **xll** if you build the project in Visual Studio - you'll then use the appropriate one for your system (meaning, check the *bitness* of your Excel version). The above link to "ExcelDna - What and why" does a very good job of explaining what an **xll** is and how it relates to the other types of Excel add-ins available. From the end-user's point of view, an **xll** add-in is just something to be *added* to Excel, in a similar fashion to how you'd *add* a **xla** or **xlam** add-in.

## Getting Started
Documentation work in progress - but it's going to be pretty much the same process one would go through with any other **xll** add-in, in general, and any other ExcelDNA add-in, in particular.

**Binary releases of CryptoStreamerDNA are hosted on GitHub:** [https://github.com/hugodiz/CryptoStreamerDNA/releases](https://github.com/hugodiz/CryptoStreamerDNA/releases)   

In principle, downloading a copy of either the 32 or 64-bit CryptoStreamerDNA **xll** binary and having Excel ready go on your end, then adding the **xll** as an "Excel addin" in the Developer tab, is all one should need to do in order to get the functions up and running.

As mentioned in the **Introduction**, there is a sister project to this one ([https://github.com/hugodiz/TextUtilsDNA]). That project is ideally suited as a first serious ExcelDNA project to study for those learning .NET with a VBA background. That project will contain a step-by-step consolidated guide on how to build an ExcelDNA project from scratch using Visual Studio. Everything in that guide will apply equally to TextUtilsDNA and CryptoStreamerDNA (work in progress). However, my instructions / guide won't preclude the need (or at least the very strong recommendation) that you have a look at the series of excellent YouTube tutorials by Govert on getting started with coding .NET functions for Excel via ExcelDNA:   
[https://www.youtube.com/user/govertvd](https://www.youtube.com/user/govertvd)

## Examples
Documentation work in progress - in the meantime, the Spec Summary already details to a considerable degree how to use the functionality. This functionality is ideally suited for someone creating an Excel custom dashboard for viewing / managing the evolution of trading information of selected crypto symbols. In due time I'll upload an example of an Excel template which would sort of complement the functionality by effectively being a fully fledged UI for it. There's nothing special about such a 'template': it is simply any Excel workbook where you've built your own custom functions and macros, leveraging CryptoStream where appropriate.

## Support and participation
Any help or feedback is greatly appreciated, including ideas and coding efforts to fix, improve or expand this suite of functions, as well as any efforts of testing and probing, to make sure the functions are indeed 100% bug-free.

Please log bugs and feature suggestions on the GitHub 'Issues' page.   

Note in particular that, since this is all in an early stage, I expect we may find a few bugs in the CryptoStreamer. I expect that if one does find a bug, it will likely manifest in one of three likely ways:

1. The CryptoStreamer keeps entering cooling down mode (which indicates I haven't been thorough enough in the code in order to avoid the streamer reaching that last resource)
2. The CryptoStreamer UI enters a weird state with mixed elements from when it's supposed to be Off and supposed to be On (which should not be possible in principle).
3. Excel crashes (either closes without warning or displays a fatal error message) - this is not dangerous but effectively means there's an unhandled exception in the code somewhere, which is a bug.

## License
The CryptoStreamerDNA VB.NET functionality is published under the standard MIT license, with the associated Excel integration relying on ExcelDNA (Zlib License):   
[https://excel-dna.net](https://excel-dna.net)           
[https://github.com/Excel-DNA/ExcelDna](https://github.com/Excel-DNA/ExcelDna)

Hugo Diz

hugodiz@gmail.com

24 May 2021
