# CryptoStreamerDNA
Near-Time-Data crypto-currency streaming functions for Excel, powered by .NET via ExcelDNA

## Spec Summary:

**=CryptoStream(Symbol, Metric):** -> core function which implements the CryptoStreamerDNA functionality

When called from an Excel cell as =CryptoStream([Symbol], [Metric]) (eg. =CryptoStream("BNBBTC", "last_price")), it creates an approximately real-time data feed which will automatically refresh periodically without needing to explicitly recalculate Excel. 

The refresh rate is dynamically adjusted so as to be as high as it reasonably can without risking exceeding the API policies on Http request limits. Such limits are defined in terms of a maximum allowed 'request weight' limit per unit time (consecutively resetting after the unitary interval has passed). The RTD server which powers the Excel RTD feeds has a number of safety measures to ensure the API policies are never violated regardless of any end-user's Internet connection speed. These measures include:

- Automatic runtime reading and interpretation of API policy limits when the Excel session is first asked to connect
- A negative feedback loop which imposes an artificial dynamic delay between requests, in order to steer the expected request weight usage to be around 75% of the maximum allowed
- A damping mechanism which progressively waits more between calls if despite the negative feedbackloop, weight usage is becoming close to the limit
- A 'last resource' automatic cut-off -> cooling timer -> restart routine if, despite the measures above, the maximum allowed weight was (nearly) reached

**Streamer** -> Excel Ribbon group which serves as the User Interface for CryptoStreamerDNA

This Excel ribbon group contains 2 sections: 

- a Doge button which controls On/Off switching of the streamer mechanism (this switch needs to be turned on **in addition** to CryptoStream() Excel feeds existing in the worksheet, in order for values to actually be streamed). The button also shows the overall status of the CryptoStreamerDNA, through various Doge status displays.

- A telemetry box which, whenever things are actually being streamed, will show the key stats regading the data in-flow taking place; this includes request weight limit information obtained from the API, as well as currently used weight during this time interval, and the average time between requests (which is dinamically adjusted in order to always remain sustainable)

**UNPACK:** 


**TEXTSPLIT:** the inverse of the built-in Excel function TEXTJOIN. TEXTSPLIT takes a single (scalar input) string and returns a row containing each piece of the string, resulting from splitting the string according to a delimiter.

**RESUB:** takes an input array (2D allowed) and returns a similarly-sized array, where each entry has undergone a regular expression transformation (similar to .NET Regex.Replace). A regular expression pattern P is specified, as well as a replacement string R. All occurrences of P in each input are replaced by R. Although R must be a literal string, it may include the usual $**G** methodology (for re-using pieces of the pattern), where **G** is an integer number representing the **G**th captured group, if specified in P.

