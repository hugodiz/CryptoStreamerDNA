# CryptoStreamerDNA
Near-Time-Data crypto-currency streaming functions for Excel, powered by .NET via ExcelDNA

## Spec Summary:

**CryptoStream:** 
(core function / handle for the functionality)
When called from an Excel cell as =CryptoStream([Symbol], [Metric]) (eg. =CryptoStream("BNBBTC", "last_price")), it creates an approximately real-time data feed which will automatically refresh periodically without needing to explicitly recalculate Excel. 

The refresh rate is dynamically adjusted so as to be as high as it reasonably can without risking exceeding the API policies on Http request limits. Such limits are defined in terms of a maximum allowed 'request weight' limit per unit time (consecutively resetting after the unitary interval has passed). The RTD server which powers the Excel RTD feeds has a number of safety measures to ensure the API policies are never violated regardless of any end-user's Internet connection speed. These measures include:

- Automatic runtime reading and interpretation of API policy limits when the Excel session is first asked to connect
- A negative feedback loop which imposes an artificial dynamic delay between requests, in order to steer the expected request weight usage to be around 75% of the maximum allowed
- A damping mechanism which progressively waits more between calls if despite the negative feedbackloop, weight usage is becoming close to the limit
- A 'last resource' automatic cut-off -> cooling timer -> restart routine if, despite the measures above, the maximum allowed weight was (nearly) reached



**UNPACK:** because LSDLOOKUP can optionally give back a match's coordinates in the lookup_array instead of the text itself, and because lookup_array is allowed to be 2D, we need a way to represent arrays in single cells (in this case containing a tuple [row index, column index]) - the convention we'll use is JSON. This function takes a JSON string representation of any 1D or 2D array(ie. [A(1), A(2), ..] or [[A(1,1), A(1,2), ..], [A(2,1), A(2,2), ..],..]) and produces an actual dynamic array from it (pieces of which can then be taken using the Excel built-in INDEX function). Note that 1D arrays are single rows (not columns) by convention.

**TEXTSPLIT:** the inverse of the built-in Excel function TEXTJOIN. TEXTSPLIT takes a single (scalar input) string and returns a row containing each piece of the string, resulting from splitting the string according to a delimiter.

**RESUB:** takes an input array (2D allowed) and returns a similarly-sized array, where each entry has undergone a regular expression transformation (similar to .NET Regex.Replace). A regular expression pattern P is specified, as well as a replacement string R. All occurrences of P in each input are replaced by R. Although R must be a literal string, it may include the usual $**G** methodology (for re-using pieces of the pattern), where **G** is an integer number representing the **G**th captured group, if specified in P.

