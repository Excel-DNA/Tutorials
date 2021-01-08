# What is RTD?

Excel's 'Real-Time Data' (RTD) feature lets us push data updates into an Excel sheet.
A typical example for RTD is to provide a real-time stock ticker feed.

You might already have noticed the `=RTD(...)` function on the list of worksheet functions, or seen me describe some Excel-DNA feature as 'RTD-based'. 
In this tutorial I'll explain what RTD is, how works under the covers and give an example of adding it to an Excel-DNA add-in.

I can think of some reasons why the RTD feature of Excel is not so well known:
* There are no built-in RTD data sources that ship with Office. There are only available from as part of third-party add-ins.
* RTD Servers cannot be created in VBA, so they need another environment like C++ or .NET to create.
* The COM-based nature of RTD means there are a few things to learn and take care with when making RTD Servers.

Despite the challenges in getting to know RTD, it is a very powerful feature of Excel that is closely integrated into the Excel calculation engine. Hence, it provides a foundation on which various high-level features can be built. But let me not run ahead of myself.

## What is an RTD Server?

An RTD data source defined by code in a COM(\*) library that implements an 'RTD Server'. An RTD Server supports the interaction between Excel and the data source by implementing the `IRtdServer` interface. An RTD Server will then expose its data as one or more 'Topics', each of which is defined by an array of strings passed to the RTD Server when connecting to a new topic, and in return provides a stream of values back to Excel.

> (*) What is COM?
>
> This is separate essay question... In brief, the Component Object Model (COM) is a standard that describes how software components can interact. The Excel COM object model is the set of interfaces that VBA uses to interact with Excel - this includes objects like `Application` and `Workbook`. Excel-DNA add-ins can also use the COM object model to interact with Excel. COM libraries are .dll libraries that work according to the COM standard. So in the context of RTD, it means that an RTD Server must follow these standard rules, so that Excel knows how to interact with it.

Being a COM class, an RTD Server is identified by its COM 'ProgId'. These strings normally have a dotted form like 'MyCompany.RtdServer' or 'MyCompany.DataLink'. Behind the scenes there is also a Guid (a 'globally unique identifier') called the COM 'ClsId' for the RTD server, which normally looks like a long hexadecimal number 'B73B68BD-9DD0-4E9D-82A1-E9B2798AF8E5'.

The combination of the `IRtdServer` interface and the `IRTDUpdateEvent` callback interface form the COM-based specification for how RTD Servers will iteract with Excel.

The `IRtdServer` interface has these members:
* `ServerStart` - create a new connection to the server (before any topics are connected)
* `ServerTerminate` - end the connection to the server (after all topics are disconnected)
* `ConnectData`- create a new topic according to the given topic strings
* `DisconnectData` - notify the server that a topic is no longer connected to Excel
* `RefreshData` - called to fetch updates for all topics
* `Heartbeat` - check that the server is still running.

The helper interface to notify Excel of any updates is called `IRTDUpdateEvent`, and is passed to the server in the `ServerStart` call.
The server then notifies Excel that new data is ready, with a call to `IRTDUpdateEvent.UpdateNotify`.

## How does Excel interact with an RTD Server?

We can now trace the interaction sequence between Excel and an RTD Server.

A basic call sequence might look like this
![RTD Call Sequence](https://user-images.githubusercontent.com/414659/104023185-f2e06280-51c9-11eb-9873-ab66cd07dae5.png)

## `ExcelRtdServer` helper class in Excel-DNA

Excel-DNA contains a base class that 


## Building on RTD


Some of the features of RTD:
* High throughput to support many data items, and high update rates
* High performance updates that do not interfere with the user's interaction with Excel
* 


## References

