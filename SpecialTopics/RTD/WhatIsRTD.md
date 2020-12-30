# What is RTD?

Excel's 'Real-Time Data' (RTD) feature lets us push data updates into an Excel sheet.
A typical example for RTD is to provide a real-time stock ticker feed.

You might already have noticed the `=RTD(...)` function on the list of worksheet functions, or seen me describe some Excel-DNA feature as 'RTD-based'. 
In this tutorial I'll explain what RTD is, how works under the covers and give an example of adding it to an Excel-DNA add-in.

## What is an RTD Server?

An RTD data source is a program defined in a COM library by implementing an 'RTD Server'. An RTD Server defines the interaction between Excel and the data source by implementing the `IRtdServer` interface. An RTD Server will then expose its data as one or more 'Topics', each of which is defined by an array of strings passed to the RTD Server when connecting to a new topic, and provides a stream of values back to Excel.

Being a COM class, an RTD Server is identified by its COM 'ProgId'.

The `IRtdServer` has these members:
* `ServerStart` - create a new connection to the server (before any topics are connected)
* `ServerTerminate` - end the connection to the server (after all topics are disconnected)
* `ConnectData`- create a new topic according to the given topic strings
* `DisconnectData` - notify the server that a topic is no longer connected to Excel
* `RefreshData` - called to fetch updates for all topics
* `Heartbeat` - check that the server is still running.

The helper interface to notify Excel of any updates is called `IRTDUpdateEvent`, and is passed to the server in the `ServerStart` call.
The server then notifies Excel that new data is ready, with a call to `IRTDUpdateEvent.UpdateNotify`.



Some of the features of RTD:
* High throughput to support many data items, and high update rates
* High performance updates that do not interfere with the user's interaction with Excel
* 
