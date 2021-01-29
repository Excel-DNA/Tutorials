 # VBA to .NET: Web Request Basics
 
 A common programming task is to retrieve some information from a remote system via an HTTP web request. This might be very easy or very difficult depending on the nature of the remote system, the type of data to be retrieved, the amount of processing required and the way the data will then be used. So this is a large topic.

This tutorial gives a high-level view of how to aproach this task, an example of how to get started with the simple case, and some pointers on where to look further.

## Introduction

I use the term 'web-based resource' for any information that can be accessed via the standard HTTP protocol. This includes:
* public web sites (reading the information would often be called 'web scraping'),
* web-based services that can be accessed with an HTTP-based API,
* internal services in a company, or even running on the same machine where it will be accessed.

(While I won't discuss it further here, other protocols like FTP might also be required to access web resources.)

There is a huge variety of web-based resources, and reasons for wanting to access these resources. For example:
* gather data from a web site into an Excel spreadsheet for further analysis or custom reporting,
* periodically read data and store in a different database for further use,
* perform some actions through a web-based interface, like uploading data or sending bulk email messages.

Like the range of resources, and the various reasons for accessing these, the ease or difficulty in interacting with web-based resources also varies a lot. Some web site are very easy to access from a script or program, while others may be extremely difficult to interact with even for a human using a browser, and more so for an automated system. Into this mix also comes various types of security involved, including authentication of a user, various ways that website try to prevent being accessed by programs ('prove you're human') and throttling to prevent large amounts of (or sometimes any) data being accessed.

This tutorial is not meant as a comprehensive guide to this wide range of cases, but rather has the following aims:
* give a high-level view of the web-request topic, helping to strucutre your thoughts and programs when taking on such a task,
* provide a basic example of how to go about the simple case,
* explain the varous ways in which things become complicated, and point to resources that might help.

## Bird's eye view

I find it useful to structure the web-request programming tasks into three aspects - **Fetch**, **Process** and **Use**.

#### Fetch 
The various calls and mechanism needed to interact with the remote system, including protocol settings, API keys, cookies or other security aspects.

#### Process 
What to do with the results returned from the fetch phase to extract the information we want, or convert into a format or data structure we want to use. For example finding particular parts in an html result page, or parsing a JSON result string into a friendly data structure.

#### Use 
This covers the program environment that will initiate the web request and receive the processed results for further use. For example this might be an Excel add-in which exposes user-defined function for accessing the web resource, or it might be a console application that is run periodically to donwload a data set and store into a local database.

I think of these aspects in a relationship like this:

:TODO:


While the **Fetch** and **Process** aspects offer quite different technical challenges and need different tools, they tend to work as a pair when developing access to a particular web site or resource. A particuler web resource would follow a pattern for access (for example security, call pattern) when fectching different bits of information. Similarly, the web site structure and processing needed to get to the right bits of data would often change together when systems are updated. So it makes sense to encapsulate or bundle them together into an `XyzClient` library that can easily be used in different. For example we might have a `YahooFinanceClient` or a `FREDClient`. Indeed for many web resources _there might already be a client library for .NET available_ from the service provider, or as a project on GitHub.

The **Use** aspect concerns the specific way we will run and use the client library and the results we get. I will not be focusing much on this aspect for the tutorial, but will pick a simple case for the example - press a button in Excel to download and put some data in a sheet. It's often useful to think of multiple ways in which the same client for a web resource can be used. Sometimes it might be interactively through a spreadsheet or other interface, while at other times the data will be fecthed and stored for later analysis, for example when a machine learning model will be trained against the data set. Of source the usage setting woud aslo drive the requirements for what information and format the client library should support.

Generally the programming tools needed in the **Use** aspects are different and independent of those needed in the web-related client library. Results of the web requested data might be used from console applications or scripts (which are easy to automate), from Excel add-ins where the are directed by end-users  or for storing in databases, which in turn might be on a local (like an Access, SQLite or SQL Server database), on a network or internet hosted service.

### Notes on program style

There are some questions of program structure and style that I think are relevant to the topic and examples, so I'll mention these briefly.

#### Using and exposing asynchronous calls

Since web request mostly cross machine and often large geographic boundaries, they tend to have high latency, so a program that is making such requests will spend a lot of time (in computing terms) just waiting for the next response.

Writing programs that are well structured but still interact well with such high-latency calls was historically quite difficult. But the excellent async/await programming model developed in .NET and now used in most programming platforms makes this kind of programming much easier. Async programming is another important topic to cover in the VBA to .NET context.

I recommend that the wrapper client library also expose async methods, as these can easily be consumed both in settings where the async aspect is useful, and in places where a synchronous interface is better suited. So I will also follow that pattern in the example here.

#### Interfaces and abstraction

One question that might arise in structuring such a project is whether to have an abstract interface that defines the abilities and usage for an `XyzClient`, maybe an `IWebClient` interface. This would allow various clients to be interchanged as components in various useage scenarios. 

If a project will need many different web-resource clients that will all be used in a similar way, it might well make sense to abstract the usage through a common base class or by implementing common interfaces.

I will not be adding this complexity in the examples here.

#### Automated testing

The web request client library often deals with external systems, which make automated testing more difficult. It can help to at least have a simple console program that uses the client library, as a way of using and testing the library witout the larger environment and other components involved. Conversely, for some projects it will help to have a mode or a mock version of the library that can be used to provide predictable results for automated test purposes, so that the mock version can be used instead of the real web requests in an automated test setting.


## The simplest thing that might work



### Fetch


### Process

### ***Use** from a console runner

### ***Use** from an Excel add-in


## Complications galore!

While the basic case is quite easy to make work, web request programming can become very difficult and lead to great frustration. In the section I give some glimpse of the ladder of complexity one might climb down, and some pointers to the tools that might be useful along the way.

### Fetching



### Processing


Single values
Tables
Download files
