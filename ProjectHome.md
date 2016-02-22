A small piece of code written in C# that uses Microsoft Powerpoint 2010 automation libraries to convert Powerpoint and Open Document Presentations to video from the command line.
I've kept the code simple so it is easy to understand. [Do What The Fuck You Want](http://en.wikipedia.org/wiki/WTFPL) with it :)


## Formats Supported ##

  1. PPT
  1. PPTX
  1. ODP

Embedded videos and fonts also work


## Requirements ##

  1. Microsoft Windows
    * Server 2008R2 and Windows 7 have the best media support for conversions
  1. Microsoft Powerpoint 2010 with developer tools installed
  1. Dot NET is also required


## Usage ##

```
PPTVideo.exe <infile> <outfile> [-d]
```

Use the '-d' switch to **delete the infile**, the outfile format is WMV and the in/out files can be relational paths to the working directory or fully qualified.

The resulting file is a perfect conversion and works much better than commercial tools that attempt to do the same thing (Probably because we are using Microsoft's conversion tool, this just provides command line access).
I then use FFMPEG to convert the video into any other required format, in this case WEBM for use on the web.

Also this is free :)