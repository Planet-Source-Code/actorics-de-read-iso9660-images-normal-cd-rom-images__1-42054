<div align="center">

## Read ISO9660 images \(normal CD\-ROM images\)

<img src="PIC200311102834867.jpg">
</div>

### Description

This article will show you, how to read the main informations of a ISO9660 image.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-07-23 16:27:04
**By**             |[actorics\.de](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/actorics-de.md)
**Level**          |Intermediate
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Read\_ISO96152144112003\.zip](https://github.com/Planet-Source-Code/actorics-de-read-iso9660-images-normal-cd-rom-images__1-42054/archive/master.zip)





### Source Code

<p><h1>Hi! Welcome to this little tutorial!</h1></p>
<p>It will show you how to read the main informations of an ISO9660 Image file.<br>
First of all, what is an ISO9660 image file?<br>
ISO9660 is the standard specification for file systems on a CD-ROM.<br>
An ISO9660 Image is a 1:1 copy of a CD-ROM to a file on your hard drive.<br>
Mostly the extension is *.iso. <br>
But there are also other formats like: *.bin, *.img, ...<br>
I only want to show you how to get informations from a normal image (*.iso)<br>
because this is a very easy image format.</p>
<p>If you open an image with word or something like this, you will see....nothing.<br>
That's because the first 16 sectors of a CD-ROM are always empty.<br>
One sector of a CD-ROM is 2048 bytes big.<br>
Note: Sometimes VCDs have bigger sectors.<br>
In this project the sector size is 2048.<br>
Then comes the header: CD001<br>
And now it gets interesting.<br>
The Volume Descriptors are coming!!!!<br>
The next 32 bytes are the title of the CD-ROM.<br>
After the title comes the System Descriptor.<br>
It's also 32 bytes long.<br>
It describes the system that you need that the CD-ROM works.<br>
For the next part of the ISO you need to know what an Endian is.<br>
It's a data format for saving binary data. There are a lot of them in an image.<br>
The most used in the whole wide world is the big endian, I think.<br>
An 32 bit endian is always 4 bytes big.<br>
In the ISO9660 format you need to convert a lot of endians back.<br>
I added a demo project that reads more information than this article.<br>
Just look at the screenshot!

