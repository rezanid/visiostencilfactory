# Visio Stencil Factory

Visio is a versatile tool to draw diagrams. The concept of *stencils* and shape metadata makes it very extendable. Whenever you need new shapes to integrate with your diagrams and you have them in SVG format it is very easy to just drag and drop those shapes to your diagram. Visio will just convert those shapes to its internal format on-the-fly and your are ready to go.

## The problem with adhoc SVG imports
If you have used Visio's built-in stenciles in your diagrams, you have probably noticed that the shapes that you drag-and-drop from these stencils are smarter than just a bunch of vectors. For example when you type a label for those shapes, the label is beautifully placed under the shape and you can even fine-tune the position of the label relative to the shape. Or when you connect these shapes to other shapes, the connecting lines are connecting to certain points around the bounding box of the shape that makes your diagram more readable. There is nothing stopping you from creating your own custom stencil and adding all the intelligence you want, but if you have many SVG files, doing it manually and one by one is not going to be easy. 

## Genesis
Recently I was designing our BCP and HADR solution for a Sitecore web application on Azure. One of my favorite drawing apps is Draw.IO and it has many Azure shapes out-of-the-box, but these shapes are very limited and they are mono-colored. I thought maybe I give Visio a go. To my surprise Visio's desktop edition does not have any Azure shapes. Though Microsoft has published a collection of shapes in SVG format, using them means the typical hassle of dealing with SVG files I mentioned above. What's even more of a pitty is that the icons you get in the collection is not based on the new theme of Azure.
Being the nerd I am, I decided to do a little more research and perhaps find a more complete and recent collection. I found out that Microsoft has made public all the icons that come as part of new Azure.com design system (https://azure.microsoft.com/en-gb/patterns/styles/glyphs-icons/). Doing a little more research lead me to this amazing website (http://code.benco.io/icon-collection/) that brings all the official and recent Azure icons together. The next step is to find a way to make stencils that look and act like the ones that come built-in with Visio and nothing less.
There are two main solutions for me.
### Solution 1
Visio like any other Microsoft Office application comes with a comprehensive well-documented automation API that is based on COM and can be called from any language (e.g. C#, PowerShell or C++). In PowerShell world it means after running the following command I can easily create and interact with Visio objects.
```
using namespace Microsoft.Office.Interop.Visio
```
For example to run Visio I can call `New-Object -ComObject Visio.Application`. The object model is very similar to the UI of Visio, but sometimes the names are different. This is understandable as the names in the UI are simpler, but the object model tries to be more precisely describe what it does. I had some experience with Office object model long time ago, but never in Powershell.
The good thing about this solution is that Visio will take care of converting SVG format if you know the right ropes, but the downside of this solution is that you need to have Visio installed on the machine you are developing. This might not be a deal breaker, since you need the stencils for Visio after all.

### Solution 2
Microsoft Office is using an open standard format to store all its files. This format is called Open Document Format (ODF). You can read about it [here](http://opendocumentformat.org/developers/). Basically every office file is a compressed package (ZIP) of XML files that are structured in a specific way. There is even a namespace in .NET that abstracts some of the complexity, but it is specific to any Office application. In fact you can use it for your own applications if you need.
The good thing about this solution is that you don't even need Visio when generating stencil files, but you habe to read, parse and translate SVG format to the visio format and to make it more difficult, working directly with open document format is not very easy.

### Decision time
Since I already knew a little bit how to work with Office object model and usually everyone who needs Visio stencils should already have Visio installed somewhere to use, I thought I'd go with the first solution.

## How to use
* You can either download all the stencils from here and add them to Visio.
* Or if you need to create your own stencils from other SVG files or tweak the shapes a little, you can get the Powershell cmdlet and use it.

### Adding stencils to Visio

[TO BE COMPLETED]

### Using New-VisioStencil cmdlet

[TO BE COMPLETED]
