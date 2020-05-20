# Azure Visio Stencils based on official icons from Microsoft
I started this project because I needed a decent set of visio stencils to draw my Azure diagrams. I wanted to use the latest official icons from Microsoft and I wanted the shapes to act smart in the diagram just like the native Visio stencils, but I failed to find any. The closest I got was [this repository](https://github.com/benc-uk/icon-collection) that has collected all the official SVG files published by Microsoft and generously shared sith everyone. Thumbs up to [Ben Coleman](https://github.com/benc-uk).

If you are here to download all the these stencils, you are in the right place.

## Stencil Galleries
The followig stencil galleries contain identical icons shared by Ben Coleman in his repository, but converted to Visio Stencils for your convenience.

> 📯 NOTE! 
> Ther might be some shaped in the stencils that won't work. This is because Visio is not able to import them unfortunately. You can check the list int [Known Issues](#known-issues-) section

### Core Azure Sets
#### [Azure Docs](stencils/azure-docs.vssx)
Scraped from main Azure docs site, all major Azure services are in here plus a couple of other useful icons. The main place to start if what you require is an icon for a top level Azure service

#### ['Official' Azure Icons Set](stencils/azure-icons.vssx)
This is the official "Microsoft Azure Icon Repository" set from the Microsoft Cloud Design Studio, the best place to look if you're after an Azure service or other icon from the Azure Portal. There will be an overlap with the Azure Docs set, but there's nearly four times the number of icons here.

#### [Azure Patterns Collection](stencils/azure-patterns.vssx)
Very large set of 1200+ icons including many Azure services, but mostly focused on other concepts, actions and gyphs. This has been scraped from https://azure.microsoft.com/en-gb/patterns/styles/glyphs-icons/ This is made public, as part of the "Sundog" Azure.com design system

#### [Microsoft 'Cloud and AI' Set (Outdated)](stencils/cloud-old.vssx)
This is the old official "Microsoft Azure Cloud and AI Symbol / Icon Set - SVG" dated 15/05/2019 fetched from here. It is extremely outdated, and has many overlaps with the other sets, but also contains many unique icons. It's included here in entirety for completeness

## Other Sets
### [Logos & Brands](stencils/azure-logos.vssx)
Various product, company & programing language logos, hand picked & fetched from various sources online. Theses may or may not be directly Azure related

### [Other Icons](stencils/azure-others.vssx)
Many other Azure & Microsoft icons sourced from various places; Azure.com pages, Azure Docs git repo, Azure portal etc. Most of these are hand picked or moved here manually, these shouldn't overlap with the other sets, but may do

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
```powershell
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

### How to use New-VisioStencil cmdlet
First thing you need is to run New-VisioStencil.ps1 powershell script. It will define a cmdlet called `New-VisioStencil` (obviously). This cmd let is very easy to use and I tried to add documentation (including examples), comments directly in the cmdlet. Here I will give you a few examples. Don't forget to run the following line before trying any of the following examples. The simplest way to use functions a Powershell script without installing any modules.
```powershell
. .\New-VisioStencil.ps1
```

#### Example 1 - Simply creating a Visio stencil from a list of SVG files
The following example will create a stencil called "Stencil1.vssx" from all SVG files in the current folder.
```powershell
New-VisioStencil (Get-ChildItem "*.svg") -StencilPath "Stencil1.vssx"
```
The same example can be written using piping like the following.
```powershell
dir "*.svg" | New-VisioStencil -StencilPath "Stencil1.vssx"
```

#### Example 4 - Custom naming for master shapes
When you simply use `New-VisioStencil` without any parameters other than `StencilPath`, it will use each SVG file's name to name the master shaped in the stencil. You might not necessarily like that. Imangine you have SVG files like the following.
* Analytics-141-SQL-Data-Warehouses.svg
* Analytics-142-HD-Insight-Clusters.svg
* Analytics-143-Data-Lake-Analytics.svg
And you would prefer if your master shapes were named like the following
* SQL Data Warehouses
* HD Insight Clusters
* Data Lake Analytics
It means that everything before the number and the hyphen after it ('-') should be removed and all the hyphens should be replaced with space ('-').
For these scenarios `New-VisioStencil` cmdlet has a special parameter called `MasterNameExtractor`. Using this parameter you can provide your own logic to extract master names from files names. Your logic needs to receive a string (the file name without extension) and return a string (the master name).
```powershell
$nameExtractor = {param($name) ($name | Select-String "^\w+?-\d+?-(.+)").Matches[0].Groups[1].Value.Replace('-',' ') }
dir *.svg" | New-VisioStencil -StencilPath "MyStencil.vssx" -MasterNameExtractor $nameExtractor
```
You could of course embed the logic directly in one line. Although it would be a bit less readable.
```powershell
dir *.svg" | New-VisioStencil -StencilPath "MyStencil.vssx" -MasterNameExtractor {param($name) ($name | Select-String "^\w+?-\d+?-(.+)").Matches[0].Groups[1].Value.Replace('-',' ') }
```

#### Example 3 - Creating multiple Visio stencils from several SVG files - Simple
Some times you might have several SVG files and you don't want to create just one stencil for all of them. You would probably want to categorize them and create a separate stencil for each group. There are several approaches to fix this issue. One approach would be to filter these files based on any criteria and create stencils for them like the following.
```powershell
.\New-VisioStencil.ps1
dir "api*.svg" | New-VisioStencil -StencilPath "API.vssx"
dir "blockchain*.svg" | New-VisioStencil -StencilPath "Bloackchain.vssx"
dir "logos\*.svg" | New-VisioStencil -StencilPath "Logos.vssx"
```

#### Example 4 - Creating multiple Visio stencils from several SVG files - Advanced
If you have SVG files that have some sort of naming pattern you can even create several stencils at once. Let's assume that your SVG files follow this name pattern: `category-number-master-shape-name` and you want to generate stencils with the name `category.vssx` and each master shape in stencil should be named `master shape name`. In other words, the first part of the file name is used for stencil name, the number is removed and the rest of the name is used for the master shape by replacing hyphens ('-') with space (' ').
For example if you have the following files in a folder.
* Analytics-141-SQL-Data-Warehouses.svg
* Analytics-142-HD-Insight-Clusters.svg
* Analytics-143-Data-Lake-Analytics.svg
* Blockchain-363-Applications.svg
* Blockchain-364-Outbound-Connection.svg
That should turn into two stenciles, first one containing the first three files and the second contianing the two other files like the following.
```
Analytics.vssx
├─ SQL Data Warehouses
├─ HD Insight Clusters
└─ Data Lake Analytics
Blockchain.vssx
├─ Applications
└─ Outbound Connection
```
You might even like to display a progress bar for the overal operation, since this is going to take a while if you have thousands of files.

```powershell
cd 'E:\Temp\Visio\Official Azure Icon Set'
$nameExtractor = {param($name) ($name | Select-String "^\w+?-\d+?-(.+)").Matches[0].Groups[1].Value.Replace('-',' ') }
$groupsOfSvgFiles = Get-ChildItem "*.svg" |
    Group-Object -Property @{
        Expression = {$_.BaseName.Substring(0,$_.BaseName.IndexOf('-'))}
    }
$groupsOfSvgFiles |
    ForEach-Object $_ -Begin {
        $i = 0
        Write-Host "Stencils to be created: $($groupsOfSvgFiles.Count)"
    } -Process {
        Write-Host "Creating stencil $($_.Name).vssx with $($_.Group.Count) masters..."
        Write-Progress -Id 1 -Activity "Creating stencils..." -Status "File $($i + 1) of $($groupsOfSvgFiles.Count)" -PercentComplete ($i / $groupsOfSvgFiles.Count * 100)
        New-VisioStencil $_.Group -StencilPath "$($_.Name).vssx" -MasterNameExtractor $nameExtractor -Verbose
        Write-Host "Stencil $($_.Name).vssx completed."
        $i++
    } -End {
        Write-Progress -Id 1 -Activity "Creating stencils..." -Completed
    }
```
If you assign '1' as the Id of your progress bar, `New-VisioStencil` will write its progress as a child of your progress.

## Known Issues 💀
1. Some SVG files cannot be imported. This is a limitation in Visio. Even if you try the drag-n-drop those SVG files in Visio it will give you the following error.
2. There is one file that if you import it to Visio, it will crash without even giving you any error. That file is notebooks.svg.
I will report these as bugs to Microsoft.

### Following is the list files that cannot be imported.
#### azure-docs
* recovery-services-vaults.svg
* security-center.svg
* service-health.svg
* signalr-service.svg
* spring-cloud.svg
* stack.svg
* storage-accounts.svg
* stream-analytics.svg
* time-series-insights-environments.svg
* traffic-manager.svg
* virtual-network-gateways.svg

### cloud-old
* cloud-old\Analysis Services.svg
* cloud-old\App Configuration.svg
* cloud-old\App Services.svg
* cloud-old\Azure Cosmos DB.svg
* cloud-old\Azure Information Protection.svg
* cloud-old\Cloud Services (Classic).svg
* cloud-old\Cloud Services.svg
* cloud-old\CloudSimple Virtual Machines.svg
* cloud-old\Cognitive Services.svg
* cloud-old\Customer Lockbox.svg
* cloud-old\Data Lake Storage.svg
* cloud-old\dedicated_event_hub.svg
* cloud-old\DeveloperTools.svg
* cloud-old\Event Hub Clusters.svg
* cloud-old\Genomics Accounts.svg
* cloud-old\Managed Applications.svg
* cloud-old\Mesh Applications.svg
* cloud-old\Recovery Services Vaults.svg
* cloud-old\Resource Explorer.svg
* cloud-old\Resource Groups.svg
* cloud-old\Resource.svg
* cloud-old\Service Endpoint Policies.svg
* cloud-old\SignalR.svg
* cloud-old\What's New.svg
* cloud-old\Windows 10 IoT Core Services.svg

### logos
* logos\bit-bucket.svg
* logos\docker-icon-wh.svg
* logos\docker-mono.svg
* logos\etcd-icon-color.svg
* logos\flask.svg
* logos\ios.svg
* logos\kafka.svg
* logos\linux-tux-mono.svg
* logos\microsoft.svg
* logos\nodejs-1.svg
* logos\nodejs-3.svg
* logos\python-colour.svg
* logos\red-hat-new.svg
* logos\red-hat.svg
* logos\spark.svg
* logos\twitter-2.svg

### Others
* other\aml-activities.svg
* other\azure-cognitive-services-color.svg
* other\azure-lbs.svg
* other\backup-archive.svg
* other\containerinstances-mono.svg
* other\cubes.svg
* other\develop.svg
* other\genomics.svg
* other\high-performance-computing.svg
* other\notebooks-alt.svg
* other\pickle.svg
* other\resource-explorer.svg
* other\ResourceDefault.svg
* other\work-how-you-want.svg
* CRASH! notebooks.svg
