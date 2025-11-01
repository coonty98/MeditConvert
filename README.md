# Introduction
In the Dental industry, the digital (CAD/CAM) workflow has taken over and become critical to delivering high-quality products to patients. However, as with all new technology and processes, the various tools and systems do not always function seamlessly together.

One such case is with the numerous Intra-Oral Scanners (IOS) and dental CAD programs. IOS programs output different file types with different data, and not everything will properly import into CAD programs.

For example, some IOS programs allow clinics to mark the margin line for Crown and Bridge cases. Unfortunately, the 3Shape CAD program does not offer a method of importing the margin line unless the scans came from Trios or iTero.

That is where this script comes in. 3Shape uses an import tool that allows users to import iTero cases, including the margin line. It reads certain information from an XML file, then imports the files indicated.

The goal of this script is to convert Medit IOS cases into an 'iTero' case so that the margin line can be imported into 3Shape.

# How it Works
> This has only been tested to work with files exported from the Medit Web portal, not the Medit Desktop app. The desktop app exports differently and is being phased out anyway (per Medit support).

## File Name
This script starts by taking the file name of the parent folder and running it through a function that extracts the patient's name. All Medit cases generally use the same naming conventions, so this script takes those into account. However, clinics may stray from this, so the output may not always be as expected.
## Tooth Number
Then, the tooth number of the case is determined. If a case has a margin file, it will have *MarginLine* in the name, so these files are isolated and the number is extracted. In the case of Medit, it uses the FDI tooth numbering system whereas 3Shape uses the US system. Therefore, the function uses a simple hash table to convert the number from FDI to US notation.

This section also converts the margin file from the Medit **.XYZ** file to a **.PTS** file that 3Shape can understand.
## XML File
The last bit of data that needs to be gathered is the **Exported Objects** which is the section of the XML file that describes the scans/margin files. 

Once all data is gathered, it is passed to the XML function which builds the XML then outputs the file. At this point, the case can be imported into 3Shape.
# Previous Versions
When I originally created this script, I had a master folder of stripped-down XML files that the script would copy to the case parent folder, then modify with the required information. This was not ideal because it required the master folder to always be present and was slower. I then switched methods so that the script was creating the XML from scratch.

I also had multiple XML files which normal iTero cases always generate, but then I learned that 3Shape only uses one of them, so I was able to simplify the script a bit.
# Adaptation
This script can also be used with other IOS programs that output margin files as long as these conditions are met:
- The parent folder name matches normal Medit naming conventions.
- The file names are converted to match Medit naming.
- The margin file can be converted to **.PTS**

# Future Improvements
This script could be improved by adding functionality to identify files from other IOS programs to expand the scope.

In the Medit Desktop app, cases can be exported with custom naming which can include information such as the case ID. It also exports an XML file with additional case information such as the patient name, tooth number, material, etc. At this time, Medit has no intention of adding this functionality the Web application, but if they do, I would like to rewrite this script to gather the required data from the XML file so that cases can be imported into 3Shape with more useful information. For example, the proper material could be automatically selected, or the proper tooth number if there isn't a defined margin line.