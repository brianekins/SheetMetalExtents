** Description
This is an Add-In that when installed will automatically create and update three custom iProperties.  It should work for Inventor 2013 and later.  The iProperties created are:

SheetMetalLength - The length of the sheet metal flat pattern.
SheetMetalWidth - The width of the sheet metal flat pattern.
SheetMetalStyle - The active sheet metal style (or rule).

It also creates the following two reference parameters:

SheetMetalLength - The length of the sheet metal flat pattern.
SheetMetalWidth - The width of the sheet metal flat pattern.

The SheetMetalLength and SheetMetalWidth iProperties are the result of the two parameters being set to be exposed as iProperties.  The output format for these parameters can be edited to change the formatting of the associated iProperties.  The parameters should not be used as input to other parameters that control the shape of the part because this can result in cycles being created.

Any suggestions or issues should be reported to Brian Ekins at brian.ekins@autodesk.com