# PDFPerLocation
EPLAN P8 script to generate a PDF file per location defined in project

To use this script:
	
- The XML files must be copied to the EPLAN $(MD_SCHEME) folder.
- The PDFPerLocation.cs script must copied to the EPLAN $(MD_SCRIPTS) folder, and loaded in EPLAN P8
- A toolbar button must be created, calling the PDFPerLocationAction, passing the destination folder as the /rootFolder parameter:

`PDFPerLocationAction /rootFolder:"C:\Temp"`

The PDFPerLocationAction action can also be called from other scripts or API Add-ins.

Description of the XML files:

- LB.Mounting_Location_List.xml : Labeling scheme to export project's Structure Identifiers. 
- FGfiSO.Mounting_Locations_only.xml : Labeling filter scheme to export only the Mounting Location Structure Identifiers
- PDs.Mounting_Locations.xml : PDF export scheme. Relies on the PBfiN.Mounting_Locations.xml Page filter
- PBfiN.Mounting_Locations.xml: Page filter based on a Mounting Location
- PBfiN.Mounting_Locations_template.xml: Template for creating the Page filters based on the project's Mounting Locations. 
