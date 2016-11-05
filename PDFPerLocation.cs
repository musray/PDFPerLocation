/*
Luc Morin (MRN), EPLAN Software and Services, November 2016
*/


/*
The following compiler directive is necessary to enable editing scripts
within Visual Studio.

It requires that the "Conditional compilation symbol" SCRIPTENV be defined 
in the Visual Studio project properties

This is because EPLAN's internal scripting engine already adds "using directives"
when you load the script in EPLAN. Having them twice would cause errors.
*/

#if SCRIPTENV
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
#endif

/*
On the other hand, some namespaces are not automatically added by EPLAN when
you load a script. Those have to be outside of the previous conditional compiler directive
*/

using System.Windows.Forms;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Training.Example00
{

    class PDFPerLocation
    {
        [DeclareAction("PDFPerLocationAction")]
        public void PDFPerLocationAction(string rootFolder)
        {

            /* This script loads the following scheme files from the $(MD_SCHEME) folder:
             * 
             * LB.Mounting_Location_List.xml : Labeling scheme to export project's Sturecture Identifiers. 
             *      Relies on the FGfiSO.Mounting_Locations_only.xml labeling filter
             * FGfiSO.Mounting_Locations_only.xml : Labeling filter scheme to export only the Mounting Location Structure Identifiers
             * PDs.Mounting_Locations.xml : PDF export scheme. Relies on the PBfiN.Mounting_Locations.xml Page filter
             * PBfiN.Mounting_Locations.xml: Page filter based on a Mounting Location
             * 
             * It also uses a "template" for the PBfiN.Mounting_Locations.xml Page filter. 
             *      It does string substitution to set a specific Mounting Location as the filter criteria.
             *      After string substitution, the resulting Page filter scheme file is reloaded to set the desired ML
             * 
             */

            CommandLineInterpreter CLI = new CommandLineInterpreter();
            ActionCallingContext ctx = new ActionCallingContext();
            Settings set = new Settings();

            try
            {
                //Load the PDF Export scheme. If a scheme with the same name already exists in P8, it will be overwritten.
                string pathPDFScheme = Path.Combine(PathMap.SubstitutePath("$(MD_SCHEME)"), "PDs.Mounting_Locations.xml");
                set.ReadSettings(pathPDFScheme);

                //Load the Page Filter scheme. If a scheme with the same name already exists in P8, it will be overwritten.
                string pathFilterScheme = Path.Combine(PathMap.SubstitutePath("$(MD_SCHEME)"), "FGfiSO.Mounting_Locations_only.xml");
                set.ReadSettings(pathFilterScheme);

                //Load the labeling scheme. If a scheme with the same name already exists in P8, it will be overwritten.
                string pathLabelingScheme = Path.Combine(PathMap.SubstitutePath("$(MD_SCHEME)"), "LB.Mounting_Location_List.xml");
                string pathMLOutput = Path.Combine(PathMap.SubstitutePath("$(TMP)"), "AllMountingLocations.txt");
                set.ReadSettings(pathLabelingScheme);

                //Export the project's Mounting Locations to temporary file, one per line
                ctx.AddParameter("configscheme", "Mounting_Location_List");
                ctx.AddParameter("destinationfile", pathMLOutput);
                ctx.AddParameter("language", "en_US");
                CLI.Execute("label", ctx);

                //Read the Mounting Location file as a List<string> for later iteration
                List<string> MLs = File.ReadLines(pathMLOutput).ToList();

                string pathTemplate = Path.Combine(PathMap.SubstitutePath("$(MD_SCHEME)"), "PBfiN.Mounting_Locations_template.xml");
                string pathPageFilterScheme = Path.Combine(PathMap.SubstitutePath("$(MD_SCHEME)"), "PBfiN.Mounting_Locations.xml");

                //Loop through each Mounting Location in the "label" file
                foreach (string ML in MLs)
                {
                    //Init Progress bar
                    using (Progress prg = new Progress("SimpleProgress"))
                    {
                        prg.BeginPart(100, "");
                        prg.SetTitle("Exporting: " + ML);
                        prg.SetAllowCancel(false);
                        prg.ShowImmediately();

                        //Read the "template" file, and substitute the real Mounting Location
                        string content = File.ReadAllText(pathTemplate);
                        content = content.Replace("!!MountingLocation!!", ML);

                        //Save the result back to disk, and load the resulting Page Filter scheme
                        File.WriteAllText(pathPageFilterScheme, content);
                        set.ReadSettings(pathPageFilterScheme);

                        //Call the PDF export, using the scheme that makes use of the Page Filter
                        ctx = new ActionCallingContext();
                        ctx.AddParameter("TYPE", "PDFPROJECTSCHEME");
                        ctx.AddParameter("EXPORTSCHEME", "Mounting_Locations");
                        ctx.AddParameter("EXPORTFILE", Path.Combine(rootFolder, ML));
                        CLI.Execute("export", ctx);

                        prg.EndPart(true);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

    }

}
