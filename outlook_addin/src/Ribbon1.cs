using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;



namespace Outlook2Aula
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

            Globals.Ribbons.Ribbon1.lblO2APath.Label = getO2AFolderPath();

            Globals.Ribbons.Ribbon1.lblBuildVersion.Label = "2023-10-25"; //Environment.GetEnvironmentVariable("ClickOnce_CurrentVersion");

            //Oprindeligt fra: https://robindotnet.wordpress.com/2010/07/11/how-do-i-programmatically-find-the-deployed-files-for-a-vsto-add-in/
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;
            Console.WriteLine(assemblyLocation);

            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = System.IO.Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
            Console.WriteLine(ClickOnceLocation);

            //Globals.Ribbons.Ribbon1.label1.Label = assemblyLocation;
            //Globals.Ribbons.Ribbon1.label2.Label = ClickOnceLocation;
        }

        private String getO2AFolderPath()
        {
            return Properties.Settings.Default.O2AFolderPath;
        }

        private void setO2AFolderPath(String O2AFolderPath)
        {
            //Sets the Value, and saves for future use. 
            Properties.Settings.Default.O2AFolderPath = O2AFolderPath;
            Properties.Settings.Default.Save();

            //Updates the Ribbon with the new location
            Globals.Ribbons.Ribbon1.lblO2APath.Label = Properties.Settings.Default.O2AFolderPath;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string strCmdText;
            strCmdText = @"/C cd " + getO2AFolderPath() + " & o2a_run.bat & pause";
            //strCmdText = @"/C cd ""C:\Users\Ole Dahl Frandsen\Documents\GitHub\O2A\"" & python src\main.py & pause""";
            System.Diagnostics.Process.Start("CMD.exe", strCmdText);
        }

        private void btnAllSettings_Click(object sender, RibbonControlEventArgs e)
        {
            string strCmdText;
            strCmdText = @"/C cd "+getO2AFolderPath()+" & python src\\main.py -g";
            //strCmdText = @"/C cd ""C:\Users\Ole Dahl Frandsen\Documents\GitHub\O2A\"" & python src\main.py -g & pause""";
            System.Diagnostics.Process.Start("CMD.exe", strCmdText);
        }

        private void btnForceUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            string strCmdText;
            strCmdText = @"/C cd " + getO2AFolderPath() + " & o2a_run_force.bat & pause";
            //strCmdText = @"/C cd ""C:\Users\Ole Dahl Frandsen\Documents\GitHub\O2A\"" & python src\main.py -f -r & pause""";
            System.Diagnostics.Process.Start("CMD.exe", strCmdText);
        }

        private void btnSelectO2AFolder_Click(object sender, RibbonControlEventArgs e)
        {
            string fileName;
            string folderName;

            //Option 1
            // FolderBrowserDialog - returns only path to the folder
            FolderBrowserDialog fileFolderBrowserDialog = new FolderBrowserDialog();
            fileFolderBrowserDialog.SelectedPath = getO2AFolderPath();
            fileFolderBrowserDialog.ShowDialog();
            folderName = fileFolderBrowserDialog.SelectedPath;

            setO2AFolderPath(folderName);
            

            //fileName = string.Concat(fileSaveAsDialog.FileName, "\\", "your generated file name.
        }

        private void btnOpenPeopleWorkbook_Click(object sender, RibbonControlEventArgs e)
        {
            // declare the application object
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;


            // open a file
            //Excel.Workbook wb = xl.Workbooks.Open(@"C:\Users\Ole Dahl Frandsen\Documents\GitHub\\O2A\personer.csv
            Excel.Workbook wb = xl.Workbooks.Open(@getO2AFolderPath()+"\\personer.csv");

            // do stuff ....

            //close the file
            //wb.Close();

            // close the application and release resources
            //xl.Quit();
        }

        private void btnOpenIgnoreFile_Click(object sender, RibbonControlEventArgs e)
        {
            // declare the application object
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;


            // open a file
            //Excel.Workbook wb = xl.Workbooks.Open(@"C:\Users\Ole Dahl Frandsen\Documents\GitHub\\O2A\personer.csv
            Excel.Workbook wb = xl.Workbooks.Open(@getO2AFolderPath() + "\\personer_ignorer.csv");
        }
    }
}
