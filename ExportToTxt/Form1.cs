using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using PassoloU;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExportToTXTForProjects
{    

    public partial class Form1 : Form
    {
        List<String> paths = new List<String>();
        String sourcePath;
        String checkPath;
        String languageSelectionText;
        String languageForPath;
        String pathForOne = "";
        String pathForTwo = "";
        String pathForThree = "";
        String pathToPrint = "";
        String formatSelection;
        string[] lpuFiles = new string[0];
        // String Standard;
        //String Glossary;
        string[] destinationPath = new string[0];
        String exactPath;
        


        public Form1()
        {
            InitializeComponent();
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            languageSelectionText = comboBox1.SelectedItem.ToString();

            if(languageSelectionText == "ESN")
            {
                languageForPath = "Spanish";
            }
            else if(languageSelectionText == "DEU")
            {
                languageForPath = "German";
            }
            else if(languageSelectionText == "ENG")
            {
                languageForPath = "UK English";
            }
       
            else if (languageSelectionText == "FRA")
            {
                languageForPath = "French";
            }
            else
                languageForPath = "Portuguese";
            
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (CheckInPaths(pathForOne) && CheckInPaths(pathForTwo) && CheckInPaths(pathForThree))
            {

                if (!string.IsNullOrEmpty(checkPath) && !string.IsNullOrEmpty(languageSelectionText))
                {
                    for (int q = 0; q < paths.Count; q++)
                    {
                        sourcePath = paths[q];
                        button2.Text = "Running...";
                        button2.Refresh();
                        StepThroughDirectories(sourcePath);
                        //string[] lpuFiles = System.IO.Directory.GetFiles(sourcePath, languageSelectionText + "*.lpu", SearchOption.AllDirectories);
                        PassoloApp PSL = new PassoloApp();
                        PSL.Visible = false;


                        for (int i = 0; i < lpuFiles.Length; i++)
                        {
                            if (lpuFiles[i].ContainsMe("\\old\\") || lpuFiles[i].ContainsMe("\\bak\\") || lpuFiles[i].ContainsMe("\\backup\\") || lpuFiles[i].ContainsMe("\\temp\\") || lpuFiles[i].ContainsMe("test"))
                            {   //ignore file
                                Console.WriteLine(lpuFiles[i]);
                            }
                            else
                            {
                                PSL.Projects.Open(lpuFiles[i]);

                                String path = @"\\FileSrvwhq2\globalops\global team\Passolo Projects\" + languageForPath;
                                if (formatSelection == "Glossary")
                                {
                                    destinationPath = System.IO.Directory.GetDirectories(path, "*Glossaries*", SearchOption.TopDirectoryOnly);

                                }
                                else
                                {
                                    destinationPath = System.IO.Directory.GetDirectories(path, "*Standard*", SearchOption.TopDirectoryOnly);
                                }
                                String dateTime = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString();
                                System.IO.Directory.CreateDirectory(System.IO.Path.Combine(destinationPath[0], dateTime));
                                System.IO.Directory.CreateDirectory(System.IO.Path.Combine(destinationPath[0], dateTime, Path.GetFileName(sourcePath)));
                                exactPath = System.IO.Path.Combine(destinationPath[0], dateTime, Path.GetFileName(sourcePath), Path.GetFileName(lpuFiles[i]).Replace(".lpu", "") + ".txt");
                                Console.WriteLine(PSL.ActiveProject.Name);
                                

                                PslTransBundle newBundle = PSL.ActiveProject.PrepareTransBundle();

                                for (int p = 0; p < PSL.ActiveProject.SourceLists.Count; p++)
                                {
                                    newBundle.AddTransList(PSL.ActiveProject.TransLists.Item(p + 1));
                                    newBundle.AddTransList(PSL.ActiveProject.SourceLists.Item(p + 1));
                                }
                                PSL.DisplayAlerts = PslAlertLevel.pslAlertsNone;
                                if (formatSelection == "Glossary")
                                {
                                    PSL.ActiveProject.Export("Trados Text Export", newBundle, exactPath);

                                    DateTime firstTime = DateTime.Now;
                                    var allLines = File.ReadAllLines(exactPath).ToList();

                                    int line = 0;

                                    for (int p = 1; p <= PSL.ActiveProject.TransLists.Count; p++)
                                    {
                                        for (int n = 1; n <= PSL.ActiveProject.TransLists.Item(p).StringCount; n++)
                                        {

                                            Boolean contains = false;
                                            while (contains == false)
                                            {
                                                if (line == allLines.Count)
                                                {
                                                    contains = true;
                                                }
                                                else if (allLines[line].Contains("<Seg L=es-ES>"))
                                                {
                                                    contains = true;
                                                    string yes = PSL.ActiveProject.TransLists.Item(p).get_String(n).TransComment.ToString();
                                                    if (yes.Equals("T9N"))
                                                    {

                                                    }
                                                    string no = PSL.ActiveProject.TransLists.Item(p).get_String(n).Text.ToString();
                                                    allLines.Insert(line + 1, "<Comment> " + PSL.ActiveProject.TransLists.Item(p).get_String(n).TransComment.ToString());
                                                    //Console.WriteLine(allLines[line - 1]);
                                                    //Console.WriteLine(allLines[line]);
                                                    //Console.WriteLine(allLines[line + 1]);
                                                    line++;
                                                }
                                                else
                                                    line++;
                                            }


                                        }
                                    }
                                    File.WriteAllLines(exactPath, allLines.ToArray());
                                    DateTime lastTime = DateTime.Now;
                                    TimeSpan total = lastTime - firstTime;
                                }
                                else if (formatSelection == "Standard")
                                {

                                    PSL.ActiveProject.Export("Customizable Text Export", newBundle, exactPath);
                                   
                                }
                                else
                                {
                                    MessageBox.Show("Please select a format");
                                    return;
                                }

                                PSL.ActiveProject.Close();
                                pathToPrint = Path.GetDirectoryName(exactPath);
                                pathToPrint = Path.GetDirectoryName(pathToPrint);
                            }

                          
                        }

                        PSL.Quit();
                        if (formatSelection == "Standard")
                        {
                            //concatanate files
                            string[] fileCount = Directory.GetFiles(System.IO.Directory.GetParent(exactPath).ToString(), "*.txt", SearchOption.AllDirectories);
                            System.Text.StringBuilder sb = new System.Text.StringBuilder();
                            if (fileCount.Length > 1)
                            {
                                for (int count = 0; count < fileCount.Count(); count++)
                                {
                                    sb.Append(System.IO.File.ReadAllText(fileCount[count], Encoding.Default));

                                }
                                string buildOutput = sb.ToString();
                                string newFilePath = System.IO.Directory.GetParent(exactPath).ToString() + @"\newMergedFile.txt";
                                System.IO.File.WriteAllText(newFilePath, buildOutput, Encoding.Default);


                                //open .txt with excel
                                Excel.Application excelApp = new Excel.Application();
                                excelApp.Workbooks.OpenText(newFilePath);
                                excelApp.Visible = true;
                                //Check for and remove duplicates
                                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                                xlWorkSheet.Columns.RemoveDuplicates(1, Excel.XlYesNoGuess.xlNo);


                            }
                        }


                        MessageBox.Show("TXT Files were exported to the location: " + pathToPrint);

                        button2.Text = "Finished";
                        button2.Refresh();
                       
                    }
                 
                }

            }
        }

        void StepThroughDirectories(string dir)
        {
            StepThroughDirectories(dir, 0);
        }

        void StepThroughDirectories(string dir, int currentDepth)
        {
            // process 'dir'
            if (languageSelectionText == "ENG")
            {
             lpuFiles = lpuFiles.Concat(Directory.GetFiles(dir, "UK"  + "*.lpu")).ToArray();
             lpuFiles = lpuFiles.Concat(Directory.GetFiles(dir, languageSelectionText  + "*.lpu")).ToArray();
            }
            else
            lpuFiles = lpuFiles.Concat(Directory.GetFiles(dir, languageSelectionText + "*.lpu")).ToArray();          

            // process subdirectories, limit depth to 3
            if (currentDepth < 3)
            {
                foreach (string subdir in Directory.GetDirectories(dir))
                    StepThroughDirectories(subdir, currentDepth + 1);
            }
        }

        private Boolean CheckInPaths(string x )
        {
            if (!string.IsNullOrEmpty(x))
            {
                if (System.IO.Directory.Exists(@x) == true)
                {
                    paths.Add(@x);
                    return true;
                }
                else
                {
                    MessageBox.Show("The path: " + x + " is not valid.");
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            
            checkPath = textBox3.Text;
            pathForThree = textBox3.Text;
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            checkPath = textBox2.Text;
            pathForTwo = textBox2.Text;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            checkPath = textBox1.Text;
            pathForOne = textBox1.Text;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            formatSelection = comboBox2.SelectedItem.ToString();
            if (formatSelection == "Standard .txt")
            {
                formatSelection = "Standard";
            }
            else if (formatSelection == "Glossary .txt")
            {
                formatSelection = "Glossary";
            }
            
        }

        
    }
       

    static class ExtensionHelpers
    {
        internal static bool ContainsMe(this string source, string toCheck)
        {
            return source.IndexOf(toCheck, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }

}