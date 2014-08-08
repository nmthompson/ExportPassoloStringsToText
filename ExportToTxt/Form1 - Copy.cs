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
        String Standard;
        String Glossary;
        


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
            else
                languageForPath = "French";
            
        }
        /*
        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                checkPath = folderBrowserDialog1.SelectedPath;
                paths.Add(folderBrowserDialog1.SelectedPath);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                checkPath = folderBrowserDialog1.SelectedPath;
                paths.Add(folderBrowserDialog1.SelectedPath);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                checkPath = folderBrowserDialog1.SelectedPath;
                paths.Add(folderBrowserDialog1.SelectedPath);
            }
        }*/

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
                        string[] lpuFiles = System.IO.Directory.GetFiles(sourcePath, languageSelectionText + "*.lpu", SearchOption.AllDirectories);
                        PassoloApp PSL = new PassoloApp();
                        PSL.Visible = true;
                        

                        for (int i = 0; i < lpuFiles.Length; i++)
                        {
                            PSL.Projects.Open(lpuFiles[i]);


                            String path = @"\\FileSrvwhq2\globalops\global team\Passolo Projects\" + languageForPath;
                            string[] destinationPath = System.IO.Directory.GetDirectories(path, "*Glossaries*", SearchOption.TopDirectoryOnly);
                            String dateTime = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString();
                            System.IO.Directory.CreateDirectory(System.IO.Path.Combine(destinationPath[0], dateTime));
                            System.IO.Directory.CreateDirectory(System.IO.Path.Combine(destinationPath[0], dateTime, Path.GetFileName(sourcePath)));
                            String exactPath = System.IO.Path.Combine(destinationPath[0], dateTime, Path.GetFileName(sourcePath), Path.GetFileName(lpuFiles[i]).Replace(".lpu", "") + ".txt");
                            //Console.WriteLine(PSL.ActiveProject.Name);


                            PslTransBundle newBundle = PSL.ActiveProject.PrepareTransBundle();

                            for (int p = 0; p < PSL.ActiveProject.SourceLists.Count; p++)
                            {
                                newBundle.AddTransList(PSL.ActiveProject.TransLists.Item(p + 1));
                                newBundle.AddTransList(PSL.ActiveProject.SourceLists.Item(p + 1));
                            }
                            PSL.DisplayAlerts = PslAlertLevel.pslAlertsNone;
                            if (formatSelection == Glossary)
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
                                                Console.WriteLine(allLines[line - 1]);
                                                Console.WriteLine(allLines[line]);
                                                Console.WriteLine(allLines[line + 1]);
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
                            else if (formatSelection == Standard)
                            {
     
                              // File.CreateText(exactPath);
                              PSL.ActiveProject.Export("Trados Text Export", newBundle, exactPath);
                               

                                //duplicates
                                //folder
                                //.txt or .csv
                            }


                            PSL.ActiveProject.Close();
                            pathToPrint = Path.GetDirectoryName(exactPath);
                            pathToPrint = Path.GetDirectoryName(pathToPrint);
                        }
                        
                    }

                    

                    MessageBox.Show("TXT Files were exported to the location: " + pathToPrint);

                    button2.Text = "Finished";
                    button2.Refresh();
                }
                else
                {
                    MessageBox.Show("You need to select at least one source path and the language field.");
                }
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
                formatSelection = Standard;
            }
            else if (formatSelection == "Glossary .txt")
            {
                formatSelection = Glossary;
            }
            
        }

        

       

      

        
    }
}
