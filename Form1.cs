using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Globalization;
using System.Diagnostics;
using System.Collections;
using System.Xml;
using System.Security;
using System.Security.AccessControl;
using System.Data.OleDb;

//1.0	    21/10/2014	    Initial version	                    C. Parlayan
//1.1       25/11/2014      Added Search textbox                C. Parlayan
//1.2       26/11/2014      Fixed UNGROUPED problem             C. Parlayan
//1.3-13    December 2014   Fixed complex rule bugs             C. Parlayan

namespace OCRuleTranslator
{
    public partial class Form1 : Form
    {
        public bool DEBUGMODE = false;
        // public bool DEBUGMODE = true;
        Thread MyThread;
        FileStream fpipdf;
        string Mynewline = System.Environment.NewLine;
        static public string workdir = "";
        static public string inputSand = "";
        static public string inputPrd = "";
        static public bool givewarn = true;
        ArrayList SourceRuleLines = new ArrayList();
        ArrayList TargetRuleLines = new ArrayList();
        ArrayList tmpTargetRuleLines = new ArrayList();
        ArrayList RuleOID = new ArrayList();
        ArrayList warnings = new ArrayList();
        public Form1()
        {
            InitializeComponent();
            Menu = new MainMenu();
            if (DEBUGMODE) MessageBox.Show("D E B U G M O D E - !!!!!", "OCRuleTranslator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            Menu.MenuItems.Add("&File");
            Menu.MenuItems[0].MenuItems.Add("&Open", new EventHandler(MenuFileOpenOnClick));
            Menu.MenuItems[0].MenuItems.Add("&Exit", new EventHandler(MenuFileExitOnClick));
            Menu.MenuItems.Add("&Help");
            Menu.MenuItems[1].MenuItems.Add("&User Manual", new EventHandler(MenuHelpHowToOnClick));
            Menu.MenuItems[1].MenuItems.Add("&About", new EventHandler(MenuHelpAboutOnClick));
            buttonExit.BackColor = System.Drawing.Color.LightGreen;
            buttonExit.Enabled = true;
            buttonSrcBrowse.BackColor = System.Drawing.Color.LightGreen;
            buttonTrgBrowse.BackColor = System.Drawing.Color.LightGreen;
            try
            {
                fpipdf = new FileStream("OCRuleTranslator.pdf", FileMode.Open, FileAccess.Read);
            }
            catch (Exception exx)
            {
                MessageBox.Show("Problem opening user manual. Message = " + exx.Message, "OCRuleTranslator", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            // workdir = Directory.GetCurrentDirectory();
        }

        void MenuFileOpenOnClick(object obj, EventArgs ea)
        {
            tbxSource.Focus();
        }

        void MenuFileExitOnClick(object obj, EventArgs ea)
        {
            fpipdf.Close();
            Close();
        }

        void MenuHelpHowToOnClick(object obj, EventArgs ea)
        {
            try
            {
                Process myProcess = new Process();
                myProcess.StartInfo.FileName = "acrord32.exe";
                myProcess.StartInfo.Arguments += '"' + fpipdf.Name + '"';
                myProcess.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(" Failed to start Acrobat reader: \n\n" + e);
            }
        }

        void MenuHelpAboutOnClick(object obj, EventArgs ea)
        {
            FormAbout Section = new FormAbout();
            Section.Show();
            //MessageBox.Show("OCRuleTranslator Version 2.x Lead programmer: C. Parlayan, tested and released by: TraiT Project team, VU Medical Center, Amsterdam, The Netherlands - 2015", Text);
        }

        private void buttonSrcBrowse_Click(object sender, EventArgs e)
        {
            MyFileDialog mfd = new MyFileDialog(workdir);
            if ((mfd.fnxml != "")) tbxSource.Text = mfd.fnxml;
            if (tbxSource.Text.Contains("\\")) workdir = tbxSource.Text.Substring(0, tbxSource.Text.LastIndexOf('\\'));
            //else workdir = Directory.GetCurrentDirectory();
        }

        private void buttonTrgBrowse_Click(object sender, EventArgs e)
        {
            MyFileDialog mfd = new MyFileDialog(workdir);
            if ((mfd.fnxml != "")) tbxTarget.Text = mfd.fnxml;
            if (tbxTarget.Text.Contains("\\")) workdir = tbxTarget.Text.Substring(0, tbxTarget.Text.LastIndexOf('\\'));
            //else workdir = Directory.GetCurrentDirectory();            
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            textBoxOutput.Text = "";
            if (tbxSource.Text == "" || tbxTarget.Text == "")
            {
                MessageBox.Show("Please enter or select correct input files: Both source and target metadata files are mandatory.", "OCRuleTranslator");
                tbxSource.Focus();
                return;
            }
            if (tbxTarget.Text.Contains("\\")) workdir = tbxTarget.Text.Substring(0, tbxTarget.Text.LastIndexOf('\\'));
            else workdir = Directory.GetCurrentDirectory();
            if (!Get_Rules_From_Source(tbxSource.Text)) return;
            if (SourceRuleLines.Count == 0)
            {
                MessageBox.Show("Source metadata file has no rule definitions.", "OCRuleTranslator");
                tbxSource.Focus();
                return;
            }
            progressBar1.Minimum = 0;
            progressBar1.Step = 1;
            progressBar1.Maximum = SourceRuleLines.Count;
            progressBar1.Value = 0;
            textBoxOutput.Text = "Source: " + tbxSource.Text + ", Target: " + tbxTarget.Text + Mynewline;
            textBoxOutput.Text += "Started in directory " + workdir + ". This may take several minutes..." + Mynewline;
            buttonStart.Enabled = false;
            buttonStart.BackColor = SystemColors.Control;
            buttonExit.Enabled = false;
            buttonExit.BackColor = SystemColors.Control;
            buttonCancel.Enabled = true;
            buttonCancel.BackColor = System.Drawing.Color.LightGreen;
            this.Cursor = Cursors.AppStarting;

            if (DEBUGMODE) DoWork();
            else
            {
                MyThread = new Thread(new ThreadStart(DoWork));
                MyThread.IsBackground = true;
                MyThread.Start();
            }
        }

        public bool Get_Rules_From_Source(string theInputFile)
        {
            // Get Rules from source
            bool record = false;
            SourceRuleLines.Clear();
            TargetRuleLines.Clear();
            tmpTargetRuleLines.Clear();
            RuleOID.Clear();
            warnings.Clear();
            try
            {
                using (StreamReader sr = new StreamReader(theInputFile))
                {
                    String line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        line = line.TrimEnd();
                        if (line.Length == 0) continue;
                        if (line.StartsWith("<OpenClinicaRules:Rules xmlns"))
                        {
                            record = true;
                            continue;
                        }
                        if (line == "</OpenClinicaRules:Rules>") record = false;
                        if (record) SourceRuleLines.Add(line);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OCRuleTranslator", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return (false);
            }
            using (StreamWriter swlog = new StreamWriter(workdir + "\\OCRuleTranslator_SourceRules.xml"))
            {
                swlog.WriteLine("*** Source Rule Lines:");
                foreach (string ss in SourceRuleLines) swlog.WriteLine(ss);
            }
            return (true);
        }

        public void DoWork()
        {
            bool error_happend = false;
            bool found = false;
            bool expSem = false;
            for (int i = 0; i < SourceRuleLines.Count; i++)
            {
                progressBar1.PerformStep();
                string theLine = SourceRuleLines[i].ToString().Replace("OpenClinicaRules:", "");
                if (tbxSearchProfile.Text == "" || theLine.Contains(tbxSearchProfile.Text)) found = true;
                if (theLine.Contains("</RuleAssignment") || theLine.Contains("</RuleDef")) // if there are not founds dont write to target
                {
                    if (!error_happend && found && tmpTargetRuleLines.Count > 0)
                    {
                        for (int p = 0; p < tmpTargetRuleLines.Count; p++)
                        {
                            TargetRuleLines.Add(tmpTargetRuleLines[p]);
                            if (tmpTargetRuleLines[p].ToString().Contains("<RuleRef OID=") || (tmpTargetRuleLines[p].ToString().Contains("<RuleDef OID="))) RuleOID.Add(tmpTargetRuleLines[p].ToString()); 
                        }
                        TargetRuleLines.Add(theLine);
                    }
                    tmpTargetRuleLines.Clear();
                    error_happend = false;
                    found = false;
                    continue;
                }   
                
                if (theLine.Contains("<Target>"))
                {
                    // extract target, get new target, add to tmpTargetRuleLines and continue.
                    string partOID = theLine.Replace("<Target>", "").Replace("</Target>", "").Trim();
                    string theNew = GetNewString(partOID);
                    if (theNew == "NOTF") error_happend = true;
                    else tmpTargetRuleLines.Add(theLine.Replace(partOID, theNew));
                    continue;
                }
                if (theLine.Contains("<DestinationProperty OID="))
                {
                    string[] tmp = theLine.Split('"');
                    string theStr = GetNewString(tmp[1]);
                    if (theStr == "NOTF") error_happend = true; 
                    else tmpTargetRuleLines.Add(theLine.Replace(tmp[1], theStr));
                    continue;
                }
                if (theLine.Contains("<Expression>")) expSem = true;
                if (expSem)
                {
                    // extract expression, get new expression with new OID's, add to tmpTargetRuleLines and continue.
                    // (!) expression can be spread to multiple lines. Thats why I use a semaphore here.
                    if (theLine.Contains("</Expression>")) 
                    {
                        expSem = false;
                        theLine = theLine.Replace("</Expression>", " </Expression>");
                    }
                    // now fish I_ or SE_ in theLine to replace with new oid's; if found replace with new; finally add to tmpTargetRuleLines.
                    string tmp = theLine.Replace('>', '$').Replace('(', '$').Replace(' ', '$');
                    string [] stmp = tmp.Split('$');
                    foreach (string one in stmp)
                    {
                        if (one.Trim().StartsWith("SE_") || one.Trim().StartsWith("I_") || one.Trim().StartsWith("F_") || one.Trim().StartsWith("IG_")) // 1.7
                        {
                            string bb = GetNewString(one);
                            if (bb == "NOTF") error_happend = true;
                            theLine = theLine.Replace(one, bb);
                        }
                    }
                    if (!error_happend) tmpTargetRuleLines.Add(theLine.Replace("&quot;", "\"")); // 1.5
                    continue;
                }
                if (!error_happend) tmpTargetRuleLines.Add(theLine.Replace("&quot;", "\""));  // 1.5
            }
            // check if all RuleDef's and RuleRefs have their partners
            foreach (string one in RuleOID)
            {
                string[] prt = one.Split('"');
                bool oidfound = false;
                foreach (string two in RuleOID)
                {
                    string [] prttwo = two.Split('"');
                    if ((prt[0] != prttwo[0]) && prt[1] == prttwo[1]) oidfound = true;   
                }
                if (!oidfound) AppendWarning("Either RuleRef or RuleDef for " + prt[1] + " is not defined or has errors to solve. This can cause errors while uploading new rules");
            }
            if (TargetRuleLines.Count == 0) textBoxOutput.Text += Mynewline + "No rules found to translate!";
            else
            {
                DateTime dt = DateTime.Now;
                string OUTFILE = "";
                if (warnings.Count == 0)
                {
                    textBoxOutput.Text += Mynewline + "Finished successfully.";
                    if (error_happend) OUTFILE = "NewRulesINCOMPLETE_" + dt.Year.ToString() + "-" + dt.Month.ToString() + "-" + dt.Day.ToString() + "-" + dt.Hour.ToString() + "-" + dt.Minute.ToString() + "-" + dt.Second.ToString() + ".xml";
                    OUTFILE = "NewRules_" + dt.Year.ToString() + "-" + dt.Month.ToString() + "-" + dt.Day.ToString() + "-" + dt.Hour.ToString() + "-" + dt.Minute.ToString() + "-" + dt.Second.ToString() + ".xml";
                }
                else
                {
                    OUTFILE = "NewRulesINCOMPLETE_" + dt.Year.ToString() + "-" + dt.Month.ToString() + "-" + dt.Day.ToString() + "-" + dt.Hour.ToString() + "-" + dt.Minute.ToString() + "-" + dt.Second.ToString() + ".xml";
                    bool war = true;
                    foreach (string one in warnings) if (one.StartsWith("Warning:") == false) war = false;
                    if (war) textBoxOutput.Text += Mynewline + "Finished with warnings: You can ignore them or manually edit the generated rule file.";
                    else textBoxOutput.Text += Mynewline + "Finished with errors. The generated rule file is INCOMPLETE and can't be used. Please correct and try again.";
                }
                using (StreamWriter swlog = new StreamWriter(workdir + "\\" + OUTFILE))
                {
                    swlog.WriteLine("<RuleImport>");
                    for (int i = 0; i < TargetRuleLines.Count; i++) swlog.WriteLine(TargetRuleLines[i]);
                    swlog.WriteLine("</RuleImport>");
                }
            }
            buttonExit.Enabled = true;
            buttonCancel.Enabled = false;
            buttonExit.BackColor = System.Drawing.Color.LightGreen;
            buttonCancel.BackColor = SystemColors.Control;
            this.Cursor = Cursors.Arrow;
            buttonStart.Enabled = true;
            buttonStart.BackColor = System.Drawing.Color.LightGreen;
        }

        private string GetNewString(string old)
        {
            string[] components = old.Trim().Split('.');
            string repidx = "";
            if (components.Length == 4) // se.crf.gr.item
            {
                if (components[1].Contains("_00")) givewarn = true;
                else givewarn = false;
                repidx = GetRepeatIndex(components[0]);
                string SEName = "";
                string NewSEOID = "";
                if (repidx != "")
                {
                    SEName = GetSEName(components[0].Replace(repidx, ""), tbxSource.Text);
                    if (SEName != "NOTF") NewSEOID = GetNewSEOID(SEName.Trim(), tbxTarget.Text);
                    else return ("NOTF");
                    NewSEOID += repidx;
                }
                else
                {
                    SEName = GetSEName(components[0], tbxSource.Text);
                    if (SEName != "NOTF") NewSEOID = GetNewSEOID(SEName.Trim(), tbxTarget.Text);
                    else return ("NOTF");
                }
                string FName = GetFName(components[1], tbxSource.Text);
                string NewFOID = "";
                if (FName != "NOTF") NewFOID = GetNewFOID(FName, tbxTarget.Text);
                else return ("NOTF");
                string GrName = "";
                string NewGrOID = "";
                repidx = GetRepeatIndex(components[2]);

                string tmp = "";
                if (repidx != "") tmp = components[2].Replace(repidx, "");
                else tmp = components[2];

                if (tmp.Contains("_UNGROUPED"))  // get group OID from Formname
                {
                    NewGrOID = GetUngroupedOID(NewFOID, tbxTarget.Text);
                    if (NewGrOID == "NOTF") return ("NOTF");
                }
                else  // get group OID from Group name
                {
                    GrName = GetGrName(tmp, tbxSource.Text);
                    if (GrName != "NOTF") NewGrOID = GetNewGrOID(GrName, NewFOID, tbxTarget.Text);
                    else return ("NOTF");
                }
                if (repidx != "") NewGrOID += repidx;
                string NewItemOID = GetNewItemOID(components[3]);
                return (NewSEOID + "." + NewFOID + "." + NewGrOID + "." + NewItemOID);
            }
            else if (components.Length == 2 || (components.Length == 1 && old.StartsWith("IG_")) ) // gr.item or gr
            {
                givewarn = false;
                string GrName = "";
                string NewGrOID = "";
                repidx = GetRepeatIndex(components[0]);
                string tmp = "";
                if (repidx != "") tmp = components[0].Replace(repidx, "");
                else tmp = components[0];
                string NewFOID = "";
                string oldFormOID = GetOldFormOID(tmp, tbxSource.Text);
                string FName = GetFName(oldFormOID, tbxSource.Text);
                if (FName != "NOTF") NewFOID = GetNewFOID(FName, tbxTarget.Text);
                else return ("NOTF");
                if (tmp.Contains("_UNGROUPED"))  // get group OID from Formname
                {
                    NewGrOID = GetUngroupedOID(NewFOID, tbxTarget.Text);
                    if (NewGrOID == "NOTF") return ("NOTF");
                }
                else  // get group OID from Group name
                {
                    GrName = GetGrName(tmp, tbxSource.Text);
                    if (GrName != "NOTF") NewGrOID = GetNewGrOID(GrName, NewFOID, tbxTarget.Text);
                    else return ("NOTF");
                }
                if (repidx != "") NewGrOID += repidx;
                if (components.Length == 1) return (NewGrOID);
                else
                {
                    string NewItemOID = GetNewItemOID(components[1]);
                    return (NewGrOID + "." + NewItemOID);
                }
            }
            else if (components.Length == 3) // crf.gr.item
            {
                if (components[0].Contains("_00")) givewarn = true;
                else givewarn = false;
                string FName = GetFName(components[0], tbxSource.Text);
                string NewFOID = "";
                if (FName != "NOTF") NewFOID = GetNewFOID(FName, tbxTarget.Text);
                else return ("NOTF");
                string GrName = "";
                string NewGrOID = "";
                repidx = GetRepeatIndex(components[1]);

                string tmp = "";
                if (repidx != "") tmp = components[1].Replace(repidx, "");
                else tmp = components[1];

                if (tmp.Contains("_UNGROUPED"))  // get group OID from Formname
                {
                    NewGrOID = GetUngroupedOID(NewFOID, tbxTarget.Text);
                    if (NewGrOID == "NOTF") return ("NOTF");
                }
                else  // get group OID from Group name
                {
                    GrName = GetGrName(tmp, tbxSource.Text);
                    if (GrName != "NOTF") NewGrOID = GetNewGrOID(GrName, NewFOID, tbxTarget.Text);
                    else return ("NOTF");
                }
                if (repidx != "") NewGrOID += repidx;
                string NewItemOID = GetNewItemOID(components[2]);
                return (NewFOID + "." + NewGrOID + "." + NewItemOID);
            }
            else
            {
                givewarn = false;
                return (GetNewItemOID(old)); // only item: can be = "NOTF"
            }
        }

        private string GetOldFormOID(string groupOID, string filename)
        {
            string tolook = "<OpenClinica:ItemGroupDetails ItemGroupOID=\"" + groupOID + "\">";
            bool sem = false;
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook))
                    {
                        sem = true;
                        continue;
                    }
                    if (sem)
                    {
                        string[] prt = line.Split('"');
                        return (prt[1].Trim());
                    }
                }
            }
            AppendWarning("Ungrouped group OID not found: " + groupOID + "in file " + filename);
            return ("NOTF");
        }

        private string GetUngroupedOID(string theFrom, string filename)
        {
            string tolook = "<FormDef OID=\"" + theFrom;
            bool sem = false;
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook) && line.Contains("Name="))
                    {
                        sem = true;
                        continue;
                    }
                    if (sem)
                    {
                        string[] prt = line.Split('"');
                        return (prt[1].Trim());
                    }
                }
            }
            AppendWarning("Group OID not found in form: " + theFrom);
            return ("NOTF");
        }

        private string GetSEName(string partOID, string filename)
        {
            string tolook = "<StudyEventDef OID=\"" + partOID + "\"";
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook) && line.Contains("Name=") && line.Contains("Repeating="))
                    {
                        string desc = line.Substring(line.IndexOf("Name=") + 6);
                        string[] pp = desc.Split('"');
                        return (pp[0].Trim());
                    }
                }
            }
            AppendWarning("Study Event not found in source: " + partOID);
            return ("NOTF");
        }

        private string GetFName(string partOID, string filename)
        {
            string tolook = "<FormDef OID=\"" + StripVersion(partOID, false, filename);
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook) && line.Contains("Name=") && line.Contains("Repeating="))
                    {
                        string desc = line.Substring(line.IndexOf("Name=") + 6);
                        return (desc.Substring(0, desc.LastIndexOf(" - ")).Trim()); // Form name has <name> - v1.0; we get rid of the last version information here
                    }
                }
            }
            AppendWarning("Form not found in source: " + partOID);
            return ("NOTF");
        }

        private string GetGrName(string partOID, string filename)
        {
            string tolook = "<ItemGroupDef OID=\"" + partOID + "\"";
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook) && line.Contains("Name=") && line.Contains("Repeating="))
                    {
                        string desc = line.Substring(line.IndexOf("Name=") + 6);
                        string[] pp = desc.Split('"');
                        return (pp[0].Trim());
                    }
                }
            }
            AppendWarning("Group not found in source: " + partOID);
            return ("NOTF");
        }

        private string GetNewSEOID(string SEName, string filename)
        {
            string tolook = "Name=\"" + SEName + "\"";
            string save1 = "";
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    string tmp = line;
                    for (int m = 0; m < tmp.Length; m++) tmp = tmp.Replace(" \"", "\""); // eliminate trainling spaces in SE name
                    if (tmp.Contains(tolook) && line.Contains("Repeating=") && line.Contains("<StudyEventDef OID="))
                    {
                        save1 = line.Substring(0, line.IndexOf("Name="));
                        string[] pp = save1.Split('"');
                        return (pp[1]);
                    }
                }
            }
            AppendWarning("Study Event not found in target: " + SEName);
            return ("NOTF");
        }

        private string GetNewFOID(string FName, string filename)
        {
            string tolook = "Name=\"" + FName + " ";
            string save1 = "";
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook) && line.Contains("Repeating=") && line.Contains("<FormDef OID="))
                    {
                        save1 = line.Substring(0, line.IndexOf("Name="));
                        string[] pp = save1.Split('"');
                        return (StripVersion(pp[1], true, filename));
                    }
                }
            }
            AppendWarning("Form not found in target: " + FName);
            return ("NOTF");
        }

        private string GetNewGrOID(string GrName, string FOID, string filename)
        {
            string tolook = "Name=\"" + GrName + "\"";
            string save1 = "";
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook) && line.Contains("Repeating=") && line.Contains("<ItemGroupDef OID="))
                    {
                        save1 = line.Substring(0, line.IndexOf("Name="));
                        string[] pp = save1.Split('"');
                        if (StripVersion(GetOldFormOID(pp[1], filename), true, filename) == FOID) return (pp[1]);
                    }
                }
            }
            AppendWarning("Group not found in target: " + GrName);
            return ("NOTF");
        }

        private string GetNewItemOID(string partOID)
        {
            string OIDWithoutTrailing = OIDContainsTrailingDigits(partOID);
            //if (OIDWithoutTrailing == "NOTD") return (partOID);// item has no trailing _dddd
            //else // item has trailing _dddd
            {
                string[] ItemNameFormID = GetNameFormOID(partOID, tbxSource.Text).Split('$');
                string ItemName = ItemNameFormID[0];
                string FormID = ItemNameFormID[1];
                if (ItemName == "NOTF" || FormID == "NOTF") return ("NOTF");
                // 1.5
                string FName = GetFName(FormID, tbxSource.Text);
                string NewFOID = "";
                if (FName != "NOTF") NewFOID = GetNewFOID(FName, tbxTarget.Text);
                else return ("NOTF");
                
                string itemandform = GetItemOIDFromTarget(ItemName, NewFOID, tbxTarget.Text);
                if (itemandform.Contains("$")) return (itemandform.Substring(0, itemandform.IndexOf('$')));
                else return (itemandform);
            }
        }

        private string StripVersion(string orig, bool all, string filename)
        {
            string strippedOID = GetParentFormOID(orig, filename);
            if ((strippedOID == "NOTF") || (orig == strippedOID)) return (orig);
            else
            {
                if (all && givewarn) AppendWarning("Warning: " + orig + " has version number: " + orig.Replace(strippedOID, "") + ". This will be ignored during translation."); 
            }
            return (strippedOID);
        }

        private string GetRepeatIndex(string orig)
        {
            if (orig.IndexOf('[') < 0) return ("");
            else return (orig.Substring(orig.IndexOf('[')));
        }

        private string GetItemOIDFromTarget(string name, string form, string filename)
        {
            string newoid = "";
            string newfor = "";
            string tolook1 = "<ItemDef OID=";
            string tolook2 = "Name=\"" + name + "\"";
            string tolook3 = "FormOIDs=\"" + form;
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook1) && line.Contains(tolook2) && (line.Contains(tolook3 + "\"") || line.Contains(tolook3 + ",") || line.Contains(tolook3 + "_"))) 
                    {
                        newfor = tolook3.Substring("FormOIDs=\"".Length);  
                        newoid = line.Trim().Substring(tolook1.Length + 1);
                        string[] tmp2 = newoid.Split('"');
                        newoid = tmp2[0];
                        return (newoid + "$" + newfor);
                    }
                }
            }
            return ("NOTF");
        }

        private string GetNameFormOID(string ItemOID, string filename)
        {
            string tolook = "<ItemDef OID=\"" + ItemOID + "\"";
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook))
                    {
                        string name = line.Substring(line.IndexOf("Name=") + 6);
                        name = name.Substring(0, name.IndexOf('"'));
                        string form = line.Substring(line.IndexOf("FormOIDs=") + 10);
                        form = form.Replace(',', '"');
                        form = form.Substring(0, form.IndexOf('"'));
                        return (name + "$" + form);
                    }
                }
            }
            AppendWarning("String not found: " + tolook + " in file " + filename);
            return ("NOTF$NOTF");
        }

        private string GetParentFormOID(string FormOID, string filename)
        {
            // <OpenClinica:FormDetails FormOID="F_TRACERLICHAM_6975_12" ParentFormOID="F_TRACERLICHAM_6975">
            string tolook = "<OpenClinica:FormDetails FormOID=\"" + FormOID + "\" ParentFormOID=";
            using (StreamReader sr = new StreamReader(filename))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.TrimEnd();
                    if (line.Length == 0) continue;
                    if (line.Contains(tolook))
                    {
                        string res = line.Substring(line.IndexOf("ParentFormOID=") + 15);
                        return (res.Substring(0, res.IndexOf('"')));
                    }
                }
            }
            return ("NOTF");
        }
        private void buttonExit_Click_1(object sender, EventArgs e)
        {
            Close();
        }

        private string OIDContainsTrailingDigits(string theOID)
        {
            string[] parts = theOID.Split('_');
            int len = parts.Length; 
            if (len > 0 && parts[len - 1].Length > 0 && IsNumber(parts[len - 1])) return (theOID.Substring(0, theOID.Length - (parts[len - 1].Length + 1))); 
            return ("NOTD");
        }

        private void buttonCancel_Click_1(object sender, EventArgs e)
        {
            if (MyThread != null && MyThread.IsAlive)
            {
                if (MessageBox.Show("Process is not finished yet. If you cancel the process now, the generated rules will be INCOMPLETE AND CAN NOT BE USED. Are you sure you want to cancel?", "OCRuleTranslator asks confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No) return;
                MyThread.Abort();
                buttonExit.Enabled = true;
                buttonExit.BackColor = System.Drawing.Color.LightGreen;
                this.Cursor = Cursors.Arrow;
                MyThread = null;
            }
        }

        public bool IsNumber(String s)
        {
            bool value = true;
            foreach (Char c in s.ToCharArray())
            {
                value = value && Char.IsDigit(c);
            }

            return value;
        }

        public void AppendWarning(string s)
        {
            bool found = false;
            foreach (string one in warnings)
            {
                if (one == s)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                warnings.Add(s);
                DateTime dt = DateTime.Now;
                textBoxOutput.Text += s + Mynewline;
            }
        }
    }
}