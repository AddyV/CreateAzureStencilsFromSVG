using System;
using System.IO;
using System.Windows.Forms;
using MOIV = Microsoft.Office.Interop.Visio;

namespace CreateAzureStencilsFromSVG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MOIV.Application a = null;
            string pf = txtDestination.Text;

            try
            {
                // Validate
                if (txtSource.Text.Length == 0)
                    throw new Exception("Source was not specified");
                else if (!Directory.Exists(txtSource.Text))
                    throw new Exception("Source does not exist");
                else if (new DirectoryInfo(txtSource.Text).GetDirectories().Length == 0)
                    throw new Exception("Azure icons were not proreply downloaded.\nExpecting sub-folders in the source folder but didnt find any.");
                else if (txtDestination.Text.Length == 0)
                    throw new Exception("Destination was not specified");

                // Prepare destination
                if (!Directory.Exists(txtDestination.Text))
                    Directory.CreateDirectory(txtDestination.Text);

                textBox1.Text = "";
                textBox1.Refresh();

                // Open Visio
                a = new MOIV.Application();

                // Loop through the downloaded folders
                foreach (string s in Directory.EnumerateDirectories(txtSource.Text))
                {
                    // Build stencil (file) name
                    DirectoryInfo di = new DirectoryInfo(s);
                    string sn = di.Name;
                    if (!sn.StartsWith("azure", StringComparison.CurrentCultureIgnoreCase))
                        sn = "Azure - " + sn;
                    else
                        sn = sn.Replace("Azure", "Azure -");
                    textBox1.Text += sn + Environment.NewLine;
                    textBox1.Refresh();

                    // Create the new stencil and add svg files
                    MOIV.Document d = a.Documents.Add(Application.StartupPath + "\\Template.vssx");
                    foreach (FileInfo fi in di.GetFiles("*.svg"))
                    {
                        string sf = fi.Name;
                        int i = sf.IndexOf("-service-", StringComparison.CurrentCultureIgnoreCase);
                        sf = sf.Substring(i + "-service-".Length);
                        sf = sf.Replace(".svg", "");
                        sf = sf.Replace("-", " ");

                        MOIV.Master m = d.Masters.Add();
                        m.Name = sf;
                        m.Import(fi.FullName);

                        textBox1.Text += "     " + sf + Environment.NewLine;
                        textBox1.Refresh();

                        Application.DoEvents();
                        Application.DoEvents();
                    }

                    // Save and close the stencil
                    d.SaveAs(pf + "\\" + sn + ".vssx");
                    d.Close();
                    d = null;
                    textBox1.Text += "---- Created";
                    textBox1.Refresh();

                    textBox1.Text += Environment.NewLine;
                    textBox1.Refresh();
                    Application.DoEvents();
                    Application.DoEvents();
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (a != null)
                {
                    a.Quit();
                    a = null;
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            System.Diagnostics.Process.Start(linkLabel1.Text);
        }
    }
}
