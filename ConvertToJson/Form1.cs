using Aspose.Cells;
using CRM.Models.Global;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms; 

namespace ConvertToJson
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        bool isvalid = true;
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                
                if (label1.Text != "" && txt1rows.Text.Trim()!="")
                {
                    FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                    DialogResult result = folderBrowserDialog.ShowDialog();

                    button1.Visible = false;
                    progressBar1.Visible = true;
                    progressBar1.Value = 1;
                    // Create a file stream containing the Excel file to be opened
                    FileStream fstream = new FileStream(label1.Text, FileMode.Open);
                    button1.Visible = false;
                    progressBar1.Visible = true;
                    // Instantiate a Workbook object
                    //Opening the Excel file through the file stream
                    Workbook workbook = new Workbook(fstream);
                    Worksheet worksheet = workbook.Worksheets[0];
                    int rows = Convert.ToInt32(txt1rows.Text.Trim());
                    DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, rows+1, 5, true);
                    fstream.Close();
                    if (dataTable != null)
                    {
                        List<SheetData> list = GlobalFunctions.ConverDataTableToList<SheetData>(dataTable);
                        if (list != null)
                        {
                            List<SheetJsonData> jsondata = new List<SheetJsonData>();
                            foreach (var item in list)
                            {
                                SheetJsonData obj = new SheetJsonData();
                                obj.name = item.name;
                                obj.description = item.description;
                                obj.external_url = item.external_url;
                                obj.image = item.image;
                                List<Attribute> attributes = new List<Attribute>();
                                if (!string.IsNullOrEmpty(item.attributes))
                                {
                                    string[] data = item.attributes.Split("|");
                                    if (data != null && data.Length > 0)
                                    {
                                        foreach (var d in data)
                                        {
                                            if (!string.IsNullOrEmpty(d))
                                            {
                                                string[] typedata = d.Split(":");
                                                if (typedata != null && typedata.Length == 2)
                                                {
                                                    Attribute attribute = new Attribute();
                                                    attribute.trait_type = typedata[0].Trim();
                                                    attribute.value = typedata[1].Trim();
                                                    attributes.Add(attribute);
                                                }
                                            }
                                        }
                                    }
                                }
                                obj.attributes = attributes;
                                jsondata.Add(obj);
                            }

                            string json = Newtonsoft.Json.JsonConvert.SerializeObject(jsondata);
                            Stream myStream;
                            int count = 0;
                            if (result == DialogResult.OK)
                            {
                               
                                string folderName = folderBrowserDialog.SelectedPath;
                                int totalcount = jsondata.Where(x => x.name != null).ToList().Count;
                                foreach (var item in jsondata.Where(x=>x.name!=null).ToList())
                                {
                                    string jsond = Newtonsoft.Json.JsonConvert.SerializeObject(item);

                                    string path = folderName+"/"+item.name+".json";

                                    // This text is added only once to the file.
                                    if (!File.Exists(path))
                                    {
                                        using (StreamWriter sw = File.CreateText(path))
                                        {
                                            sw.WriteLine(jsond);
                                        }
                                    }
                                    else
                                    {
                                        using (StreamWriter sw = File.AppendText(path))
                                        {
                                            sw.WriteLine(jsond);
                                        }
                                    }
                                    count++;
                                    int pbv = count / totalcount * 100;
                                    if (pbv < 100)
                                    {
                                        progressBar1.Value = pbv;
                                    }
                                    else
                                    {
                                        progressBar1.Value = 100;
                                    }
                                    
                                }
                                
                                MessageBox.Show("File saved successfully");
                                button1.Visible = true;
                                progressBar1.Visible = false;


                            }
                        }
                    }
                }
                else
                {   
                    if(label1.Text == "")
                    {
                        MessageBox.Show("Please select file");
                    }
                    if (txt1rows.Text.Trim() == "")
                    {
                        MessageBox.Show("Please enter no of rows to be convert");
                    }
                }
            }
            catch (Exception ex)
            {
                button1.Visible = true;
                isvalid = false;
                MessageBox.Show("File not saved: "+ex.Message);
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            isvalid = false;
            //To where your opendialog box get starting location. My initial directory location is desktop.
            openFileDialog1.InitialDirectory = "C://Desktop";
            //Your opendialog box title name.
            openFileDialog1.Title = "Select file to be upload.";
            //which type file format you want to upload in database. just add them.
            openFileDialog1.Filter = "Select Valid Document(*.xlsx;)|*.xlsx;";
            //FilterIndex property represents the index of the filter currently selected in the file dialog box.
            openFileDialog1.FilterIndex = 1;
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog1.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(openFileDialog1.FileName);
                        label1.Text = path;
                        isvalid = true;
                    }
                }
                else
                {
                    MessageBox.Show("Please Upload document.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
