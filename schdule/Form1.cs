using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Code7248;

namespace schdule
{
    public partial class Form1 : Form
    {
        #region defiend_vraibles

        List<mydata> alldata = new List<mydata>();
        int index = 0;
        List<data> unediteddata = new List<data>();
        bool SaveWorkTimeOnClosing = false;
        int rowsbypage = 25;
        int cnum = 0;
        string text = "";
        string filename = "";
        #endregion

        #region defiend clsses
        class mydata
        {
          public  string date,text,
             text_type,
             text_classfication,
             subject,
             titel,
             auther,
             tages,
             notes,id;

            public mydata(string id, string date,string text,string text_type,string text_classification,string subject,string titel,string auther,string tages,string notes)
            {
                this.id = id;
                this.date = date; this.text = text; this.text_type = text_type; this.text_classfication = text_classification;
                this.subject = subject; this.titel = titel; this.auther = auther; this.tages = tages; this.notes = notes;
            }
        }
        class data
        {
            public string text, time;
            public data(string text,string time)
            {
                this.text = text;
                this.time = time;
            }
           
        }
        #endregion
        public Form1()
        {
            InitializeComponent();
           
            this.AllowDrop = true;
            
        }

        
      

        //DataGridViewComboBoxCell v = new DataGridViewComboBoxCell();


        bool exiest(string d)
        {
            for (int i = 0; i < unediteddata.Count; i++)
            {
                if (unediteddata[i].text == d)
                    return true;
            }
            return false;
        }

        void write()
        {
            timer1.Stop();
            SaveWorkTimeOnClosing = false;
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet_mkalat = null ;
            Microsoft.Office.Interop.Excel._Worksheet oSheet_makolat;
            Microsoft.Office.Interop.Excel._Worksheet oSheet_kabsolat;
            Microsoft.Office.Interop.Excel._Worksheet oSheet_moade3_5asa;
           
           
            //Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;
                
                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet_mkalat = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                //==========================مقالات====================
                oSheet_mkalat.Name = "مقالات";
                //date class titel text tages nootes
                //Add table headers going cell by cell.
                oSheet_mkalat.Cells[1, 1] = "التاريخ";
                oSheet_mkalat.Cells[1, 2] = "التصنيف";
                oSheet_mkalat.Cells[1, 3] = "العنوان";
                oSheet_mkalat.Cells[1, 4] = "نص المقال";
                oSheet_mkalat.Cells[1, 5] = "الوسوم";
                oSheet_mkalat.Cells[1, 6] = "ملاحظات";
                //Format A1:D1 as bold, vertical alignment = center.
                oSheet_mkalat.get_Range("A1", "F1").Font.Bold = true;
                oSheet_mkalat.get_Range("A1", "F1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                oSheet_mkalat.get_Range("A1", "F1").Interior.Color = Color.LightBlue;
                
             
                int m = 2;
                foreach (var item in alldata)
                {

                    if (item.text_type == "مقالات") {

                        oSheet_mkalat.Cells[m, 1] = item.date;
                        oSheet_mkalat.Cells[m, 2] = item.text_classfication;
                        oSheet_mkalat.Cells[m, 3] = item.titel;
                        oSheet_mkalat.Cells[m, 4] = item.text;
                        oSheet_mkalat.Cells[m, 5] = item.tages;
                        oSheet_mkalat.Cells[m++, 6] = item.notes;
                    }
                }
                //===========================مقولات===============================
                oSheet_makolat = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets.Add(After:oWB.Sheets[oWB.Sheets.Count]);
                oSheet_makolat.Cells[1, 1] = "التاريخ";
                oSheet_makolat.Cells[1, 2] = "التصنيف";
                oSheet_makolat.Cells[1, 3] = "المقولة";
                oSheet_makolat.Cells[1, 4] = "القائل";
                oSheet_makolat.Cells[1, 5] = "الملاحظات";
                oSheet_makolat.Name = "مقولات";

                oSheet_makolat.get_Range("A1", "E1").Font.Bold = true;
                oSheet_makolat.get_Range("A1", "E1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                oSheet_makolat.get_Range("A1", "E1").Interior.Color = Color.LightBlue;
                
                int m1 = 2;
                foreach (var item in alldata)
                {
                    
                    if (item.text_type == "مقولات")
                    {
                        
                        oSheet_makolat.Cells[m1, 1] = item.date;
                        oSheet_makolat.Cells[m1, 2] = item.text_classfication;
                        oSheet_makolat.Cells[m1, 3] = item.text;
                        oSheet_makolat.Cells[m1, 4] = item.auther;
                        oSheet_makolat.Cells[m1++, 5] = item.notes;
                    }
                }
                

                //===================كبسولات==================
                oSheet_kabsolat = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets.Add(After: oWB.Sheets[oWB.Sheets.Count]); ;
                oSheet_kabsolat.Name = "كبسولات";
                oSheet_kabsolat.Cells[1, 1] = "التاريخ";
                oSheet_kabsolat.Cells[1, 2] = "التصنيف";
                oSheet_kabsolat.Cells[1, 3] = " الكبسولة";
                oSheet_kabsolat.Cells[1, 4] = "الوسوم";
                oSheet_kabsolat.Cells[1, 5] = "ملاحظات";

                oSheet_kabsolat.get_Range("A1", "E1").Font.Bold = true;
                oSheet_kabsolat.get_Range("A1", "E1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                oSheet_kabsolat.get_Range("A1", "E1").Interior.Color = Color.LightBlue;
                

                
                int m11 = 2;
                foreach (var item in alldata)
                {
                    
                    if (item.text_type == "كبسولات")
                    {
                        oSheet_kabsolat.Cells[m11, 1] = item.date;
                        oSheet_kabsolat.Cells[m11, 2] = item.text_classfication;
                        oSheet_kabsolat.Cells[m11, 3] = item.text;
                        oSheet_kabsolat.Cells[m11, 4] = item.tages;
                        oSheet_kabsolat.Cells[m11++, 5] = item.notes;
                    }
                }
                //==============================مواضيع خاصة===================

                oSheet_moade3_5asa = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets.Add(After: oWB.Sheets[oWB.Sheets.Count]); 
                //==========================مقالات====================
                oSheet_moade3_5asa.Name = "مواضيع خاصة";
                //date class titel text tages nootes
                //Add table headers going cell by cell.
                oSheet_moade3_5asa.Cells[1, 1] = "التاريخ";
                oSheet_moade3_5asa.Cells[1, 2] = "التصنيف";
                oSheet_moade3_5asa.Cells[1, 3] = "الموضوع";
                oSheet_moade3_5asa.Cells[1, 4] = "المقولة";
                oSheet_moade3_5asa.Cells[1, 5] = "الوسوم";
                oSheet_moade3_5asa.Cells[1, 6] = "ملاحظات";
                //Format A1:D1 as bold, vertical alignment = center.

                oSheet_moade3_5asa.get_Range("A1", "F1").Font.Bold = true;
                oSheet_moade3_5asa.get_Range("A1", "F1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                oSheet_moade3_5asa.get_Range("A1", "F1").Interior.Color = Color.LightBlue;

                int m3 = 2;
                foreach (var item in alldata)
                {

                    if (item.text_type == "مواضيع خاصة")
                    {

                        oSheet_moade3_5asa.Cells[m3, 1] = item.date;
                        oSheet_moade3_5asa.Cells[m3, 2] = item.text_classfication;
                        oSheet_moade3_5asa.Cells[m3, 3] = item.subject;
                        oSheet_moade3_5asa.Cells[m3, 4] = item.text;
                        oSheet_moade3_5asa.Cells[m3, 5] = item.tages;
                        oSheet_moade3_5asa.Cells[m3++, 6] = item.notes;
                    }
                }
                
                
                /*
                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";

                saNames[4, 1] = "Johnson";
                
                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet_mkalat.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet_mkalat.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet_mkalat.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet_mkalat.get_Range("A1", "D1");
                */
                 // oRng.EntireColumn.AutoFit();

                //oXL.Visible = false;
                //oXL.UserControl = false;

                string excelname = Directory.GetCurrentDirectory()+"\\Excel\\"+filename.Trim()+".xlsx";
        oWB.SaveAs(excelname, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
         false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
         Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //oWB.Close();

                //...
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        void removeandsplit()
        {
            string[] mon = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            /*
            




Saturday,
Sunday,
Monday,
Tuesday,
Wednesday,
Thursday,
Friday,
            */
            string[] days = { "Saturday,", "Sunday,", "Monday,", "Tuesday,", "Wednesday,", "Thursday,", "Friday," };
            bool fromPage = false;
            bool fromfacbookdirect = false;
            string[] tokens = null;
            if (!text.Contains("Permalink"))
                fromPage = true;
            if (text.Contains("UTC"))
                fromfacbookdirect = true;
            groupBox1.Hide();
            groupBox4.Show();
            groupBox3.Show();
            groupBox2.Show();
            this.WindowState = FormWindowState.Maximized;

            text = Regex.Replace(text, @"http[^\s]+", string.Empty);
            text = Regex.Replace(text, @"\b[0-9]*\s*(likes?|comments?)\s*[0-9]*\b", "").Replace("<", string.Empty);

            text = Regex.Replace(text, @"[Vv]ia\s*[a-zA-Z_ ]+", string.Empty);
            text = Regex.Replace(text, @"\b[wW]{3}.[a-zA-Z]+.[a-zA-Z]+\b", "");
            text = Regex.Replace(text, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);
            if (!fromPage && !fromfacbookdirect)
                tokens = Regex.Split(text, "Permalink");
            else if (!fromfacbookdirect)
            {

                foreach (string mons in mon)
                {
                    text = text.Replace(mons.ToUpper(), mons.ToUpper() + " ##$$##$$ ");
                }
                tokens = text.Split(new string[] { "##$$##$$" }, StringSplitOptions.RemoveEmptyEntries);
            }
            else {
                tokens = text.Split(days,StringSplitOptions.RemoveEmptyEntries);
            }
            //    MessageBox.Show(tokens.Length.ToString());
            int ii = 0;
            List<string> dd = new List<string>();
           
            foreach (string item in tokens)
            {
                

                if (!fromPage||fromfacbookdirect)
                {

                    string tex = item;
                    if (!tex.Contains("UTC"))
                    {
                        
                        //وائل عزيز - WAEL AZIZ
                        string[] t = tex.Split(new string[] { "Wael Aziz" }, StringSplitOptions.RemoveEmptyEntries);
                        if (t.Length >= 2)
                        {
                            t[1] = Regex.Replace(t[1], "\b[a-zA-Z]+.[a-zA-Z]{3}\b", string.Empty).Trim();

                            if (t[0].Contains("وائل عزيز - "))
                                t[0] = t[0].Replace("وائل عزيز - ", string.Empty);

                            if (!t[0].Contains(',') && t[0].Trim().Length >= 9)
                            {
                                string year = t[0].Trim().Substring(t[0].Trim().Length - 4);
                                string fd = t[0].Trim().Substring(0, t[0].Trim().Length - 4);
                                t[0] = fd + ", " + year;

                            }
                        }
                        if (t.Length >= 3 && (Regex.IsMatch(t[1], @"[A-Za-z]{3} [0-9]{1,2},?\s?[0-9]{4}")))
                        {
                            string n = Regex.Matches(t[1], @"[A-Za-z]{3} [0-9]{1,2},?\s?[0-9]{4}")[0].Value;
                            if (!exiest(t[2].Trim()))
                                unediteddata.Add(new data(t[2].Trim(), n.Trim()));
                        }
                        else
                            if (t.Length >= 2 && t[1].Trim() != "")
                            if (!exiest(t[1].Trim()))
                                unediteddata.Add(new data(t[1].Trim(), t[0].Trim()));
                    }
                    else {
                        //at 7:27am UTC+02
                        string[] splites = Regex.Split(tex, @"at\s[0-9]+:[0-9]+[ap]m\sUTC[+-][0-9]+");
                        if (splites.Length > 1)
                        {
                            if (!exiest(splites[1].Trim()))
                                unediteddata.Add(new data(splites[1].Trim(), splites[0].Trim()));
                        }
                    }
                }
                else
                {

                    string dat="";
                    string myitem = "";
                    if(ii>0)
                    myitem = Regex.Replace(item, @"\b[0-9]{1,2}\s[A-Z]+\b", string.Empty);
                    if (Regex.IsMatch(item, @"\b[0-9]{1,2}\s[A-Z]+\b"))
                    {
                        dat = Regex.Matches(item, @"\b[0-9]{1,2}\s[A-Z]+\b")[0].Value;
                        dd.Add(dat);

                    }
                    if (ii > 0) {
                        string postdat = dd[ii - 1];

                        myitem = Regex.Replace(myitem, "وائل عزيز - ", string.Empty).Replace("updated their status.", string.Empty);
                        string[] posts = myitem.Split(new string[] { "Wael Aziz" }, StringSplitOptions.RemoveEmptyEntries);
                        string fil = filename.Replace("and", "");
                        string ss=fil.Substring(fil.Length-4);
                        
                        foreach (var post in posts)
                        {
                            if (post.Length > 10)
                            {
                                string pp = post;
                               pp= Regex.Replace(pp, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);
                                unediteddata.Add(new data(pp, postdat + " 20" + ss));
                            }
                        }
                    }
                }
                ii++;
            }
            for (int i = index; i < unediteddata.Count; i++)
            {

                if (index < rowsbypage)
                    dataGridView1.Rows.Add(new object[] { ++index, unediteddata[i].time, unediteddata[i].text,"مواضيع خاصة","الشأن العام","سياسي" });   
                else
                    break;
            }
            
            dataGridView1.Enabled = false;
            lblpnum.Text = ((unediteddata.Count + cnum) % rowsbypage == 0) ? (((unediteddata.Count + cnum) / rowsbypage).ToString()) : (((((unediteddata.Count + cnum) / rowsbypage) + 1).ToString()));
            lblpremain.Text = ((unediteddata.Count + cnum - alldata.Count) % rowsbypage == 0) ? (((unediteddata.Count + cnum - alldata.Count) / rowsbypage).ToString()) : (((((unediteddata.Count + cnum - alldata.Count) / rowsbypage) + 1).ToString()));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Docx|*.docx;*.doc";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                savefilename(Path.GetFileNameWithoutExtension(openFileDialog1.FileName));
                text = new Code7248.word_reader.TextExtractor(openFileDialog1.FileName).ExtractText(); 
            
            }
            if (text.Length >= 1)
                removeandsplit();
            else
            {
                MessageBox.Show("برجاء تحميل الملف");
            }
            
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            string[] authers = File.ReadAllLines(@"data\auther.dat");
            aut.AddRange(authers);
            label4.Parent = groupBox1;
            int fx = this.Size.Width;
            button1.Location = new Point(fx/2 - (button1.Size.Width / 2), button1.Location.Y);
            lbldrage.Location = new Point(fx/2 - (lbldrage.Size.Width / 2), lbldrage.Location.Y+25);
            label4.Location = new Point(fx / 2 - (label4.Size.Width / 2), lbldrage.Location.Y - 125);
            groupBox1.Show();
            groupBox2.Hide(); groupBox3.Hide(); groupBox4.Hide(); 
           
            if (File.Exists(@"data\info.dat"))
                filename = File.ReadAllText(@"data\info.dat");
            groupBox4.Hide();
            if (File.Exists(@"data\worktime.dat"))
            {
                string[] times = File.ReadAllText(@"data\worktime.dat").Split(';');
                hh = int.Parse(times[0]);
                mm = int.Parse(times[1]);
                ss = int.Parse(times[2]);
            }
            if (!File.Exists(@"data\savfin.dat"))
                return;
            string[] editedsaveddata = File.ReadAllText(@"data\savfin.dat").Split(new string[] { "@#+=" },StringSplitOptions.RemoveEmptyEntries);
            groupBox2.Show();
            groupBox3.Show();
            dataGridView1.Enabled = false;
            cnum = editedsaveddata.Length;
            if (editedsaveddata.Length >= 1)
            {
                foreach (var item in editedsaveddata)
                {
                    string[] tokens = item.Split(new string[] { "$&;" },StringSplitOptions.None);
                    mydata dat = new mydata(tokens[0], tokens[1], tokens[2], tokens[3], tokens[4], tokens[5], tokens[6], tokens[7], tokens[8], tokens[9]);       
                    alldata.Add(dat);
                }
            }
            int ind = cnum;
            int dis = rowsbypage;
            string[] uneditedsaveddata = File.ReadAllText(@"data\sav.dat").Split(new string[] { "@#+=" },StringSplitOptions.RemoveEmptyEntries);
            
            if (uneditedsaveddata.Length >= 1)
            {
                foreach (var item in uneditedsaveddata)
                {
                    string[] tokens = item.Split(new string[] { "$&;" }, StringSplitOptions.None);
                    unediteddata.Add(new data(tokens[1], tokens[0]));
                    if (dis-- >= 1)
                    {
                      
                        dataGridView1.Rows.Add(new object[] { ++ind, tokens[0], tokens[1],"مواضيع خاصة","الشأن العام" ,"سياسي"});
                        index++;        
                    }
                }
                groupBox1.Hide();
                groupBox4.Show();
                this.WindowState = FormWindowState.Maximized;
                
                lblpnum.Text = ((unediteddata.Count + cnum) % rowsbypage == 0) ? (((unediteddata.Count + cnum) / rowsbypage).ToString()) : (((((unediteddata.Count + cnum) / rowsbypage) + 1).ToString()));
                lblpremain.Text = ((unediteddata.Count + cnum - alldata.Count) % rowsbypage == 0) ? (((unediteddata.Count + cnum - alldata.Count) / rowsbypage).ToString()) : (((((unediteddata.Count + cnum - alldata.Count )/ rowsbypage) + 1).ToString()));
            }
        }
        AutoCompleteStringCollection aut = new AutoCompleteStringCollection();
        
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 7)
            {
               
                (e.Control as TextBox).Multiline = false;
                (e.Control as TextBox).AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                (e.Control as TextBox).AutoCompleteSource = AutoCompleteSource.CustomSource;
                
                (e.Control as TextBox).AutoCompleteCustomSource = aut;
                
            }


            if (e.Control is ComboBox)
            {
                // remove handler first to avoid attaching twice
                ((ComboBox)e.Control).SelectedIndexChanged -= Form1_SelectedIndexChanged;
                ((ComboBox)e.Control).SelectedIndexChanged += Form1_SelectedIndexChanged;
            }
            
        }
        
        void Form1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string type = (sender as ComboBox).SelectedItem.ToString();
           // dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
         //   dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            switch (type)
            {
                case "مقالات":

                    dataGridView1.Columns[5].Visible=false;       
                    dataGridView1.Columns[7].Visible = false;
                    dataGridView1.Columns[6].Visible = true;
                    dataGridView1.Columns[8].Visible = true;
                    
                    break;
                case "مقولات":
                    
                       
                        dataGridView1.Columns[5].Visible = false;
                        dataGridView1.Columns[7].Visible = true;
                        dataGridView1.Columns[6].Visible = false;
                        dataGridView1.Columns[8].Visible = false;
                    
                    ; break;
                case "كبسولات":
                   
                    
                        dataGridView1.Columns[5].Visible = false;
                        dataGridView1.Columns[7].Visible = false;
                        dataGridView1.Columns[6].Visible = false;
                        dataGridView1.Columns[8].Visible = true;
                   
                    ; break;
                case "مواضيع خاصة":
                   
                        dataGridView1.Columns[5].Visible = true;
                        dataGridView1.Columns[7].Visible = false;
                        dataGridView1.Columns[6].Visible = false;
                        dataGridView1.Columns[8].Visible = true;
                   
                    break;
                default:
                    break;
            }
     
           
            
  
     
        }

        bool Savedata() {
            
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                string id = dataGridView1[0, j].Value.ToString();
                string date = (dataGridView1[1, j].Value == null) ? "" : dataGridView1[1, j].Value.ToString();
             
                string text = (dataGridView1[2, j].Value == null) ? "" : dataGridView1[2, j].Value.ToString();
                if (dataGridView1[3, j].Value == null) {
                    dataGridView1.Columns[3].Visible = true;
                    dataGridView1.CurrentCell = dataGridView1[3, j];
                    MessageBox.Show(this,"لا بد من تحديد نوع النص","نوع النص مفقود",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    return false;
                }
                string text_type = (dataGridView1[3, j].Value == null) ? "" : dataGridView1[3, j].Value.ToString();
            //    MessageBox.Show(text_type);
                if (dataGridView1[4, j].Value == null)
                {
                    dataGridView1.Columns[4].Visible = true;
                    dataGridView1.CurrentCell = dataGridView1[4, j];
                     
                    MessageBox.Show(this, "لا بد من تحديد تصنيف النص", "تصنيف النص مفقود", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                string text_classification = (dataGridView1[4, j].Value == null) ? "" : dataGridView1[4, j].Value.ToString();
              //  MessageBox.Show(text_classification);

                string ttype = "";
                    ttype=dataGridView1[3,j].Value.ToString();
                if (ttype=="مواضيع خاصة"&& dataGridView1[5, j].Value == null)
                {
                    dataGridView1.Columns[5].Visible = true;
                    dataGridView1.CurrentCell = dataGridView1[5, j];

                    MessageBox.Show(this, "لا بد من تحديد نوع الموضوع", "نوع الموضوع مفقود", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        return false;
                }
                string subject = (dataGridView1[5, j].Value == null) ? "" : dataGridView1[5, j].Value.ToString();
               // MessageBox.Show(subject);
                string titel = (dataGridView1[6, j].Value == null) ? "" : dataGridView1[6, j].Value.ToString();
                string ttype2 ="";
                   ttype2 = dataGridView1[3, j].Value.ToString();
                if (ttype2=="مقولات"&& dataGridView1[7, j].Value == null)
                {
                    dataGridView1.Columns[7].Visible = true;
                    dataGridView1.CurrentCell = dataGridView1[7, j];
                    
                 DialogResult d=   MessageBox.Show(this, "هل يوجد قائل لهذه المقولة ؟\n"+dataGridView1[2,j].Value.ToString(), "القائل مفقود",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                    
                    
                    if (d == DialogResult.Yes)
                    {

                        return false;
                    }
                    else
                        dataGridView1.Rows[j].Selected = false;
                }
                string auther = (dataGridView1[7, j].Value == null) ? "" : dataGridView1[7, j].Value.ToString();
                if (auther != "" && !aut.Contains(auther))
                {
                    string d = File.ReadAllText(@"data\auther.dat");
                    
                    File.WriteAllText(@"data\auther.dat", d+"\n" + auther);
                    aut.Add(auther);
                }
                string tages = (dataGridView1[8, j].Value == null) ? "" : dataGridView1[8, j].Value.ToString();
                string notes = (dataGridView1[9, j].Value == null) ? "" : dataGridView1[9, j].Value.ToString();
             if(!inalldata(text))
                alldata.Add(new mydata(id,date, text, text_type, text_classification, subject, titel, auther, tages, notes));
            }
            return true;
        }
        bool inalldata(string text)
        {
            foreach (mydata item in alldata)
            {
                if (text == item.text)
                    return true;
            }
            return false;
        }
        private void button2_Click(object sender, EventArgs e)
        {
           
            if (!savedatatodisk())
                return;
            //write();
            int c = 0;
            dataGridView1.Rows.Clear();
            if (index >= unediteddata.Count)
            {
                MessageBox.Show("لقد انتهيت من تصنيف جميع النصوص بالملف الذى تم تحميله \n  سيتم انشاء ملف الاكسل الآن");
                timer1.Stop();
                creatreport();
                SaveWorkTimeOnClosing = false;   
                write();
                deletefiles();
                Close();
                return;
            }
                for (int i = index; i <unediteddata.Count ; i++)
            {
                if (c++ <= rowsbypage)
                {
                    index++;
                    

                    dataGridView1.Rows.Add(new object[] { i + cnum + 1, unediteddata[i].time, unediteddata[i].text, "مواضيع خاصة", "الشأن العام", "سياسي" });
                }
                else
                    break;
            }

                lblpnum.Text = ((unediteddata.Count + cnum) % rowsbypage == 0) ? (((unediteddata.Count + cnum) / rowsbypage).ToString()) : (((((unediteddata.Count + cnum) / rowsbypage) + 1).ToString()));
                lblpremain.Text = ((unediteddata.Count + cnum - alldata.Count) % rowsbypage == 0) ? (((unediteddata.Count + cnum - alldata.Count) / rowsbypage).ToString()) : (((((unediteddata.Count + cnum - alldata.Count) / rowsbypage) + 1).ToString()));

        }
        void creatreport()
        {
            StreamWriter writreport = new StreamWriter(@"data\report.txt");
            writreport.WriteLine("عدد ساعات العمل الكلية");
            writreport.WriteLine(hh+" hours");
            writreport.WriteLine(mm+" min");
         
            writreport.Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (!Savedata())
                return;
            if (alldata.Count < unediteddata.Count)
            {
                MessageBox.Show("باقي "+(unediteddata.Count-alldata.Count)+"نص لم يتم تصنيفهم من فضلك اكمل تصنيف الملف كامل\nأو اضغط  حفظ البيانات لاستكمال في وقت لاحق");
                return;
            }

            
            creatreport();
            write();
            deletefiles();

            Close();
            
        }

        void deletefiles()
        {
           
            if (File.Exists(@"data\sav.dat"))       
                File.Delete(@"data\sav.dat");
            if (File.Exists(@"data\savfin.dat"))
            {
                string na = filename;
                if (File.Exists(@"staticdata\" + na + ".dat"))
                {
                    na += new Random().Next(1, 10000).ToString();
                   
                }
                File.WriteAllLines(@"staticdata\"+na+".dat", File.ReadAllLines(@"data\savfin.dat"));
             
                File.Delete(@"data\savfin.dat");
            }
            if (File.Exists(@"data\worktime.dat"))
                File.Delete(@"data\worktime.dat");
            if (File.Exists(@"data\info.dat"))
                File.Delete(@"data\info.dat");
        
        }
        void tembrorsave()
        {
            
            //================edited data=============
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1[3, i].Value == null || dataGridView1[4, i].Value == null)
                    break;
                string b=(dataGridView1[3,i].Value==null)?"":dataGridView1[3,i].Value.ToString();
                if (b == "مواضيع خاصة" && dataGridView1[5, i] == null)
                    break;

               
                if (b == "مقولات" && dataGridView1[7, i] == null)
                    break;
                string id = dataGridView1[0, i].Value.ToString();
                string date = dataGridView1[1, i].Value.ToString();
                string text = dataGridView1[2, i].Value.ToString();
               
                  
                string text_type =dataGridView1[3, i].Value.ToString();
                //    MessageBox.Show(text_type);
                
                string text_classification = dataGridView1[4, i].Value.ToString();
                //  MessageBox.Show(text_classification);

                
                string subject =(dataGridView1[5, i].Value==null)?"":dataGridView1[5, i].Value.ToString();
                // MessageBox.Show(subject);
                string titel = (dataGridView1[6, i].Value==null)?"":dataGridView1[6, i].Value.ToString();

                string auther = (dataGridView1[7, i].Value==null) ? "" : dataGridView1[7, i].Value.ToString();
                string tages = (dataGridView1[8, i].Value==null) ? "" : dataGridView1[8, i].Value.ToString();
                string notes = (dataGridView1[9, i].Value==null) ? "" : dataGridView1[9, i].Value.ToString();
                alldata.Add(new mydata(id, date, text, text_type, text_classification, subject, titel, auther, tages, notes));
                
            }
            StreamWriter w = new StreamWriter(@"data\savfin.dat");

            
            foreach (var item in alldata)
            {
                w.Write(item.id + "$&;" + item.date + "$&;" + item.text + "$&;" + item.text_type + "$&;" + item.text_classfication + "$&;" + item.subject + "$&;" + item.titel + "$&;" + item.auther + "$&;" + item.tages + "$&;" + item.notes + "@#+=");
            }
            w.Close();
            //================================= not edited yet===========
            StreamWriter w1 = new StreamWriter(@"data\sav.dat");
            
            for (int i = alldata.Count-cnum; i < unediteddata.Count; i++)
            {
                w1.Write(unediteddata[i].time + "$&;" + unediteddata[i].text + "@#+=");
            }
            w1.Close();
            MessageBox.Show("Data Saved");
        
        }
        bool savedatatodisk()
        {
            if (!Savedata())
                return false;
            //================edited data=============
            StreamWriter w = new StreamWriter(@"data\savfin.dat");

            foreach (var item in alldata)
            {
                w.Write(item.id + "$&;" + item.date + "$&;" + item.text + "$&;" + item.text_type + "$&;" + item.text_classfication + "$&;" + item.subject + "$&;" + item.titel + "$&;" + item.auther + "$&;" + item.tages + "$&;" + item.notes + "@#+=");
            }
            w.Close();
            //================================= not edited yet===========
            StreamWriter w1 = new StreamWriter(@"data\sav.dat");

            for (int i = alldata.Count-cnum; i < unediteddata.Count; i++)
            {
                w1.Write(unediteddata[i].time+"$&;"+unediteddata[i].text+"@#+=");
            }
            w1.Close();
            
            return true;
        }
        private void button4_Click(object sender, EventArgs e)
        {

            savedatatodisk();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SaveWorkTimeOnClosing)
            {
                StreamWriter worktime = new StreamWriter(@"data\worktime.dat");
                worktime.WriteLine(hh + ";" + mm + ";" + ss);
                worktime.Close();
            }
            if(unediteddata.Count>=1&&SaveWorkTimeOnClosing)
            if (MessageBox.Show("حفظ البيانات", "الحفظ",
        MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.Yes)
            {
                // Cancel the Closing event from closing the form.
               
                tembrorsave();
               
                // Call method to save file...
               
            }
        }
        int ss = 0;
        int mm = 0;
        int hh = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            ss++; 
            if (ss == 60)
            {
                ss = 0;
                mm++;
            }
            if (mm == 60)
            {
                mm = 0;
                hh++;
 
            }
            lbltiimess.Text = ss.ToString();
            lbltimemen.Text = mm.ToString();
            lblworktime.Text = hh.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SaveWorkTimeOnClosing = true;
            if (button5.Text == "بدأ العمل")
            {
                timer1.Start();
                button5.Text = "توقف العمل";
                dataGridView1.Enabled = true;
            }
            else
            {
                button5.Text = "بدأ العمل";
                timer1.Stop();
                dataGridView1.Enabled = false;
            
            }
        }

        private void نصوصToolStripMenuItem_Click(object sender, EventArgs e)
        {
            نصToolStripMenuItem.Checked = نصToolStripMenuItem1.Checked = نصToolStripMenuItem3.Checked = نصوصToolStripMenuItem.Checked = نصوصToolStripMenuItem1.Checked = false;
            (sender as ToolStripMenuItem).Checked = true;
            int t= int.Parse( (sender as ToolStripMenuItem).Tag.ToString());
            rowsbypage = t;
        }

        void savefilename(string filename)
        {
            this.filename = filename;
            StreamWriter fnam = new StreamWriter(@"data\info.dat");
            fnam.Write(filename);
            fnam.Close();
            
        }
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {   string[] fileNames = null;
        StringBuilder stn = new StringBuilder();
        StringBuilder path = new StringBuilder();
            try
            {
                fileNames = e.Data.GetData(DataFormats.FileDrop) as string[];
                
                for (int i = 0; i < fileNames.Length; i++)
                {
                   stn.Append( new Code7248.word_reader.TextExtractor(fileNames[i]).ExtractText());
                   path.Append(Path.GetFileNameWithoutExtension(fileNames[i]));
                   path.Append(" and ");

                       
                }
              
                text = stn.ToString();
                File.WriteAllText(@"text.txt", text);
              
            }
               
            catch (Exception E)
            {
                MessageBox.Show(E.Message);
             
            }
            if (text.Length >= 1)
            {
                savefilename(path.ToString());
                removeandsplit();
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;

            

        }

        void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void lblpnum_Click(object sender, EventArgs e)
        {

        }

       
    }
}