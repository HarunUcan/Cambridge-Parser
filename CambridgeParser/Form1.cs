using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace CambridgeParser
{
    public partial class Form1 : Form
    {
        string word;

        const string URL = "https://dictionary.cambridge.org/dictionary/english-turkish/";

        const string typeOfWordXpath = "/html/body/div[2]/div/div[1]/div[2]/article/div[2]/div[1]/div[2]/div[2]/div/span/div/span/div[1]/div/div";
        
        const string pronuncationXpath = "/html/body/div[2]/div/div[1]/div[2]/article/div[2]/div[1]/div[2]/div[2]/div/span/div[1]/span/div";
               
        const string translationXpath = "/html/body/div[2]/div/div[1]/div[2]/article/div[2]/div[1]/div[2]/div[2]/div/span/div[1]/div[3]/div[1]/div[2]/div/div[3]/span";
        
        const string exSentenceXpath = "/html/body/div[2]/div/div[1]/div[2]/article/div[2]/div[1]/div[2]/div[2]/div/span/div/div[3]/div[1]/div/div/div[3]/div";

        const string definitionXpath = "/html/body/div[2]/div/div[1]/div[2]/article/div[2]/div[1]/div[2]/div[2]/div/span/div/div[3]/div[1]/div[2]/div/div[2]/div";

        public string? pronunTxt;
        public string? meanTxt;
        public string? exSentenceTxt;
        public string? definitionTxt;
        public string? typeOfWordTxt;

        const string fileName = "WordsSaveFolder.txt";
        string path = Environment.CurrentDirectory + fileName;
        public Form1()
        {
            InitializeComponent();
            
        }
        public void GetDocument(string word)
        {
            HtmlWeb htmlWeb = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument htmlDocument = htmlWeb.Load(URL + word);

            HtmlNodeCollection translationNodes = htmlDocument.DocumentNode.SelectNodes(translationXpath);
            HtmlNodeCollection pronuncationNodes = htmlDocument.DocumentNode.SelectNodes(pronuncationXpath);
            HtmlNodeCollection typeOfWordNodes = htmlDocument.DocumentNode.SelectNodes(typeOfWordXpath);
            HtmlNodeCollection exSentenceNodes = htmlDocument.DocumentNode.SelectNodes(exSentenceXpath);
            HtmlNodeCollection definitionNodes = htmlDocument.DocumentNode.SelectNodes(definitionXpath);

            if (translationNodes != null)
            {
                foreach (HtmlNode translationNode in translationNodes)
                {
                    
                    label3.Text = translationNode.InnerText.Trim();
                    meanTxt = translationNode.InnerText.Trim(); 
                    
                }
                
                
                foreach (HtmlNode pronuncationNode in pronuncationNodes)
                {
                    
                    int i = 0;
                    string pron = "";
                    foreach (char c in pronuncationNode.InnerText.Trim())
                    {

                        if (c == '/' && i == 0)
                        {
                            pron += c;
                            i++;
                        }
                        else if (i == 1 && c != '/')
                        {
                            pron += c;
                        }
                        else if(c=='/' && i == 1)
                        {
                            pron += c;
                            break;
                        }

                    }
                    //Debug.WriteLine(pron);
                    //return translationNode.InnerText;
                    
                    
                    label4.Text = pron;
                    pronunTxt = pron;
                    continue;
                    
                    
                }
                
                foreach (HtmlNode typeOfWordNode in typeOfWordNodes)
                {
                    label5.Text = typeOfWordNode.InnerText.Trim();
                    typeOfWordTxt = typeOfWordNode.InnerText.Trim();
                    //Debug.WriteLine(typeOfWordNode.InnerText.Trim());
                    break;
                }
                if(exSentenceNodes != null)
                {
                    foreach (HtmlNode exSentenceNode in exSentenceNodes)
                    {
                        label6.Text = exSentenceNode.InnerText.Trim();
                        exSentenceTxt = exSentenceNode.InnerText.Trim();
                        //Debug.WriteLine(exSentenceNode.InnerText.Trim());
                        break;
                    }
                }
                else
                {
                    label6.Text = " ";
                    exSentenceTxt = " ";
                }
                
                foreach (HtmlNode definitionNode in definitionNodes)
                {
                    label7.Text = definitionNode.InnerText.Trim();
                    definitionTxt = definitionNode.InnerText.Trim();
                    //Debug.WriteLine(definitionNode.InnerText.Trim());
                    break;
                }
            }
            else
            {
                MessageBox.Show("Kelime Bulunamadý!","HATA!",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning);
            }
            
            
            //return "HATA!!";


            //Debug.WriteLine(htmlDocument.Text);
            //label2.Text = htmlDocument.Text;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            word = textBox1.Text;
            //label3.Text = GetDocument(word).Trim();
            GetDocument(word);
        }

        private void SaveData()
        {
            try
            {
                FileStream fs = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Write);
                StreamWriter sw = new StreamWriter(fs);
                
                sw.WriteLine(word + "@" + meanTxt + "@" + pronunTxt + "@" + typeOfWordTxt + "@" + exSentenceTxt + "@" + definitionTxt);                
                
                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }


        }

        private void CreateTable()
        {
            String line;
            //DataTable table = new DataTable();

            dataGridView1.Rows.Clear();


            try
            {

                StreamReader sr = new StreamReader(path);

                line = sr.ReadLine();

                while (line != null)
                {

                    string[] cells = line.Split('@');

                    /*
                    foreach (string cell in cells)
                    {
                        Debug.WriteLine(cell);
                        
                    }*/
                    for (int i = 0; i < cells.Length; i++)
                    {
                        switch (i)
                        {
                            case 0:
                                word = cells[i];
                                break;
                            case 1:
                                meanTxt = cells[i];
                                break;
                            case 2:
                                pronunTxt = cells[i];
                                break;
                            case 3:
                                typeOfWordTxt = cells[i];
                                break;
                            case 4:
                                exSentenceTxt = cells[i];
                                break;
                            case 5:
                                definitionTxt = cells[i];
                                break;
                        }

                    }
                    dataGridView1.Rows.Add(word, meanTxt, pronunTxt, typeOfWordTxt, exSentenceTxt, definitionTxt);


                    //Read the next line
                    line = sr.ReadLine();
                    Debug.Write("\n\n\n");
                }

                //close the file
                sr.Close();

            }
            catch (Exception c)
            {
                Console.WriteLine("Exception: " + c.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveData();
            CreateTable();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook wb = excel.Workbooks.Add(Missing);
            Worksheet worksheet = wb.Sheets[1];

            int col = 1;
            int row = 1;

            for(int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)worksheet.Cells[row, col + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            row++;
            for(int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Range myRange = (Range)worksheet.Cells[row+i, col + j];
                    myRange.Value2 = dataGridView1[j,i].Value == null ? "" : dataGridView1[j,i].Value;
                    myRange.Select();
                }
            }
        }
    }
}