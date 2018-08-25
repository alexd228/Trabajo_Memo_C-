using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Printing;
//using CrystalReportsCommonObjectModelLib;
using System.Windows;
using System.Resources;
using System.Timers;
using Microsoft.Office.Interop.Excel;

namespace Trabajodememo
{
    public partial class Form1 : Form
    {
        public class db
        {
            public double counter;
            public double count;
            public int batch;

            public string name;
            public string lastname;
            public string address;
            public double phonenumber;
            public double age;
            public string city;
            public string state;
            public string company;
            public string resultado;
            public string resultadoalcohol;
            public string genero;
            public string DER;
            public double SS;
            public string License;
            public string datecreated;
            public string razon;
            public double ID;
            public double money;
            public string status;
            public double dotnumber;
            public string DOB;
            

            public Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            // xlApp.Visible = true;
            public Excel.Workbook xlWorkBook;
            public Excel.Worksheet xlWorkSheet;
            //for printing we need to work on two sheets
            public Excel.Application xlApp2 = new Microsoft.Office.Interop.Excel.Application();
            public Excel.Workbook xlWorkBook2;
            public Excel.Worksheet xlWorkSheet2;
            //for generating list 3 sheets
            public Excel.Application xlApp3 = new Microsoft.Office.Interop.Excel.Application();
            public Excel.Workbook xlWorkBook3;
            public Excel.Worksheet xlWorkSheet3;
        }
       
        object misValue = System.Reflection.Missing.Value;
        List<string> list = new List<string>();
        List<db> objarray = new List<db>();
        FileInfo fi = new FileInfo(File.ReadLines(@"C:\MemoExcelfolder\PathForDataBase.txt").Take(1).First());
        string fi2 = File.ReadLines(@"C:\MemoExcelfolder\PathForDataBase.txt").Take(1).First();
        string fi3 = File.ReadLines(@"C:\MemoExcelfolder\PathForDataBase.txt").Skip(1).Take(1).First();
        string fi4 = File.ReadLines(@"C:\MemoExcelfolder\PathForDataBase.txt").Skip(2).Take(1).First();
        string fi5photo = File.ReadLines(@"C:\MemoExcelfolder\PathForDataBase.txt").Skip(3).Take(1).First();
        string fi6startphoto = File.ReadLines(@"C:\MemoExcelfolder\PathForDataBase.txt").Skip(4).Take(1).First();
        string fi7 = File.ReadLines(@"C:\MemoExcelfolder\PathForDataBase.txt").Skip(5).Take(1).First();
        db D = new db();
                Random rAnD = new Random();
                List<string> arr = new List<string>();
                List<string> arr2 = new List<string>();
                List<string> temp = new List<string>();
                List<DateTime> allDates = new List<DateTime>();

                string[] whatever2 = new string[20000];
                string companyforprint;
                string[] RandDrugsAlcohol = new string[10] {"Drugs","Alcohol","Drugs & Alcohol","Drugs","Drugs","Drugs","Drugs","Drugs & Alcohol","Drugs & Alcohol","Drugs & Alcohol"};
                Random drugalco = new Random();
                int totaldrugs = 0;
                int totalalcohol = 0;
             //   int totaldrugsalcohol = 0;
        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {
          //  this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MinimumSize = new Size(674, 674);
            this.MaximumSize = new Size(1744,674);
            
            this.MaximizeBox = false;
            if (fi.Exists)
            {
                // D.xlWorkBook = D.xlApp.Workbooks.Open(@"C:\Users\Adrian\Desktop\MemoExcel2.xls", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                // D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                //D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                //D.count = D.counter;
            }
            else
            {
                
                MessageBox.Show("MemoExcel was not found so a new Excel document will be created");
                try
                {

                    string root = fi4;
                    if (!Directory.Exists(root))
                    {
                        Directory.CreateDirectory(root);
                    }

                    D.xlWorkBook = D.xlApp.Workbooks.Add(misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.xlWorkSheet.Cells[1, 11] = 1;
                    D.xlWorkSheet.Cells[1, 1] = "Nombre"; D.xlWorkSheet.Cells[1, 2] = "Apellido";
                    D.xlWorkSheet.Cells[1, 3] = "Edad"; D.xlWorkSheet.Cells[1, 4] = "Telefono";
                    D.xlWorkSheet.Cells[1, 5] = "Direccion"; D.xlWorkSheet.Cells[1, 6] = "Ciudad";
                    D.xlWorkSheet.Cells[1, 7] = "Estado"; D.xlWorkSheet.Cells[1, 8] = "Compañia";
                    D.xlWorkSheet.Cells[1, 9] = "Resultado-Drogas"; D.xlWorkSheet.Cells[1, 10] = "Resultado-Alcohol";
                    D.xlWorkSheet.Cells[1, 12] = "Dia Creado"; D.xlWorkSheet.Cells[1, 13] = "DER";
                    D.xlWorkSheet.Cells[1, 14] = "Genero"; D.xlWorkSheet.Cells[1, 15] = "#SS";
                    D.xlWorkSheet.Cells[1, 16] = "#Licensia"; D.xlWorkSheet.Cells[1, 17] = "Razon";
                    D.xlWorkSheet.Cells[1, 18] = "ID"; D.xlWorkSheet.Cells[1, 19] = "Se debe $";
                    D.xlWorkSheet.Cells[1, 20] = "Status"; D.xlWorkSheet.Cells[1, 21] = "DOT #";
                    D.xlWorkSheet.Cells[1, 22] = "DOB";
                    D.xlWorkBook.SaveAs(fi, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    // xlWorkBook.Close(true, misValue, misValue);
                    // xlApp.Quit();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

               
            
            } 
            try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;

                    textBox26.Text = D.xlWorkSheet.Cells[2, 11].Value.ToString();

                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        public void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || textBox9.Text == "" || textBox10.Text == "" || textBox11.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "" || comboBox4.Text == "" || textBox14.Text == "" || textBox25.Text == "")
            { MessageBox.Show("Write the information first!"); }
            else
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    objarray.Add(new db());
                    for (int y = objarray.Count; y < objarray.Count + 1; y++)
                    {

                        objarray[y - 1].name = textBox2.Text;
                        objarray[y - 1].lastname = textBox3.Text;
                        objarray[y - 1].phonenumber = Convert.ToDouble(textBox1.Text);
                        objarray[y - 1].age = Convert.ToDouble(textBox4.Text);
                        objarray[y - 1].address = textBox5.Text;
                        objarray[y - 1].city = textBox6.Text;
                        objarray[y - 1].state = textBox7.Text;
                        objarray[y - 1].company = textBox8.Text;
                        objarray[y - 1].resultado = comboBox1.Text;
                        objarray[y - 1].resultadoalcohol = comboBox2.Text;
                        objarray[y - 1].datecreated = dateTimePicker2.Value.ToString("MM.dd.yyyy");
                        objarray[y - 1].DER = textBox10.Text;
                        objarray[y - 1].genero = comboBox3.Text;
                        objarray[y - 1].SS = Convert.ToDouble(textBox11.Text);
                        objarray[y - 1].License = textBox9.Text;
                        objarray[y - 1].razon = comboBox4.Text;
                        objarray[y - 1].ID = D.count;
                        objarray[y - 1].money = Convert.ToDouble(textBox14.Text);
                        objarray[y - 1].status = "Activo";
                        objarray[y - 1].dotnumber = Convert.ToDouble(textBox25.Text);
                        objarray[y - 1].DOB = dateTimePicker1.Value.ToString("MM.dd.yyyy");

                        MessageBox.Show("Añadidos: " + Convert.ToString(objarray.Count));

                        D.xlWorkSheet.Cells[D.count + 1, 1] = objarray[y - 1].name;
                        D.xlWorkSheet.Cells[D.count + 1, 2] = objarray[y - 1].lastname;
                        D.xlWorkSheet.Cells[D.count + 1, 3] = objarray[y - 1].age;
                        D.xlWorkSheet.Cells[D.count + 1, 4] = objarray[y - 1].phonenumber;
                        D.xlWorkSheet.Cells[D.count + 1, 5] = objarray[y - 1].address;
                        D.xlWorkSheet.Cells[D.count + 1, 6] = objarray[y - 1].city;
                        D.xlWorkSheet.Cells[D.count + 1, 7] = objarray[y - 1].state;
                        D.xlWorkSheet.Cells[D.count + 1, 8] = objarray[y - 1].company;
                        D.xlWorkSheet.Cells[D.count + 1, 9] = objarray[y - 1].resultado;
                        D.xlWorkSheet.Cells[D.count + 1, 10] = objarray[y - 1].resultadoalcohol;
                        D.xlWorkSheet.Cells[D.count + 1, 12] = objarray[y - 1].datecreated;
                        D.xlWorkSheet.Cells[D.count + 1, 13] = objarray[y - 1].DER;
                        D.xlWorkSheet.Cells[D.count + 1, 14] = objarray[y - 1].genero;
                        D.xlWorkSheet.Cells[D.count + 1, 15] = objarray[y - 1].SS;
                        D.xlWorkSheet.Cells[D.count + 1, 16] = objarray[y - 1].License;
                        D.xlWorkSheet.Cells[D.count + 1, 17] = objarray[y - 1].razon;
                        D.xlWorkSheet.Cells[D.count + 1, 18] = objarray[y - 1].ID;
                        D.xlWorkSheet.Cells[D.count + 1, 19] = objarray[y - 1].money;
                        D.xlWorkSheet.Cells[D.count + 1, 20] = objarray[y - 1].status;
                        D.xlWorkSheet.Cells[D.count + 1, 21] = objarray[y - 1].dotnumber;
                        D.xlWorkSheet.Cells[D.count + 1, 22] = objarray[y - 1].DOB;
                        
                        D.count++;
                        D.xlWorkSheet.Cells[1, 11] = D.count;
                        D.xlWorkBook.Save();
                        D.xlWorkBook.Close(true, misValue, misValue);
                        // xlWorkBook.Close(true, misValue, misValue);
                        // xlApp.Quit();
                    }//end of for
                }//end of try

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }


            }//end of else
        }//end of void

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionTabs = new int[]  {43,131,221,250,304,375,640,713,756,906,1030,1090,1166,1241,1318,1459,1530,1590,1638};
        }
        string[] whatever = new string[100000];
        private void button2_Click(object sender, EventArgs e)//Actualizar boton
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" +
                        D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" +
                        D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                        D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" + 
                        D.xlWorkSheet.Cells[g, 12].Value.ToString()+ "\t" +
                        D.xlWorkSheet.Cells[g,20].Value.ToString() + "\t" +
                        D.xlWorkSheet.Cells[g,21].Value.ToString() + "\t" +
                        D.xlWorkSheet.Cells[g,22].Value.ToString());
                    richTextBox1.AppendText(Environment.NewLine);
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }


     // Printing Part
        private void button4_Click(object sender, EventArgs e)
        {   
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                D.batch = (int)(D.xlWorkSheet.Cells[3, 11] as Excel.Range).Value;
                int x = 27;
                //opening MemoPrint to print in format desired
                D.xlWorkBook2 = D.xlApp2.Workbooks.Open(fi3, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet2 = (Excel.Worksheet)D.xlWorkBook2.Worksheets.get_Item(1);
                //erase everything first
            
                while (x<500)
                {

                    D.xlWorkSheet2.Cells[x, 1] = "";
                            D.xlWorkSheet2.Cells[x, 2] = "";
                            D.xlWorkSheet2.Cells[x, 3] = "";
                            D.xlWorkSheet2.Cells[x, 4] = "";
                            D.xlWorkSheet2.Cells[x, 5] = "";
                            D.xlWorkSheet2.Cells[x, 6] = "";
                            D.xlWorkSheet2.Cells[x, 7] = "";
                            
                            D.xlWorkSheet2.get_Range("A" + x, "G" + x).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            D.xlWorkSheet2.get_Range("A" + x, "G" + x).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            D.xlWorkSheet2.get_Range("A" + x, "G" + x).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            D.xlWorkSheet2.get_Range("A" + x, "G" + x).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone; 
                    x++;
                }
                x = 27;

                for (int y = 2; y <= D.count; y++)
                {   
                    for (int h = 0; h < (totaldrugs); h++)
                    {
                        if (D.xlWorkSheet.Cells[y, 18].Value.ToString() == arr[h])
                        {
                            int number2 = drugalco.Next(0,RandDrugsAlcohol.Length);

                            D.xlWorkSheet2.Cells[20, 2] = D.xlWorkSheet.Cells[y, 13];
                            D.xlWorkSheet2.Cells[x, 1] = D.xlWorkSheet.Cells[y, 18];
                            D.xlWorkSheet2.Cells[x, 2] = D.xlWorkSheet.Cells[y, 15];
                            D.xlWorkSheet2.Cells[x, 3] = D.xlWorkSheet.Cells[y, 1];
                            D.xlWorkSheet2.Cells[x, 4] = D.xlWorkSheet.Cells[y, 2];
                            D.xlWorkSheet2.Cells[x, 5] = "Drugs";
                            D.xlWorkSheet2.Cells[x, 6] = "No";
                            D.xlWorkSheet2.Cells[x, 7] = DateTime.Today.ToString("MM.dd.yyyy");
                            x++;
                        }
                    }
                }
                for (int y = 2; y <= D.count; y++)
                {
                    for (int h = 0; h < (totalalcohol); h++)
                    {
                        if (D.xlWorkSheet.Cells[y, 18].Value.ToString() == temp[h])
                        {
                            //int number2 = drugalco.Next(0, RandDrugsAlcohol.Length);

                            D.xlWorkSheet2.Cells[20, 2] = D.xlWorkSheet.Cells[y, 13];
                            D.xlWorkSheet2.Cells[x, 1] = D.xlWorkSheet.Cells[y, 18];
                            D.xlWorkSheet2.Cells[x, 2] = D.xlWorkSheet.Cells[y, 15];
                            D.xlWorkSheet2.Cells[x, 3] = D.xlWorkSheet.Cells[y, 1];
                            D.xlWorkSheet2.Cells[x, 4] = D.xlWorkSheet.Cells[y, 2];
                            D.xlWorkSheet2.Cells[x, 5] = "Alcohol";
                            D.xlWorkSheet2.Cells[x, 6] = "No";
                            D.xlWorkSheet2.Cells[x, 7] = DateTime.Today.ToString("MM.dd.yyyy");
                            x++;
                        }
                    }
                }
               
                x++;
                D.xlWorkSheet2.Cells[x, 4] = (totalalcohol + totaldrugs) + " Participant(s) in company: " + companyforprint;



                D.xlWorkSheet2.get_Range("A" + x, "G" + x).Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 4d;
                D.xlWorkSheet2.get_Range("A" + x, "G" + x).Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;
                x = x + 2;
                D.xlWorkSheet2.get_Range("A" + x, "C" + (x+1)).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, (Excel.XlColorIndex)1, ColorTranslator.ToOle(Color.Black), Type.Missing);
                D.xlWorkSheet2.Cells[x, 1] = "Report Totals: 1 Company.";
                x++;
                D.xlWorkSheet2.Cells[x, 1] = "               " + (totalalcohol + totaldrugs) + " Participant(s).";
                x = x + 2;
                D.xlWorkSheet2.get_Range("A" +x,"G"+x).Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 2d;
                

                D.xlWorkSheet2.Cells[x, 1] = DateTime.Today.ToString("MM.dd.yyyy");
                D.xlWorkSheet2.Cells[x, 4] = "Random Selection - Selected Participants";
                x++;

                D.xlWorkSheet2.Cells[x, 1] = "§382.305: Random testing.";
                x++;
                D.xlWorkSheet2.Cells[x, 1] = "(a) Every employer shall comply with the requirements of this section. Every driver shall submit to random";
                x++;
                D.xlWorkSheet2.Cells[x, 1] = "alcohol and controlled substance testing as required in this section.";
                x++;
                D.xlWorkSheet2.Cells[x, 1] = "(b)(1) Except as provided in paragraphs (c) through (e) of this section, the minimum annual percentage rate";
                x++;
                D.xlWorkSheet2.Cells[x, 1] = "for random alcohol testing shall be 10 percent of the average number of driver positions.";
                x++;
                D.xlWorkSheet2.Cells[x, 1] = "(2) Except as provided in paragraphs (f) through (h) of this section, the minimum annual percentage rate for";
                x++;
                D.xlWorkSheet2.Cells[x, 1] = "random controlled substances testing shall be 50 percent of the average number of driver positions.";

                D.xlWorkSheet2.Cells[19, 2] = companyforprint;
                D.xlWorkSheet2.Cells[16, 3] = companyforprint;
             //   D.xlWorkSheet2.Cells[20, 4] = ;
                D.xlWorkSheet2.Cells[7, 6] = "Random Batch RB" + D.xlWorkSheet.Cells[3, 11].Value.ToString();
                D.batch++;
                D.xlWorkSheet.Cells[3, 11] = D.batch;
                D.xlWorkSheet2.Cells[17, 3] = DateTime.Today.ToString("MM.dd.yyyy");
                x = 27;
                D.xlWorkBook2.Save();
  

                D.xlWorkBook2.Close(true, misValue, misValue);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook2.Close(true, misValue, misValue);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            try
            {
             //  D.xlWorkBook2 = D.xlApp2.Workbooks.Open(fi3, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //    D.xlWorkSheet2 = (Excel.Worksheet)D.xlWorkBook2.Worksheets.get_Item(1);
                System.Diagnostics.Process.Start(fi3);
          //     D.xlWorkSheet2.Application.ActiveSheet.PrintPreview();
               // ((D.xlWorkSheet2).Application.ActiveSheet).PrintOut(misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

               // System.Threading.Thread.Sleep(10000);
             //   D.xlWorkBook2.Close(true, misValue, misValue);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook2.Close(true, misValue, misValue);
            }
            
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = ""; textBox4.Text = "";
            textBox5.Text = ""; textBox6.Text = "";
            textBox7.Text = ""; textBox8.Text = "";
            textBox11.Text = ""; textBox10.Text = "";
            textBox9.Text = "";
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    if (string.Equals(D.xlWorkSheet.Cells[g, 1].Value.ToString(), textBox12.Text,StringComparison.CurrentCultureIgnoreCase))
                    {
                        richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 12].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 20].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 21].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 22].Value.ToString());
                        richTextBox1.AppendText(Environment.NewLine);
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    if (string.Equals(D.xlWorkSheet.Cells[g, 2].Value.ToString(), textBox12.Text,StringComparison.CurrentCultureIgnoreCase))
                    {
                        richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 12].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 20].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 21].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 22].Value.ToString());
                        richTextBox1.AppendText(Environment.NewLine);
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)//Boton de borrar el ultimo registro
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;

                for (int g = 2; g <= D.count; g++)
                {
                    if (g == D.count)
                    {
                        D.xlWorkSheet.Cells[g, 1] = ""; D.xlWorkSheet.Cells[g, 2] = "";
                        D.xlWorkSheet.Cells[g, 3] = "";
                        D.xlWorkSheet.Cells[g, 18] = ""; D.xlWorkSheet.Cells[g, 4] = "";
                        D.xlWorkSheet.Cells[g, 17] = ""; D.xlWorkSheet.Cells[g, 5] = "";
                        D.xlWorkSheet.Cells[g, 16] = ""; D.xlWorkSheet.Cells[g, 6] = "";
                        D.xlWorkSheet.Cells[g, 15] = ""; D.xlWorkSheet.Cells[g, 7] = "";
                        D.xlWorkSheet.Cells[g, 14] = ""; D.xlWorkSheet.Cells[g, 8] = "";
                        D.xlWorkSheet.Cells[g, 13] = ""; D.xlWorkSheet.Cells[g, 9] = "";
                        D.xlWorkSheet.Cells[g, 12] = ""; D.xlWorkSheet.Cells[g, 10] = "";
                        D.xlWorkSheet.Cells[g, 19] = ""; D.xlWorkSheet.Cells[g, 20] = "";
                        D.xlWorkSheet.Cells[1, 11] = D.count - 1;
                        MessageBox.Show("Last register has been erased");
                    }
                }

                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    if (string.Equals(D.xlWorkSheet.Cells[g, 8].Value.ToString(), textBox12.Text,StringComparison.CurrentCultureIgnoreCase))
                    {
                        richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" +
                            D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                            D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" +
                            D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" + 
                            D.xlWorkSheet.Cells[g, 12].Value.ToString() + "\t" +
                            D.xlWorkSheet.Cells[g,20].Value.ToString() + "\t" +
                            D.xlWorkSheet.Cells[g,21].Value.ToString() + "\t" +
                            D.xlWorkSheet.Cells[g,22].Value.ToString());
                        //  richTextBox1.AppendText(whatever[g]);
                        richTextBox1.AppendText(Environment.NewLine);
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    if (D.xlWorkSheet.Cells[g, 15].Value.ToString() == textBox12.Text)
                    {
                        richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 12].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 20].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 21].Value.ToString() + "\t" +
                           D.xlWorkSheet.Cells[g, 22].Value.ToString());
                        richTextBox1.AppendText(Environment.NewLine);
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    if (D.xlWorkSheet.Cells[g, 16].Value.ToString() == textBox12.Text)
                    {
                        richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 12].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 20].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 21].Value.ToString() + "\t" +
                         D.xlWorkSheet.Cells[g, 22].Value.ToString());
                        richTextBox1.AppendText(Environment.NewLine);
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox12.Text)
                    {
                        richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 12].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 20].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 21].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 22].Value.ToString());
                        richTextBox1.AppendText(Environment.NewLine);
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox1.Text = "";
                //MessageBox.Show(Convert.ToString(objarray.Count));
                for (int g = 2; g <= D.count; g++)
                {
                    if (D.xlWorkSheet.Cells[g, 12].Value.ToString() == textBox12.Text)
                    {
                        richTextBox1.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 3].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 14].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 4].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 5].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 6].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 7].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 16].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 9].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 10].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 17].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 12].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 20].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 21].Value.ToString() + "\t" +
                          D.xlWorkSheet.Cells[g, 22].Value.ToString());
                        richTextBox1.AppendText(Environment.NewLine);
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void printPreviewControl1_Click(object sender, EventArgs e)
        {

        }   

        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {

        }
        private void richTextBox2_TextChanged_1(object sender, EventArgs e)
        {
            richTextBox2.SelectionTabs = new int[] { 48, 120,200 };
        }
               
        private void button13_Click(object sender, EventArgs e)//Actualizar Random boton
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                richTextBox2.Text = ""; richTextBox3.Text = "";
                richTextBox2.AppendText(Environment.NewLine);
                richTextBox3.AppendText(Environment.NewLine);
                arr.Clear();
                arr2.Clear();
                temp.Clear();
                double totalselected = 0;

                companyforprint = textBox13.Text;
                
           

                for (int g = 2; g <= D.count; g++)
                {
                    if (string.Equals(D.xlWorkSheet.Cells[g, 8].Value.ToString(), textBox13.Text, StringComparison.CurrentCultureIgnoreCase) && D.xlWorkSheet.Cells[g,20].Value.ToString() == "Activo")
                    {
                        richTextBox2.AppendText(whatever[g] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 15].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 16].Value.ToString());
                        richTextBox2.AppendText(Environment.NewLine);
                        arr2.Add(D.xlWorkSheet.Cells[g, 18].Value.ToString());
                        totalselected++;
                        textBox31.Text = Convert.ToString(totalselected);
                    }
                }
                totaldrugs = Convert.ToInt16(totalselected * (.125));
                totalalcohol = Convert.ToInt16(totalselected * (.025));
                /////////////////////DRUGS/////////////////////
                for (int i = 0; i < totaldrugs; i++)
                {//richTextBox3.Text += "doing";
                    string number = Convert.ToString(rAnD.Next(0, arr2.Count));

                    if (!arr.Contains(arr2[Convert.ToInt16(number)]))
                    {
                        arr.Add((arr2[Convert.ToInt16(number)]));
                    }
                    else
                    {
                        i--;
                    }
                }
     
                richTextBox3.AppendText("Drugs:");
                richTextBox3.AppendText(Environment.NewLine);
                richTextBox3.AppendText(Environment.NewLine);
                for (int y = 2; y <= D.count; y++)
                {
                    for (int h = 0; h < totaldrugs; h++)
                    {
                        if (D.xlWorkSheet.Cells[y, 18].Value.ToString() == arr[h])
                        {
                          
                            richTextBox3.AppendText(whatever2[y] = D.xlWorkSheet.Cells[y, 18].Value.ToString() + "\t" + D.xlWorkSheet.Cells[y, 15].Value.ToString() + "\t" + D.xlWorkSheet.Cells[y, 1].Value.ToString() + "\t" + D.xlWorkSheet.Cells[y, 16].Value.ToString());
                            richTextBox3.AppendText(Environment.NewLine);

                        }
                    }
                }

                //////////////ALCOHOL/////////////////
                for (int i = 0; i < totalalcohol; i++)
                {//richTextBox3.Text += "doing";
                    string number2 = Convert.ToString(drugalco.Next(0, arr2.Count));
                    number2 = Convert.ToString(rAnD.Next(0, arr2.Count));
                    if (!temp.Contains(arr2[Convert.ToInt16(number2)]))
                    {
                        temp.Add((arr2[Convert.ToInt16(number2)]));
                    }
                    else
                    {
                        i--;
                    }
                }
                
                richTextBox3.AppendText(Environment.NewLine);
                richTextBox3.AppendText("Alcohol:");
                richTextBox3.AppendText(Environment.NewLine);
                richTextBox3.AppendText(Environment.NewLine); 
                for (int y = 2; y <= D.count; y++)
                {
                    for (int h = 0; h < totalalcohol; h++)
                    {
                        if (D.xlWorkSheet.Cells[y, 18].Value.ToString() == temp[h])
                        {
                                                  
                            richTextBox3.AppendText(whatever2[y] = D.xlWorkSheet.Cells[y, 18].Value.ToString() + "\t" + D.xlWorkSheet.Cells[y, 15].Value.ToString() + "\t" + D.xlWorkSheet.Cells[y, 1].Value.ToString() + "\t" + D.xlWorkSheet.Cells[y, 16].Value.ToString());
                            richTextBox3.AppendText(Environment.NewLine);

                        }
                    }
                }

                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            richTextBox3.SelectionTabs = new int[] { 48, 120,200 };
        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("www.facebook.com/adrianleall");
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (var series in chart1.Series)
                {
                    series.Points.Clear();
                }
                foreach (var series in chart2.Series)
                {
                    series.Points.Clear();
                }
                int countpre = 0;
                int countran = 0; int countcause = 0; int countacc = 0;
                int countduty = 0; int countfollow = 0; int countother = 0;
                int countNA = 0;
                int countNONpre = 0; int countNONran = 0; int countNONcause = 0;
                int countNONacc = 0; int countNONduty = 0; int countNONfollow = 0;
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                for (int g = 2; g <= D.count; g++)
                {
                    if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "N/A", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countNA++;
                    }
   
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "Pre Employment", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countpre++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "Random", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countran++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "Reasonable Cause", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countcause++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "Post Accident", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countacc++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "Return to Duty", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countduty++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "Follow Up", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countfollow++;
                    }
                        //NON new
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "NON DOT Pre Employment", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countNONpre++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "NON DOT Random", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countNONran++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "NON DOT Reasonable Cause", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countNONcause++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "NON DOT Post Accident", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countNONacc++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "NON DOT Return to Duty", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countNONduty++;
                    }
                    else if (string.Equals(D.xlWorkSheet.Cells[g, 17].Value.ToString(), "NON DOT Follow Up", StringComparison.CurrentCultureIgnoreCase))
                    {
                        countNONfollow++;
                    }
                    else
                    {
                        countother++;
                    }
                }
                this.chart1.Series["Total"].Points.AddXY("N/A", countNA);
                this.chart1.Series["Total"].Points.AddXY("Pre-E", countpre);
                this.chart1.Series["Total"].Points.AddXY("Random", countran);
                this.chart1.Series["Total"].Points.AddXY("R-Cause", countcause);
                this.chart1.Series["Total"].Points.AddXY("P-Accident", countacc);
                this.chart1.Series["Total"].Points.AddXY("Ret-Duty", countduty);
                this.chart1.Series["Total"].Points.AddXY("Follow Up", countfollow);
                this.chart1.Series["Total"].Points.AddXY("NON Pre-E", countNONpre);
                this.chart1.Series["Total"].Points.AddXY("NON Random", countNONran);
                this.chart1.Series["Total"].Points.AddXY("NON R-Cause", countNONcause);
                this.chart1.Series["Total"].Points.AddXY("NON P-Accident", countNONacc);
                this.chart1.Series["Total"].Points.AddXY("NON Ret-Duty", countNONduty);
                this.chart1.Series["Total"].Points.AddXY("NON Follow Up", countNONfollow);
                this.chart1.Series["Total"].Points.AddXY("Other", countother);
                chart1.ChartAreas["ChartArea1"].AxisX.Interval = 1;

                this.chart2.Series["Total"].Points.AddXY("N/A", countNA);
                this.chart2.Series["Total"].Points.AddXY("Pre-E", countpre);
                this.chart2.Series["Total"].Points.AddXY("Random", countran);
                this.chart2.Series["Total"].Points.AddXY("R-Cause", countcause);
                this.chart2.Series["Total"].Points.AddXY("P-Accident", countacc);
                this.chart2.Series["Total"].Points.AddXY("Ret-Duty", countduty);
                this.chart2.Series["Total"].Points.AddXY("Follow Up", countfollow);
                this.chart2.Series["Total"].Points.AddXY("NON Pre-E", countNONpre);
                this.chart2.Series["Total"].Points.AddXY("NON Random", countNONran);
                this.chart2.Series["Total"].Points.AddXY("NON R-Cause", countNONcause);
                this.chart2.Series["Total"].Points.AddXY("NON P-Accident", countNONacc);
                this.chart2.Series["Total"].Points.AddXY("NON Ret-Duty", countNONduty);
                this.chart2.Series["Total"].Points.AddXY("NON Follow Up", countNONfollow);
                this.chart2.Series["Total"].Points.AddXY("Other", countother);
                chart2.ChartAreas["ChartArea1"].AxisX.Interval = 1;

                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void label45_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void chart2_Click(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            
            if (textBox16.Text == "")
            {
                double totalmoney = 0;
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox4.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    
                    string date1 = textBox15.Text;
                    string date2 = textBox17.Text;

                    for (int p = 2; p <= D.count; p++)
                    {
                        if (D.xlWorkSheet.Cells[p, 12].Value.ToString() == date1)
                        {
                            for (int g = p; g <= D.count; g++)
                            {
                                

                                richTextBox4.AppendText(whatever[p] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" + "$" + D.xlWorkSheet.Cells[g, 19].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 12].Value.ToString());
                                richTextBox4.AppendText(Environment.NewLine);
                                totalmoney += D.xlWorkSheet.Cells[g, 19].Value;
                               if (Convert.ToString(D.xlWorkSheet.Cells[g + 1, 12].Value.ToString()) == date2)
                                {
                                    p = g;
                                    break;
                                }
                            }
                        }
                        if (Convert.ToString(D.xlWorkSheet.Cells[p + 1, 12].Value.ToString()) == date2)
                        {
                            break;
                        }
                    }
                    textBox18.Text = "$" + Convert.ToString(totalmoney);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    textBox18.Text = "$" + Convert.ToString(totalmoney);
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                double totalmoney2 = 0;
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox4.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    
                    string date1 = textBox15.Text;
                    string date2 = textBox17.Text;

                    for (int p = 2; p <= D.count; p++)
                    {
                        if (D.xlWorkSheet.Cells[p, 12].Value.ToString() == date1)
                        {
                            for (int g = p; g <= D.count; g++)
                            {

                                if (string.Equals(D.xlWorkSheet.Cells[g, 8].Value.ToString(), textBox16.Text, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    richTextBox4.AppendText(whatever[p] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" + "$" + D.xlWorkSheet.Cells[g, 19].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 12].Value.ToString());
                                    richTextBox4.AppendText(Environment.NewLine);
                                    totalmoney2 += D.xlWorkSheet.Cells[g, 19].Value;
                                }
                                if (Convert.ToString(D.xlWorkSheet.Cells[g + 1, 12].Value.ToString()) == date2)
                                {
                                    p = g;
                                    break;
                                }
                            }
                        }
                        if (Convert.ToString(D.xlWorkSheet.Cells[p + 1, 12].Value.ToString()) == date2)
                        {
                            break;
                        }
                    }
                    textBox18.Text = "$" + Convert.ToString(totalmoney2);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    textBox18.Text = "$" + Convert.ToString(totalmoney2);
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {
            richTextBox4.SelectionTabs = new int[] { 46, 125, 220, 262, 438, 587};
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 1].Value.ToString();
                            D.xlWorkSheet.Cells[g, 1] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 1].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {

        }

        private void label66_Click(object sender, EventArgs e)
        {

        }

        private void button32_Click(object sender, EventArgs e)
        {
            if (groupBox1.Visible == false)
            {
                groupBox1.Visible = true;
            }
            else if (groupBox1.Visible == true)
            {
                groupBox1.Visible = false;
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 2].Value.ToString();
                            D.xlWorkSheet.Cells[g, 2] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 2].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 3].Value.ToString();
                            D.xlWorkSheet.Cells[g, 3] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 3].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 14].Value.ToString();
                            D.xlWorkSheet.Cells[g, 14] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 14].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 4].Value.ToString();
                            D.xlWorkSheet.Cells[g, 4] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 4].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 5].Value.ToString();
                            D.xlWorkSheet.Cells[g, 5] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 5].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 6].Value.ToString();
                            D.xlWorkSheet.Cells[g, 6] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 6].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 7].Value.ToString();
                            D.xlWorkSheet.Cells[g, 7] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 7].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 8].Value.ToString();
                            D.xlWorkSheet.Cells[g, 8] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 8].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 13].Value.ToString();
                            D.xlWorkSheet.Cells[g, 13] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 13].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 16].Value.ToString();
                            D.xlWorkSheet.Cells[g, 16] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 16].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 9].Value.ToString();
                            D.xlWorkSheet.Cells[g, 9] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 9].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 10].Value.ToString();
                            D.xlWorkSheet.Cells[g, 10] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 10].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 17].Value.ToString();
                            D.xlWorkSheet.Cells[g, 17] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 17].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 20].Value.ToString();
                            D.xlWorkSheet.Cells[g, 20] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 20].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 19].Value.ToString();
                            D.xlWorkSheet.Cells[g, 19] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 19].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure?", "Completely sure?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (textBox16.Text == "")
                {

                    double totalmoney = 0;
                    try
                    {
                        D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                        D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                        D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                        D.count = D.counter;
                        richTextBox4.Text = "";
                        //MessageBox.Show(Convert.ToString(objarray.Count));

                        string date1 = textBox15.Text;
                        string date2 = textBox17.Text;

                        for (int p = 2; p <= D.count; p++)
                        {
                            if (D.xlWorkSheet.Cells[p, 12].Value.ToString() == date1)
                            {
                                for (int g = p; g <= D.count; g++)
                                {

                                    D.xlWorkSheet.Cells[g, 19] = 0;
                                    richTextBox4.AppendText(whatever[p] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" + "$" + D.xlWorkSheet.Cells[g, 19].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 12].Value.ToString());
                                    richTextBox4.AppendText(Environment.NewLine);
                                    totalmoney += D.xlWorkSheet.Cells[g, 19].Value;
                                    if (Convert.ToString(D.xlWorkSheet.Cells[g + 1, 12].Value.ToString()) == date2)
                                    {
                                        p = g;
                                        break;
                                    }
                                }
                            }
                            if (Convert.ToString(D.xlWorkSheet.Cells[p + 1, 12].Value.ToString()) == date2)
                            {
                                break;
                            }
                        }
                        textBox18.Text = "$" + Convert.ToString(totalmoney);
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
                    catch (Exception ex)
                    {
                        textBox18.Text = "$" + Convert.ToString(totalmoney);
                        MessageBox.Show(ex.Message);
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
                }
                else
                {
                    double totalmoney2 = 0;
                    try
                    {
                        D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                        D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                        D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                        D.count = D.counter;
                        richTextBox4.Text = "";
                        //MessageBox.Show(Convert.ToString(objarray.Count));

                        string date1 = textBox15.Text;
                        string date2 = textBox17.Text;

                        for (int p = 2; p <= D.count; p++)
                        {
                            if (D.xlWorkSheet.Cells[p, 12].Value.ToString() == date1)
                            {
                                for (int g = p; g <= D.count; g++)
                                {

                                    if (string.Equals(D.xlWorkSheet.Cells[g, 8].Value.ToString(), textBox16.Text, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        D.xlWorkSheet.Cells[g, 19] = 0;
                                        richTextBox4.AppendText(whatever[p] = D.xlWorkSheet.Cells[g, 18].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 1].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 2].Value.ToString() + "\t" + "$" + D.xlWorkSheet.Cells[g, 19].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 8].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 13].Value.ToString() + "\t" + D.xlWorkSheet.Cells[g, 12].Value.ToString());
                                        richTextBox4.AppendText(Environment.NewLine);
                                        totalmoney2 += D.xlWorkSheet.Cells[g, 19].Value;
                                    }
                                    if (Convert.ToString(D.xlWorkSheet.Cells[g + 1, 12].Value.ToString()) == date2)
                                    {
                                        p = g;
                                        break;
                                    }
                                }
                            }
                            if (Convert.ToString(D.xlWorkSheet.Cells[p + 1, 12].Value.ToString()) == date2)
                            {
                                break;
                            }
                        }
                        textBox18.Text = "$" + Convert.ToString(totalmoney2);
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
                    catch (Exception ex)
                    {
                        textBox18.Text = "$" + Convert.ToString(totalmoney2);
                        MessageBox.Show(ex.Message);
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {

            }
        }

        private void label69_Click(object sender, EventArgs e)
        {

        }
      
        private void button34_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(fi5photo);
              try
                    {
                        D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                        D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                        D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                        D.count = D.counter;

                         textBox23.Text = Convert.ToString((D.count -1));
                    
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
            
        }
        private void button35_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start( fi6startphoto +"Picture " + textBox24.Text + ".jpg");
            }
            catch (Exception ex)
            {              
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {

        }

        private void button36_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure?", "Completely sure?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
               

                    try
                    {
                        D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                        D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                        D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                        D.count = D.counter;

                        if (File.Exists( fi6startphoto + "Picture " + (D.count - 1) + ".jpg"))
                        {
                            File.Delete( fi6startphoto + "Picture " + (D.count - 1) + ".jpg");
                        }
                        else
                        {
                            MessageBox.Show("There does not exist a photo of the last ID registered","Something is wrong");
                        }
                    
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
                    catch (Exception ex)
                    {
                    
                        MessageBox.Show(ex.Message);
                        D.xlWorkBook.Close(true, misValue, misValue);
                    }
                
            }
            else if (dialogResult == DialogResult.No)
            {

            }
        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {

        }

        private void button37_Click(object sender, EventArgs e)
        {
            if (label1.Text == "Nombre:")
            {
                tabPage1.Text = "Register"; tabPage2.Text = "Data";
                tabPage3.Text = "Graphs"; tabPage4.Text = "Print Random";
                tabPage5.Text = "Money/Edit Data"; tabPage7.Text = "Account/ Print Bills";
                label1.Text = "Name:"; label2.Text = "Last Name:";
                label3.Text = "Age:"; label26.Text = "Gender:";
                label6.Text = "Phone#:"; label5.Text = "Address:";
                label8.Text = "City:"; label7.Text = "State";
                label10.Text = "Company:"; label23.Text = "SSN:";
                label24.Text = "License #:"; label47.Text = "Pay:";
                label9.Text = "Drug result:"; label20.Text = "Alcohol result:";
                label33.Text = "Reason"; label72.Text = "ID of this photo:";
                button3.Text = "Erase..."; button1.Text = "Register";
                button37.Text = "Spanish/English"; button35.Text = "See photo by ID";
                button34.Text = "Take photo"; button36.Text = "Erase last photo taken";
                label79.Text = "Bill #:";

                button2.Text = "Update"; label31.Text = "Search by:";
                button5.Text = "Name"; button6.Text = "Last Name";
                button7.Text = "Company"; button8.Text = "SSN";
                button9.Text = "License"; button10.Text = "Date";
                button12.Text = "Erase last register";
                label11.Text = "Name:"; label12.Text = "Last Name:";
                label13.Text = "Age:"; label27.Text = "Gender:";
                label14.Text = "Phone#:"; label15.Text = "Address:";
                label16.Text = "City:"; label17.Text = "State:";
                label18.Text = "Company:"; label29.Text = "SSN:";
                label30.Text = "License#:"; label19.Text = "Result-D:";
                label21.Text = "Result-A:"; label34.Text = "Reason:";
                label22.Text = "Date:";

                label35.Text = "Name of the Company:";
                button13.Text = "Update Random"; button4.Text = "Print";
                label36.Text = "=List of the Company="; label37.Text = "=Selected=";
                label39.Text = "SSN:"; label38.Text = "Name:"; label57.Text = "License#:";
                label41.Text = "SSN:"; label40.Text = "Name:"; label58.Text = "License#:";
                label44.Text = "Made by:";

                button14.Text = "Update"; label45.Text = "Graphs to observe the total for every reason of why the test was made";

                label49.Text = "Date 1:"; label50.Text = "Date 2:";
                label51.Text = "Company:"; label52.Text = "Name:";
                label53.Text = "Last name:"; label55.Text = "Amount";
                label59.Text = "Company:"; label56.Text = "Date:";
                label62.Text = "Note: Use format (mm.dd.yyyy).";
                label61.Text = "Total of this list is:";
                button33.Text = "Edit money in this list to $0"; button15.Text = "Update";
                groupBox1.Text = "Edit";
                button32.Text = "Edit"; label63.Text = "Edit by ID any of the following data";
                label65.Text = "Change to"; label67.Text = "Changed from:"; label68.Text = "To:";
                label66.Text = "Note: If ID is unknown, search by name, last name, company," + "\n" + "social security number or license number in 'Data'.";
                label69.Text = "Note: Care with whatever is edited, it can cause inestability if" + "\n" + "numbers are changed for letters and viceversa (example: money)." + "\n" + "If any mistake is made, edit immediately.";
                button16.Text = "Name"; button17.Text = "Last Name"; button18.Text = "Age";
                button23.Text = "Gender"; button19.Text = "Phone"; button29.Text = "Address";
                button28.Text = "City"; button21.Text = "State"; button22.Text = "Company";
                button25.Text = "License#"; button27.Text = "Result-D"; button20.Text = "Result-A";
                button26.Text = "Reason"; button30.Text = "Money";
                button41.Text = "Date";

                //tab 7 cuenta imprimir
                label75.Text = "Bank Account";
                label77.Text = "Sum/Substract:";
                label78.Text = "Change total to:";
                button38.Text = "Update";
                button39.Text = "Save Changes";
                label80.Text = "Print Bill";
                button40.Text = "Print";

            }
            else if (label1.Text == "Name:")
            {
                tabPage1.Text = "Registro"; tabPage2.Text = "Ver Datos";
                tabPage3.Text = "Graficas"; tabPage4.Text = "Imprimir Random";
                tabPage5.Text = "Dinero/Editar Registro";
                label1.Text = "Nombre:"; label2.Text = "Apellido:";
                label3.Text = "Edad:"; label26.Text = "Genero:";
                label6.Text = "#Telefono:"; label5.Text = "Direccion:";
                label8.Text = "Ciudad:"; label7.Text = "Estado";
                label10.Text = "Compañia:"; label23.Text = "# SS:";
                label24.Text = "# Licencia:"; label47.Text = "Se Debe:";
                label9.Text = "Resultado de drogas:"; label20.Text = "Resultado de alcohol:";
                label33.Text = "Razon"; label72.Text = "ID de esta foto:";
                button3.Text = "Borrar..."; button1.Text = "Registrar";
                button37.Text = "Español/Ingles"; button35.Text = "Ver foto por ID";
                button34.Text = "Tomar Foto"; button36.Text = "Borrar ultima foto tomada";
                label79.Text = "Factura #:";

                button2.Text = "Actualizar"; label31.Text = "Buscar por:";
                button5.Text = "Nombre"; button6.Text = "Apellido";
                button7.Text = "Compañia"; button8.Text = "#SS";
                button9.Text = "Licencia"; button10.Text = "Fecha";
                button12.Text = "Borrar ultimo registro";
                label11.Text = "Nombre:"; label12.Text = "Apellido:";
                label13.Text = "Edad:"; label27.Text = "Genero:";
                label14.Text = "Telefono:"; label15.Text = "Direccion:";
                label16.Text = "Ciudad:"; label17.Text = "Estado:";
                label18.Text = "Compa♫ia:"; label29.Text = "#SS:";
                label30.Text = "#Licensia:"; label19.Text = "Resultado-D:";
                label21.Text = "Resultado-A:"; label34.Text = "Razon:";
                label22.Text = "Fecha:";

                label35.Text = "Nombre de la compañia:";
                button13.Text = "Actualizar Random"; button4.Text = "Imprimir";
                label36.Text = "=Lista de la Compañia="; label37.Text = "=Seleccionados=";
                label39.Text = "#SS:"; label38.Text = "Nombre:"; label57.Text = "#Licencia:";
                label41.Text = "#SS:"; label40.Text = "Nombre:"; label58.Text = "#Licencia:";
                label44.Text = "Hecho por:";

                button14.Text = "Actualizar"; label45.Text = "Grafica para observar el total de cada razon por la cual se hicieron el test";

                label49.Text = "De Fecha:"; label50.Text = "A Fecha:";
                label51.Text = "Compañia:"; label52.Text = "Nombre:";
                label53.Text = "Apellido:"; label55.Text = "Se Debe";
                label59.Text = "Compañia:"; label56.Text = "Fecha:";
                label62.Text = "Nota: Usar formato (mm.dd.yyyy).";
                label61.Text = "Total de esta lista es:";
                button33.Text = "Editar todo el dinero en esta lista a $0"; button15.Text = "Actualizar";
                groupBox1.Text = "Editar";
                button32.Text = "Editar"; label63.Text = "Editar por ID cualquiera de los siguientes campos";
                label65.Text = "Cambiar por:"; label67.Text = "Se cambio de:"; label68.Text = "A:";
                label66.Text = "Nota: Si no se sabe el ID, buscar por nombre, apellido, compañia," + "\n" + "numero de seguro social o numero de licencia en 'Ver Datos'.";
                label69.Text = "Nota: Mucho cuidado con lo que se edita, podria causar" + "\n" + "errores si cambia numeros por letras o viceversa (ejemplo: dinero)." + "\n" + "Si se comete una equivocacion, editar denuevo imediatamente.";
                button16.Text = "Nombre"; button17.Text = "Apellido"; button18.Text = "Edad";
                button23.Text = "Genero"; button19.Text = "Telefono"; button29.Text = "Direccion";
                button28.Text = "Ciudad"; button21.Text = "Estado"; button22.Text = "Compañia";
                button25.Text = "#Licencia"; button27.Text = "Resultado-D"; button20.Text = "Resultado-A";
                button26.Text = "Razon"; button30.Text = "Dinero";
                button41.Text = "Fecha";

                //tab 7 cuenta imprimir
                label75.Text = "Cuenta Bancaria";
                label77.Text = "Suma/Resta:";
                label78.Text = "Cambiar total a:";
                button38.Text = "Actualizar";
                button39.Text = "Guardar Cambios";
                label80.Text = "Imprimir Factura";
                button40.Text = "Imprimir";

            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label47_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void label72_Click(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label34_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void label70_Click(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        private void label39_Click(object sender, EventArgs e)
        {

        }

        private void label38_Click(object sender, EventArgs e)
        {

        }

        private void label57_Click(object sender, EventArgs e)
        {

        }

        private void label41_Click(object sender, EventArgs e)
        {

        }

        private void label40_Click(object sender, EventArgs e)
        {

        }

        private void label58_Click(object sender, EventArgs e)
        {

        }

        private void label44_Click(object sender, EventArgs e)
        {

        }

        private void label49_Click(object sender, EventArgs e)
        {

        }

        private void label50_Click(object sender, EventArgs e)
        {

        }

        private void label51_Click(object sender, EventArgs e)
        {

        }

        private void label62_Click(object sender, EventArgs e)
        {

        }

        private void label52_Click(object sender, EventArgs e)
        {

        }

        private void label53_Click(object sender, EventArgs e)
        {

        }

        private void label55_Click(object sender, EventArgs e)
        {

        }

        private void label59_Click(object sender, EventArgs e)
        {

        }

        private void label60_Click(object sender, EventArgs e)
        {

        }

        private void label56_Click(object sender, EventArgs e)
        {

        }

        private void label61_Click(object sender, EventArgs e)
        {

        }

        private void label63_Click(object sender, EventArgs e)
        {

        }

        private void label65_Click(object sender, EventArgs e)
        {

        }

        private void label67_Click(object sender, EventArgs e)
        {

        }

        private void label68_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate(toolStripTextBox1.Text);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            webBrowser1.GoBack();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            webBrowser1.GoForward();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            webBrowser1.Refresh();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            webBrowser1.Stop();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            webBrowser1.GoHome();
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {
            this.toolStripTextBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(CheckEnter);
        }
        private void CheckEnter(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                webBrowser1.Navigate(toolStripTextBox1.Text);
            }
        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser1.ScriptErrorsSuppressed = true;
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void button38_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;

                double total1 = Convert.ToDouble(textBox26.Text);
                double sum = Convert.ToDouble(textBox27.Text);
                double total2 = total1 + sum;
                textBox28.Text = "$" + total2;

                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {

        }

        private void button39_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;

                double total1 = Convert.ToDouble(textBox26.Text);
                double sum = Convert.ToDouble(textBox27.Text);
                double total2 = total1 + sum;
                textBox28.Text = "$" + total2;
                D.xlWorkSheet.Cells[2, 11] = total2;
                textBox26.Text = D.xlWorkSheet.Cells[2, 11].Value.ToString();

                D.xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
        }

        private void label75_Click(object sender, EventArgs e)
        {

        }

        private void label77_Click(object sender, EventArgs e)
        {

        }

        private void label78_Click(object sender, EventArgs e)
        {

        }

        private void label79_Click(object sender, EventArgs e)
        {

        }

        private void button40_Click(object sender, EventArgs e)
        {

        }

        private void label80_Click(object sender, EventArgs e)
        {

        }

        private void textBox29_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {

        }

        private void button41_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 12].Value.ToString();
                            D.xlWorkSheet.Cells[g, 12] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 12].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }
        private void textBox31_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void label46_Click(object sender, EventArgs e)
        {

        }

        private void button42_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 22].Value.ToString();
                            D.xlWorkSheet.Cells[g, 22] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 22].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox19.Text != "")
            {
                try
                {
                    D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                    D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                    D.count = D.counter;
                    richTextBox1.Text = "";
                    //MessageBox.Show(Convert.ToString(objarray.Count));
                    for (int g = 2; g <= D.count; g++)
                    {
                        if (D.xlWorkSheet.Cells[g, 18].Value.ToString() == textBox19.Text)
                        {
                            textBox21.Text = D.xlWorkSheet.Cells[g, 15].Value.ToString();
                            D.xlWorkSheet.Cells[g, 15] = textBox20.Text;
                            textBox22.Text = D.xlWorkSheet.Cells[g, 15].Value.ToString();
                        }
                        else if (Convert.ToDouble(textBox19.Text) >= D.count)
                        {
                            MessageBox.Show("The ID entered does not exist", "You are Wrong!");
                            break;
                        }
                    }
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    D.xlWorkBook.Close(true, misValue, misValue);
                }
            }
            else
            {
                MessageBox.Show("Write in the two boxes first");
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void button44_Click(object sender, EventArgs e)
        {
            try
            {
                D.xlWorkBook = D.xlApp.Workbooks.Open(fi2, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet = (Excel.Worksheet)D.xlWorkBook.Worksheets.get_Item(1);
                D.counter = (double)(D.xlWorkSheet.Cells[1, 11] as Excel.Range).Value;
                D.count = D.counter;
                //MessageBox.Show(Convert.ToString(objarray.Count));
                D.xlWorkBook3 = D.xlApp3.Workbooks.Open(fi7, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                D.xlWorkSheet3 = (Excel.Worksheet)D.xlWorkBook3.Worksheets.get_Item(1);

                allDates.Clear();
                int x = 6;
                string date1 = textBox32.Text;
                string date2 = textBox33.Text;
              
                D.xlWorkSheet3.Cells[2, 1] = textBox32.Text + " to " + textBox33.Text;

                DateTime date = Convert.ToDateTime(date1);
                DateTime date22 = Convert.ToDateTime(date2);
               // dateTimePicker2.Value.ToString("MM.dd.yyyy");
                for (date=date; date <= date22; date = date.AddDays(1))
                {
                    allDates.Add(date);
                }
                while (x < 500)
                {

                    D.xlWorkSheet3.Cells[x, 1] = "";
                    D.xlWorkSheet3.Cells[x, 2] = "";
                    D.xlWorkSheet3.Cells[x, 3] = "";
                    D.xlWorkSheet3.Cells[x, 4] = "";
                    D.xlWorkSheet3.Cells[x, 5] = "";
                    D.xlWorkSheet3.Cells[x, 6] = "";
                    D.xlWorkSheet3.Cells[x, 7] = "";
                    x++;
                }
                x = 6;
               
                for (int p = 2; p <= D.count; p++)
                {
                    for (int g = 0; g < allDates.Count; g++)
                    {

                      if (Convert.ToDateTime(D.xlWorkSheet.Cells[p, 12].Value.ToString()) == allDates[g])
                      {
                       
                            if (D.xlWorkSheet.Cells[p, 19].Value.ToString() != "0")
                            {

                             
                                D.xlWorkSheet3.Cells[x, 1] = D.xlWorkSheet.Cells[p, 18];
                                D.xlWorkSheet3.Cells[x, 2] = D.xlWorkSheet.Cells[p, 1];
                                D.xlWorkSheet3.Cells[x, 3] = D.xlWorkSheet.Cells[p, 2];
                                D.xlWorkSheet3.Cells[x, 4] = D.xlWorkSheet.Cells[p, 19];
                                D.xlWorkSheet3.Cells[x, 5] = D.xlWorkSheet.Cells[p, 8];
                                D.xlWorkSheet3.Cells[x, 6] = D.xlWorkSheet.Cells[p, 13];
                                D.xlWorkSheet3.Cells[x, 7] = D.xlWorkSheet.Cells[p, 12];
                                x++;
                                break;
                            }
                       }
                    }
                }
                D.xlWorkBook.Close(true, misValue, misValue);
                D.xlWorkBook3.Close(true, misValue, misValue);
                x = 6;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook3.Close(true, misValue, misValue);
                D.xlWorkBook.Close(true, misValue, misValue);
            }
            try
            {
                //  D.xlWorkBook2 = D.xlApp2.Workbooks.Open(fi3, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                //    D.xlWorkSheet2 = (Excel.Worksheet)D.xlWorkBook2.Worksheets.get_Item(1);
                System.Diagnostics.Process.Start(fi7);
                //     D.xlWorkSheet2.Application.ActiveSheet.PrintPreview();
                // ((D.xlWorkSheet2).Application.ActiveSheet).PrintOut(misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                // System.Threading.Thread.Sleep(10000);
                //   D.xlWorkBook2.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                D.xlWorkBook3.Close(true, misValue, misValue);
            }
            
        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void textBox34_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void button45_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Bitmap bitmap1;
            bitmap1 = (Bitmap)pictureBox2.Image;
            bitmap1.RotateFlip(RotateFlipType.Rotate90FlipNone);
            pictureBox2.Image = bitmap1;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Bitmap bitmap2;
            bitmap2 = (Bitmap)pictureBox4.Image;
            bitmap2.RotateFlip(RotateFlipType.Rotate90FlipNone);
            pictureBox4.Image = bitmap2;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Bitmap bitmap3;
            bitmap3 = (Bitmap)pictureBox3.Image;
            bitmap3.RotateFlip(RotateFlipType.Rotate90FlipNone);
            pictureBox3.Image = bitmap3;
        }
        //END
    }
}
