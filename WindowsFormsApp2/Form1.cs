using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.IO;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Sacha";
        static readonly string SpreadsheetId = "1RGGGdqUb-O137YLCGbrMvEBb7MslNqEiAWn4yE5R6bE";
        static readonly string sheet = "Class Data";
        static SheetsService service;
        int clickCounter = 1;
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            GoogleCredential credential;
            using (var stream = new FileStream("My First Project-2dfed0050064.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            var start = dateTimePicker1.Value;
            var end = dateTimePicker2.Value;

            var valueRange = new ValueRange();
            logic g = new logic();
            List<string> dates = new List<string>();
            g.initializeDates(start, end, ref dates);
            char ch = 'G';
            int x = (int)ch + dates.Count;
            ch = (char)x;
            var range1 = "Class Data!A:" + ch;

            var objectList2 = new List<object>();
            objectList2.Add("Гостиница");
            objectList2.Add("Кол-во чел.");
            objectList2.Add("ФИО");
            objectList2.Add("Номер");
            objectList2.Add("Заезд");
            objectList2.Add("Выезд");

            string numOfCategories = textBox2.Text;
            string numOfRooms = textBox5.Text;
            foreach (var i in dates)
            {
                objectList2.Add(i);
            }
            valueRange.Values = new List<IList<object>> { objectList2 };
            var appendRequest2 = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
            appendRequest2.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse2 = appendRequest2.Execute();

            range1 = "Class Data!A2:A" + numOfRooms;

            var objectList = new List<object>();
            objectList.Add(textBox1.Text);
            valueRange.Values = new List<IList<object>> { objectList };
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            for (int i = 0; i < int.Parse(numOfRooms); i++)
            {
                var appendResponse = appendRequest.Execute();
            }

            var range3 = "Class Data!D" + (int.Parse(numOfRooms) + 3) + ":" + ch;
            var objectList3 = new List<object>();
            objectList3.Add("Блок в отеле");
            g.updateRequest(range3, objectList3);           
            clickCounter = clickCounter + 3 + int.Parse(numOfRooms); 

            var range4 = "Class Data!D" + (int.Parse(numOfRooms) + 5 + int.Parse(textBox2.Text)) + ":" + ch;
            var objectList4 = new List<object>();
            objectList4.Add("Фактический блок");
            g.updateRequest(range4, objectList4);

            var range5 = "Class Data!D" + (int.Parse(numOfRooms) + 7 + int.Parse(textBox2.Text) * 2) + ":" + ch;
            var objectList5 = new List<object>();
            objectList5.Add("Разница");
            g.updateRequest(range5, objectList5);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        

        private void button2_Click(object sender, EventArgs e)
        {
            int numOfCategories = int.Parse(textBox2.Text);
            if(clickCounter == numOfCategories)
            {
                throw new Exception();
            }
            GoogleCredential credential;
            using (var stream = new FileStream("My First Project-2dfed0050064.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            logic updater = new logic();
            string categoryName = textBox6.Text;
            string roomCapacity = textBox4.Text;
            string numOfRooms = textBox3.Text;
            var objectList = new List<object>();
            objectList.Add(categoryName);
            objectList.Add(roomCapacity);
            objectList.Add("");
            var start = dateTimePicker1.Value;
            var end = dateTimePicker2.Value;
            var dates = end - start;
            char ch = 'D';
            int x = (int)ch + dates.Days + 4;
            ch = (char)x;
            for (int i = 0; i <= dates.Days + 1; i++)
            {
                objectList.Add(numOfRooms);
            }
            string range = "Class Data!D" + clickCounter + ":" + ch + clickCounter;
            updater.updateRequest(range, objectList);

            var objectList2 = new List<object>();
            objectList2.Add(categoryName);
            objectList2.Add(roomCapacity);
            objectList2.Add("");
            string range2 = "Class Data!D" + (clickCounter + numOfCategories + 2) + ":" + ch + (clickCounter + numOfCategories + 2);
            for (int i = 0; i <= dates.Days + 1; i++)
            {
                objectList2.Add("0");
            }
            updater.updateRequest(range2, objectList2);

            string range3 = "Class Data!D" + (clickCounter + numOfCategories * 2 + 4) + ":E" + (clickCounter + numOfCategories * 2 + 4);
            var objectList3 = new List<object>();
            objectList3.Add(categoryName);
            objectList3.Add(roomCapacity);
            updater.updateRequest(range3, objectList3);

            for (int j = 0; j <= dates.Days + 1; j++)
            {
                char column = 'G';
                x = (int)column + j;
                column = (char)x;
                var cell = "Class Data!" + column + (clickCounter + numOfCategories + 7) + ":" + column;

                string result = "=" + column + clickCounter + " - " + column + (clickCounter + numOfCategories + 2);

                var objectList4 = new List<object>();
                objectList4.Add(result);
                updater.updateRequest(cell, objectList4);
                }

            clickCounter++;
        }
    }
}
