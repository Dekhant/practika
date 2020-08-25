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

            String range = "Class Data!A:F";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
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
            objectList2.Add("Группа");
            objectList2.Add("Погоняло");
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
            valueRange.Values = new List<IList<object>> { objectList3 };
            var appendRequest3 = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range3);
            appendRequest3.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse3 = appendRequest3.Execute();
            clickCounter = clickCounter + 3 + int.Parse(numOfRooms); 
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

            string categoryName = textBox6.Text;
            string roomCapacity = textBox4.Text;
            string numOfRooms = textBox3.Text;
            var objectList = new List<object>();
            objectList.Add(textBox6.Text);
            objectList.Add(textBox4.Text);
            objectList.Add("");
            var start = dateTimePicker1.Value;
            var end = dateTimePicker2.Value;
            var dates = end - start;
            char ch = 'D';
            int x = (int)ch + dates.Days + 4;
            ch = (char)x;
            for (int i = 0; i <= dates.Days + 1; i++)
            {
                objectList.Add(roomCapacity);
            }
            string range = "Class Data!D" + clickCounter + ":" + ch + clickCounter;
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { objectList };
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            clickCounter++;
        }
    }
}
