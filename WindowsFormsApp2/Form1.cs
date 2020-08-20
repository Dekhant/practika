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
            string name = textBox1.Text;
            string category = textBox2.Text;
            string start = textBox3.Text;
            string end = textBox4.Text;

            var valueRange = new ValueRange();
            logic g = new logic();
            List<string> dates = new List<string>();
            g.CreateDates(start, end, ref dates);
            char ch = 'A';
            int x = (int)ch + dates.Count;
            ch = (char)x;
            var range1 = "Class Data!A:" + ch;

            var objectList2 = new List<object>();
            objectList2.Add("название");
            objectList2.Add("кол-во категорий");
            foreach (var i in dates)
            {
                objectList2.Add(i);
            }
            valueRange.Values = new List<IList<object>> { objectList2 };
            var appendRequest2 = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
            appendRequest2.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse2 = appendRequest2.Execute();

            var objectList = new List<object>() { name, category };
            valueRange.Values = new List<IList<object>> { objectList };
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
