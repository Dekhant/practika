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
            var start = dateTimePicker1.Value;
            var end = dateTimePicker2.Value;

            var valueRange = new ValueRange();
            logic g = new logic();
            List<string> dates = new List<string>();
            g.initializeDates(start, end, ref dates);
            char ch = 'A';
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
            var numOfCategory = textBox3.Text.Split(' ');
            var roomCapacity = textBox4.Text.Split(' ');
            foreach (var i in dates)
            {
                objectList2.Add(i);
            }
            valueRange.Values = new List<IList<object>> { objectList2 };
            var appendRequest2 = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
            appendRequest2.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse2 = appendRequest2.Execute();

            range1 = "Class Data!A2:A" + g.tableMarkup(numOfCategories, numOfCategory, roomCapacity);

            var objectList = new List<object>();
            objectList.Add(textBox1.Text);
            for (int i = 0; i < g.tableMarkup(numOfCategories, numOfCategory, roomCapacity); i++)
            {
                valueRange.Values = new List<IList<object>> { objectList };
                var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
