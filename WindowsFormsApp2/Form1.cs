using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;

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
            if (true)
            {
                Console.WriteLine("Name, Major");
                string name = textBox1.Text;
                string gender = textBox2.Text;
                string cval = textBox3.Text;
                string state = textBox4.Text;
                string spec = textBox5.Text;
                string club = textBox6.Text;
                var range1 = "Class Data!A:F";

                var valueRange = new ValueRange();

                var objectList = new List<object>() { name, gender, cval, state, spec, club };
                valueRange.Values = new List<IList<object>> { objectList };
                var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                var appendResponse = appendRequest.Execute();

            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
