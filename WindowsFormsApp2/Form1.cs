using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Newtonsoft.Json;
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
        static readonly string sheet = "Main Table";
        static SheetsService service;
        int addedCategoies = 1;
        int addName = 2;
        int allRooms = 0;

        public Form1()
        {
            InitializeComponent();
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

            Google.Apis.Sheets.v4.Data.ClearValuesRequest requestBody = new Google.Apis.Sheets.v4.Data.ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, "Main Table!A1:Z100");
            Google.Apis.Sheets.v4.Data.ClearValuesResponse response = request.Execute();

            Google.Apis.Sheets.v4.Data.ClearValuesRequest requestBody2 = new Google.Apis.Sheets.v4.Data.ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request2 = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, "Data Table!A1:Z100");
            Google.Apis.Sheets.v4.Data.ClearValuesResponse response2 = request2.Execute();
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
            string numOfCategories = textBox2.Text;

            if (numOfCategories == "" && textBox1.Text == "")
            {
                MessageBox.Show(
                    "Все поля должны быть заполнены",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification);
                return;
            }

            var valueRange = new ValueRange();
            logic g = new logic();
            List<string> dates = new List<string>();
            g.initializeDates(start, end, ref dates);
            char ch = 'G';
            int x = (int)ch + dates.Count;
            ch = (char)x;
            var range1 = "Main Table!A:" + ch;

            var objectList2 = new List<object>();
            objectList2.Add("Гостиница");
            objectList2.Add("Кол-во чел.");
            objectList2.Add("ФИО");
            objectList2.Add("Номер");
            objectList2.Add("Заезд");
            objectList2.Add("Выезд");


            foreach (var i in dates)
            {
                objectList2.Add(i);
            }
            valueRange.Values = new List<IList<object>> { objectList2 };
            var appendRequest2 = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range1);
            appendRequest2.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse2 = appendRequest2.Execute();

            var range3 = "Data Table!D1" + ":" + ch;
            var objectList3 = new List<object>();
            objectList3.Add("Блок в отеле");
            objectList3.Add("");
            objectList3.Add("");
            foreach (var i in dates)
            {
                objectList3.Add(i);
            }
            g.updateRequest(range3, objectList3);

            var range4 = "Data Table!D" + (3 + int.Parse(numOfCategories)) + ":" + ch;
            var objectList4 = new List<object>();
            objectList4.Add("Фактический блок");
            g.updateRequest(range4, objectList4);

            var range5 = "Data Table!D" + (5 + int.Parse(numOfCategories) * 2) + ":" + ch;
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

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        { }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            var hotelName = textBox1.Text;
            if (hotelName == "")
            {
                return;
            }

            var categoryName = textBox8.Text;
            var roomCapacity = textBox7.Text;
            var roomsNum = textBox5.Text;
            logic updater = new logic();

            int res;

            if (categoryName == "" || roomCapacity == "" || roomsNum == "")
            {
                MessageBox.Show(
                    "Все поля должны быть заполнены",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification);
                return;
            }

            if (!int.TryParse(roomCapacity, out res) || !int.TryParse(roomsNum, out res))
            {
                MessageBox.Show(
                    "Ввдите корректные данные",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification);
                return;
            }

            if (addedCategoies == (int.Parse(textBox2.Text) + 1))
            {
                MessageBox.Show(
                    "Все категории заполнены",
                    "Внимание",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification);
                return;
            }

            int n = dataGridView1.Rows.Add();
            dataGridView1.Rows[n].Cells[0].Value = categoryName;
            dataGridView1.Rows[n].Cells[1].Value = roomCapacity;
            dataGridView1.Rows[n].Cells[2].Value = roomsNum;
            allRooms += int.Parse(roomsNum);
            int numOfCategories = int.Parse(textBox2.Text);

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
                objectList.Add(roomsNum);
            }
            string range = "Data Table!D" + (addedCategoies + 1) + ":" + ch + (addedCategoies + 1);
            updater.updateRequest(range, objectList);

            var objectList2 = new List<object>();
            objectList2.Add(categoryName);
            objectList2.Add(roomCapacity);
            objectList2.Add("");
            string range2 = "Data Table!D" + (addedCategoies + numOfCategories + 3) + ":" + ch + (addedCategoies + numOfCategories + 3);
            for (int i = 0; i <= dates.Days + 1; i++)
            {
                objectList2.Add("0");
            }
            updater.updateRequest(range2, objectList2);

            string range3 = "Data Table!D" + (addedCategoies + numOfCategories * 2 + 5) + ":E" + (addedCategoies + numOfCategories * 2 + 5);
            var objectList3 = new List<object>();
            objectList3.Add(categoryName);
            objectList3.Add(roomCapacity);
            updater.updateRequest(range3, objectList3);

            for (int j = 0; j <= dates.Days + 1; j++)
            {
                char column = 'G';
                x = (int)column + j;
                column = (char)x;
                var cell = "Data Table!" + column + (addedCategoies + numOfCategories * 2 + 5) + ":" + column;

                string result = "=" + column + (addedCategoies + 1) + " - " + column + (addedCategoies + numOfCategories + 3);

                var objectList4 = new List<object>();
                objectList4.Add(result);
                updater.updateRequest(cell, objectList4);
            }

            var range5 = "Main Table!A:D";
            var objectList5 = new List<object>();
            objectList5.Add(hotelName);
            objectList5.Add("");
            objectList5.Add("");
            objectList5.Add(categoryName);
            ValueRange valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>> { objectList5 };
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range5);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            for (int i = 0; i < int.Parse(roomsNum); i++)
            {
                var appendResponse = appendRequest.Execute();
            }

            addedCategoies++;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var menNum = textBox4.Text;
            var names = textBox6.Text;
            var numOfcategories = int.Parse(textBox2.Text);
            logic updater = new logic();

            int res;

            if (menNum == "" || names == "")
            {
                MessageBox.Show(
                    "Все поля должны быть заполнены",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification);
                return;
            }

            if (!int.TryParse(menNum, out res))
            {
                MessageBox.Show(
                    "Ввдите корректные данные",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification);
                return;
            }

            if (addedCategoies == (int.Parse(textBox2.Text) + 1))
            {
                MessageBox.Show(
                    "Все поля заполнены",
                    "Внимание",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification);
                return;
            }

            int n = dataGridView2.Rows.Add();
            dataGridView2.Rows[n].Cells[0].Value = menNum;
            dataGridView2.Rows[n].Cells[1].Value = names;

            logic g = new logic();
            var range1 = "Main Table!B" + addName + ":C" + addName;
            var objectList1 = new List<object>();
            objectList1.Add(menNum);
            objectList1.Add(names);
            g.updateRequest(range1, objectList1);

            var start = dateTimePicker1.Value;
            var end = dateTimePicker2.Value;
            var dates = end - start;
            char ch = 'G';
            int x = (int)ch + dates.Days;
            ch = (char)x;

            int bisy = 0;
            int sumRooms = 0;
            int j = 0;

            while (bisy >= sumRooms)
            {
                String range = "Data Table!G" + (numOfcategories + 3 + j) + "G" + (numOfcategories + 3 + j);
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(SpreadsheetId, range);
                ValueRange response = request.Execute();
                IList<IList<Object>> values = response.Values;
                bisy = int.Parse(values[0][0].ToString());

                String range2 = "Data Table!G" + (addName + j) + "G" + (addName + j);
                SpreadsheetsResource.ValuesResource.GetRequest request2 =
                        service.Spreadsheets.Values.Get(SpreadsheetId, range2);
                ValueRange response2 = request.Execute();
                IList<IList<Object>> values2 = response2.Values;
                sumRooms = int.Parse(values2[0][0].ToString());

                j++;
            }
            Console.WriteLine(bisy.ToString(), sumRooms.ToString());
            if (bisy < sumRooms)
            {
                bisy++;
                var range3 = "Data Table!G" + addName + ch + addName;
                var objectList2 = new List<object>();
                for (int i = 0; i < dates.Days; i++)
                {
                    objectList2.Add(bisy);
                }
                g.updateRequest(range1, objectList1);
            }
            else
            {
                if (addedCategoies == (int.Parse(textBox2.Text) + 1))
                {
                    MessageBox.Show(
                        "Нет свободныз мест!",
                        "Внимание",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.ServiceNotification);
                    return;
                }
            }

            var range5 = "Main Table!G:" + addName + ":" + ch + addName;

            addName++;
        }
    }
}
