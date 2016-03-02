using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb; //Read excel file as a databases

namespace ExcelToAccess
{
    public partial class Form1 : Form
    {
        OleDbConnection con;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dgShow.DataSource = null;
            dgShow.Refresh();

            //Set up the connection and put it in the data grid
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ txtFileName.Text +";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';");
            OleDbDataAdapter da = new OleDbDataAdapter("select * from [Export Worksheet$]", con);
            da.Fill(dsInfo);
            dgShow.DataSource = dsInfo.Tables[0];

            //Query the information from the worksheet
            string selectQuery = "SELECT * from [Export Worksheet$]";
            OleDbCommand myCommand = new OleDbCommand(selectQuery, con);
            myCommand.Connection.Open();

            //Read through each row and store all the person objects into an array
            var peopleResults = new List<Person>();
            OleDbDataReader myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection);

            Client cli = new Client();
            cli.initClient();
            while(myReader.Read())
            {
                Person person = new Person();
                if (myReader[@"PERSON$_CODE"] != DBNull.Value)
                {
                    person.PERSON_CODE = (double)myReader[@"PERSON$_CODE"];
                }
                if (myReader["FIRST_NAME"] != DBNull.Value)
                {
                    person.FIRST_NAME = (string)myReader["FIRST_NAME"];
                }
                if(myReader["LAST_NAME"] != DBNull.Value)
                {
                    person.LAST_NAME = (string)myReader["LAST_NAME"];
                }
                if(myReader["DOB"] != DBNull.Value)
                {
                    person.DOB = (string)myReader["DOB"];
                }
                if(myReader["ADDRESS_1"] != DBNull.Value)
                {
                    person.ADDRESS_1 = (string)myReader["ADDRESS_1"];
                }
                if (myReader["ADDRESS_2"] != DBNull.Value)
                {
                    person.ADDRESS_2 = (string)myReader["ADDRESS_2"];
                }
                if (myReader["ADDRESS_TYPE"] != DBNull.Value)
                {
                    person.ADDRESS_TYPE = (string)myReader["ADDRESS_TYPE"];
                }
                if (myReader["CITY"] != DBNull.Value)
                {
                    person.CITY = (string)myReader["CITY"];
                }
                if (myReader["STATE"] != DBNull.Value)
                {
                    person.STATE = (string)myReader["STATE"];
                }
                if (myReader["ZIP_CODE"] != DBNull.Value)
                {
                    person.ZIP_CODE = (string)myReader["ZIP_CODE"];
                }
                if (myReader["EMAIL"] != DBNull.Value)
                {
                    person.EMAIL = (string)myReader["EMAIL"];
                }
                if (myReader["TELEPHONE_1"] != DBNull.Value)
                {
                    person.TELEPHONE_1 = (string)myReader["TELEPHONE_1"];
                }
                if (myReader["ADMIT_DATE"] != DBNull.Value)
                {
                    person.ADMIT_DATE = (string)myReader["ADMIT_DATE"];
                }
                peopleResults.Add(person);
            }
            cli.uploadPatients(peopleResults);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnReadMedication_Click(object sender, EventArgs e)
        {
            dgShow.DataSource = null;
            dgShow.Refresh();
            //Set up the connection and put it in the data grid
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ txtFileName.Text +";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';");
            OleDbDataAdapter da = new OleDbDataAdapter("select * from [Export Worksheet$]", con);
            da.Fill(dsInfo);
            dgShow.DataSource = dsInfo.Tables[0];

            //Query the information from the worksheet
            string selectQuery = "SELECT * from [Export Worksheet$]";
            OleDbCommand myCommand = new OleDbCommand(selectQuery, con);
            myCommand.Connection.Open();

            //Read through each row and store all the person objects into an array
            var medicationResults = new List<excelMedication>();
            OleDbDataReader myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader.Read())
            {
                excelMedication medication = new excelMedication();
                if (myReader[@"PERSON$_CODE"] != DBNull.Value)
                {
                    medication.PERSON_CODE = (double)myReader[@"PERSON$_CODE"];
                }
                if (myReader["FIRST_NAME"] != DBNull.Value)
                {
                    medication.FIRST_NAME = (string)myReader["FIRST_NAME"];
                }
                if (myReader["LAST_NAME"] != DBNull.Value)
                {
                    medication.LAST_NAME = (string)myReader["LAST_NAME"];
                }
                if (myReader["DOB"] != DBNull.Value)
                {
                    medication.DOB = (string)myReader["DOB"];
                }
                if (myReader["MED_NAME"] != DBNull.Value)
                {
                    medication.MED_NAME = (string)myReader["MED_NAME"];
                }
                if (myReader[@"MEDICATIONS$_TEXT"] != DBNull.Value)
                {
                   medication.MEDICATION_TEXT = (string)myReader[@"MEDICATIONS$_TEXT"];
                }
                medicationResults.Add(medication);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            xxx2fhir.json2Resource(txtFileName.Text);
        }
    }

  
}
