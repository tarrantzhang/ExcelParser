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
            var medicationResults = new List<Medication>();
            OleDbDataReader myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection);
            while (myReader.Read())
            {
                Medication medication = new Medication();
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
    }

    class Person
    {
        public double PERSON_CODE = 0;
        public string FIRST_NAME = "";
        public string LAST_NAME = "";
        public string DOB = "";
        public string ADDRESS_1 = "";
        public string ADDRESS_2 = "";
        public string ADDRESS_TYPE = "";
        public string CITY = "";
        public string STATE = "";
        public string ZIP_CODE = "";
        public string EMAIL = "";
        public string TELEPHONE_1 = "";
        public string ADMIT_DATE = "";
    }

    class Medication
    {
        public double PERSON_CODE = 0;
        public string FIRST_NAME = "";
        public string LAST_NAME = "";
        public string DOB = "";
        public string MED_NAME = "";
        public string MEDICATION_TEXT = "";
    }

    class ALLERGY
    {
        public double PERSON_CODE = 0;
        public string FIRST_NAME = "";
        public string LAST_NAME = "";
        public string DOB = "";
        public string ALLERGY_CODE = "";
        public string ALLERGY_TEXT = "";
        public string ALLERGY_ADDED_DATE = "";
        public string ADMIT_DATE = "";
    }

    class PROBLEM
    {
        public double PERSON_CODE = 0;
        public string ADMIT_DATE = "";
        public string PROBLEMM_CODE = "";
        public string VISIT_CODE = "";
        public string PROBLEM_TEXT = "";
        public string PROBLEM_PREV = "";
        public string PROBLEM_ACTIVITY = "";
        public string PROBLEM_ADDED_DATE = "";
        public string PROBLEM_ADD_DDX_MACRO = "";
        public string PROBLEM_ADD_DISCUSSION_MACRO = "";
        public string PROBLEM_ADJUDICATION_TEXT = "";
        public string PROBLEM_ASSESSMENT_TEXT = "";
        public string PROBLEM_CLASSIFICATION = "";
        public string PROBLEM_CODE = "";
        public string PROBLEM_COMMENT = "";
        public string PROBLEM_DEFAULT_SCRIPT = "";
        public string PROBLEM_DESCRIPTION = "";
        public string PROBLEM_DISCUSSION_FREETEXT = "";
        public string PROBLEM_END_DATE = "";
        public string PROBLEM_EXCLUDED_COMMENT = "";
        public string PROBLEM_HISTORY_TEXT = "";
        public string PROBLEM_HX_ACTIVITY = "";
        public string PROBLEM_LAST_MODIFIED_DATE = "";
        public string PROBLEM_NOTE = "";
        public string PROBLEM_OLD_REVISION = "";
        public string PROBLEM_ONSET_DATE = "";
        public string PROBLEM_ORIENTATION_TEXT = "";
        public string PROBLEM_PLANS_TEXT = "";
        public string PROBLEM_REPORT_SHOW = "";
        public string PROBLEM_SHOW_IN_PROBLEMS = "";
        public string PROBLEM_SOURCE = "";
        public string PROBLEM_STATUS = "";
        public string PROBLEM_STATUS1_SUBJECTIVE = "";
        public string PROBLEM_STATUS2 = "";
        public string PROBLEM_STATUS2_SUBJECTIVE = "";
        public string PROBLEM_USER_ID_ADDED = "";
        public string PROBLEM_USER_ID_LAST_MODIF = "";
        public string TAG_ORDER = "";
        public string TAG_CREATEDATE = "";
        public string TAG_SYSTEMDATE = "";
        public string TAG_SYSTEMUSER = "";
        public string PROBLEM_DISCUSSION_TOPIC = "";
        public string PROBLEM_HISTORY_TOPIC = "";
        public string PROBLEM_ONSET_DATE_PARTIAL = "";
        public string PROBLEM_STATUS2_FLAG = "";
        public string PROBLEM_STATUS2_COPY = "";
        public string PROBLEM_DISCUSSION_DIFFERENTIA = "";
        public string PROBLEM_DURATION_TEXT = "";
        public string PROBLEM_SEVERITY_FLAG = "";
    }
}
