using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hl7.Fhir.Model;
using Hl7.Fhir.Rest;

namespace ExcelToAccess
{
    class Client
    {
        private FhirClient client;

        public void initClient()
        {
            client = new FhirClient("http://localhost:49911/fhir");
        }

        public void uploadPatients(List<Person> pList)
        {
            var patientList = new List<Patient>();
            int i = 0;
            foreach(Person person in pList)
            {
                Patient patient = new Patient();
                //adding name
                var name = new HumanName();
                name.Family = new List<string> { person.LAST_NAME };
                name.Given = new List<string> { person.FIRST_NAME };
                patient.Name = new List<HumanName> { name };

                //adding address
                Address add = new Address();
                add.Text = person.ADDRESS_1 + person.ADDRESS_2;
                add.City = person.CITY;
                add.State = person.STATE;
                add.PostalCode = person.ZIP_CODE;
                patient.Address = new List<Address> { add };

                //adding phone&email
                ContactPoint phone = new ContactPoint();
                phone.System = ContactPoint.ContactPointSystem.Phone;
                phone.Value = person.TELEPHONE_1;
                ContactPoint email = new ContactPoint();
                email.System = ContactPoint.ContactPointSystem.Email;
                email.Value = person.EMAIL;
                patient.Telecom = new List<ContactPoint> { phone, email };

                patientList.Add(patient);
                try
                {
                    var patEntry = client.Create(patient);
                    string id = patEntry.Id;
                }
                catch (FhirOperationException ex)
                {
                }
                catch (Exception ex)
                {
                }
            }
            Console.WriteLine("done uploading");
        }

        public void uploadMedication(List<excelMedication> mList)
        {
            var medList = new List<MedicationOrder>();
            foreach (excelMedication med in mList)
            {
                MedicationOrder prescription = new Hl7.Fhir.Model.MedicationOrder();

                //adding who this prescription is for
                /*ResourceReference rf = new ResourceReference();
                rf.Reference = 
                Patient patient = new Patient();
                var name = new HumanName();
                name.Family = new List<string> { med.LAST_NAME };
                name.Given = new List<string> { med.FIRST_NAME };
                patient.Name = new List<HumanName> { name };
                prescription.Patient = ;*/

                //adding
                ResourceReference rf = new ResourceReference();
                


            }
        }
    }
}
