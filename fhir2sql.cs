using Hl7.Fhir.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAccess
{
    class fhir2sql
    {
        public string[] prepareImportPatient(Patient patient, string consentDate,string population)
        {
            string[] rtvl = new string[5];
            string id = patient.Identifier.ElementAt(0).Value;
            string firstName = patient.Name.ElementAt(0).Given.First();
            string site = patient.Address.First().City;
            rtvl[0] = id;
            rtvl[1] = firstName;
            rtvl[2] = consentDate;
            rtvl[3] = population;
            rtvl[4] = site;
            return rtvl;
        }

        public string[] prepareCreatePopulation(string population, string IRB_approve, string IRB_expire, string min_trigger, string max_trigger,string no_connect_trigger)
        {
            return new string[6] { population, IRB_approve, IRB_expire, min_trigger, max_trigger, no_connect_trigger };
        }
        public string[] prepareAssignPatientPopulation(string patientID, string SiteID,string populationID,string consentdate)
        {
            return new string[4] { patientID, SiteID, populationID, consentdate};
        }

        public string[] prepareImportPrescription(MedicationOrder mo, string population, )
        {
            ResourceReference medRef = (ResourceReference) mo.Medication;
            string reference = medRef.Reference;
        }
    }
}
