using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Xml;
using System.Xml.Serialization;

namespace Bulk_Report_Tool_V3
{
    public class FullCollection
    {
        public string AssetTag;
        public string AssetMakeModel;
        public string SerialNumber;
        public string Location;
        public string Owner;
        public string Status;
        public string TaxLocation;
        public string Created;
        public string JSSID;
        public string JSSComputerName;
        public string JSSManaged;
        public string JSSUsername;
        public string JSSComputerModel;
        public string JSSDepartment;
        public string JSSMACAddress;
        public string JSSUDID;
        public string JSSSerialNumber;
        public string JSSReportDate;
        public string JiraAssetJSON;
        public string AssetType;
        bool HasOwner;
        public FullCollection(string JiraAssetJSON1)
        {
            JiraAssetJSON = JiraAssetJSON1;
            //Create Collection based on Asset Insight data
            PopulateJiraData();
        }
        private void PopulateJiraData()
        {
            var jPerson = JsonConvert.DeserializeObject<dynamic>(JiraAssetJSON);
            if (jPerson.objectEntries[0].objectType.name == "Laptop")
            {
                //Assign Laptop asset data from Insight json
                AssetType = "Laptop";
                try { AssetTag = jPerson.objectEntries[0].attributes[11].objectAttributeValues[0].displayValue; } catch { AssetTag = "Unknown"; }
                try { SerialNumber = jPerson.objectEntries[0].attributes[13].objectAttributeValues[0].displayValue; } catch { SerialNumber = "Unknown"; }
                try { Location = jPerson.objectEntries[0].attributes[5].objectAttributeValues[0].displayValue; } catch { Location = "Unknown"; }
                try { Owner = jPerson.objectEntries[0].attributes[4].objectAttributeValues[0].displayValue;
                    if (this.Owner == "#N/A" || this.Owner == "Not Listed.") { HasOwner = true; } } catch { Owner = "Unknown"; }
                try { Status = jPerson.objectEntries[0].attributes[12].objectAttributeValues[0].displayValue; } catch { Status = "Unknown"; }
                try { TaxLocation = jPerson.objectEntries[0].attributes[7].objectAttributeValues[0].displayValue; } catch { TaxLocation = "Unknown"; }
                try { Created = jPerson.objectEntries[0].attributes[1].objectAttributeValues[0].displayValue; } catch { Created = "Unknown"; }
                try { AssetMakeModel = jPerson.objectEntries[0].attributes[9].objectAttributeValues[0].displayValue; } catch { AssetMakeModel = "Unknown"; }
            }
            else if (jPerson.objectEntries[0].objectType.name == "Thin Clients")
            {
                //Assign Thin Clients asset data from Insight json
                AssetType = "Thin Clients";
                try
                { 
                    string test = jPerson.objectEntries[0].attributes[1].objectAttributeValues[0].displayValue;
                    test = test.Split(' ')[0].Trim();
                    AssetTag = test;
                } 
                catch { AssetTag = "Unknown"; }
                try { SerialNumber = jPerson.objectEntries[0].attributes[14].objectAttributeValues[0].displayValue; } catch { SerialNumber = "Unknown"; }
                try { Location = jPerson.objectEntries[0].attributes[5].objectAttributeValues[0].displayValue; } catch { Location = "Unknown"; }
                try { Owner = jPerson.objectEntries[0].attributes[4].objectAttributeValues[0].displayValue;
                    if (this.Owner == "#N/A" || this.Owner == "Not Listed.") { HasOwner = true; } } catch { Owner = "Unknown"; }
                try { Status = jPerson.objectEntries[0].attributes[13].objectAttributeValues[0].displayValue; } catch { Status = "Unknown"; }
                try { TaxLocation = jPerson.objectEntries[0].attributes[8].objectAttributeValues[0].displayValue; } catch { TaxLocation = "Unknown"; }
                try { Created = jPerson.objectEntries[0].attributes[2].objectAttributeValues[0].displayValue; } catch { Created = "Unknown"; }
                try { AssetMakeModel = jPerson.objectEntries[0].attributes[10].objectAttributeValues[0].displayValue; } catch { AssetMakeModel = "Unknown"; }
            }
            else if (jPerson.objectEntries[0].objectType.name == "Network")
            {
                //Assign Network asset data from Insight json
                AssetType = "Network";
                try
                {
                    string test = jPerson.objectEntries[0].attributes[1].objectAttributeValues[0].displayValue;
                    test = test.Split(' ')[0].Trim();
                    AssetTag = test;
                }
                catch { AssetTag = "Unknown"; }
                try { SerialNumber = jPerson.objectEntries[0].attributes[6].objectAttributeValues[0].displayValue; } catch { SerialNumber = "Unknown"; }
                try { Location = jPerson.objectEntries[0].attributes[10].objectAttributeValues[0].displayValue; } catch { Location = "Unknown"; }
                try { Owner = jPerson.objectEntries[0].attributes[23].objectAttributeValues[0].displayValue; } catch { Owner = "Unknown"; }//Primary Contact
                try { Status = jPerson.objectEntries[0].attributes[20].objectAttributeValues[0].displayValue; } catch { Status = "Unknown"; }
                try { TaxLocation = jPerson.objectEntries[0].attributes[9].objectAttributeValues[0].displayValue; } catch { TaxLocation = "Unknown"; }
                try { Created = jPerson.objectEntries[0].attributes[2].objectAttributeValues[0].displayValue; } catch { Created = "Unknown"; }
                try { AssetMakeModel = jPerson.objectEntries[0].attributes[9].objectAttributeValues[0].displayValue; } catch { AssetMakeModel = "Unknown"; }
            }
            else if (jPerson.objectEntries[0].objectType.name == "Tablet")
            {
                //Assign Tablet asset data from Insight json
                AssetType = "Tablet";
                try { AssetTag = jPerson.objectEntries[0].attributes[11].objectAttributeValues[0].displayValue; } catch { AssetTag = "Unknown"; }
                try { SerialNumber = jPerson.objectEntries[0].attributes[13].objectAttributeValues[0].displayValue; } catch { SerialNumber = "Unknown"; }
                try { Location = jPerson.objectEntries[0].attributes[5].objectAttributeValues[0].displayValue; } catch { Location = "Unknown"; }
                try { Owner = jPerson.objectEntries[0].attributes[4].objectAttributeValues[0].displayValue; } catch { Owner = "Unknown"; }
                try { Status = jPerson.objectEntries[0].attributes[12].objectAttributeValues[0].displayValue; } catch { Status = "Unknown"; }
                try { TaxLocation = jPerson.objectEntries[0].attributes[7].objectAttributeValues[0].displayValue; } catch { TaxLocation = "Unknown"; }
                try { Created = jPerson.objectEntries[0].attributes[1].objectAttributeValues[0].displayValue; } catch { Created = "Unknown"; }
                try { AssetMakeModel = jPerson.objectEntries[0].attributes[10].objectAttributeValues[0].displayValue; } catch { AssetMakeModel = "Unknown"; }
            }
            else
            {
                AssetType = "Unknown Asset Type";
                AssetTag = "Not a Laptop/Thin Client/Network";
                SerialNumber = "Not a Laptop/Thin Client/Network";
                Location = "Not a Laptop/Thin Client/Network";
                Owner = "Not a Laptop/Thin Client/Network";
                Status = "Not a Laptop/Thin Client/Network";
                AssetTag = "Not a Laptop/Thin Client/Network";
                TaxLocation = "Not a Laptop/Thin Client/Network";
            }
        }
        public void PopulateJSSData(string JssFullData)
        {
            string Name = SerialNumber.ToUpper();
            List<string> AssetData = new List<string>();
            Name = Name.ToUpper();
            if (JssFullData != null && JssFullData.Length >= 3)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(JssFullData);
                foreach (XmlNode node in doc.DocumentElement)
                {
                    if (node.Name == "computer")
                    {
                        XmlNodeList thisNode = node.SelectNodes("serial_number");
                        string TempCheck = thisNode.Item(0).InnerText;
                        if (TempCheck == Name)
                        {
                            foreach (XmlNode child in node.ChildNodes)
                            {
                                if (child.Name == "id") { JSSID = child.InnerText; }
                                if (child.Name == "name") { JSSComputerName = child.InnerText; }
                                if (child.Name == "managed") { JSSManaged = child.InnerText; }
                                if (child.Name == "username") { JSSUsername = child.InnerText; }
                                if (child.Name == "model") { JSSComputerModel = child.InnerText; }
                                if (child.Name == "department") { JSSDepartment = child.InnerText; }
                                if (child.Name == "mac_address") { JSSMACAddress = child.InnerText; }
                                if (child.Name == "udid") { JSSUDID = child.InnerText; }
                                if (child.Name == "serial_number") { JSSSerialNumber = child.InnerText; }
                                if (child.Name == "report_date_utc") { JSSReportDate = child.InnerText; }
                            }
                        }
                    }
                }
            }
            else
            {

            }
        }
        public string CreateAssetWarnings(bool JssData)
        {
            string ToReturn = "";
            //Issues for every machine:
            if (this.Location.Contains("WITH USER") && (this.Owner == "#N/A" || this.Owner == "Not Listed." || this.Owner == "Unknown"))
            { ToReturn += $"{this.AssetTag} - Has no owner, but is listed as WITH USER\r\n"; }
            if (!this.Location.Contains("WITH USER") && (this.Owner != "#N/A" || this.Owner != "Not Listed." || this.Owner == "Unknown"))
            { ToReturn += $"{this.AssetTag} - Has owner, but is not listed as WITH USER\r\n"; }
            //--------------------
            if (JssData == true)
            {
                //Issues if Jss Data is present:
                if ((this.AssetMakeModel.Contains("Mac") || this.AssetMakeModel.Contains("mac")) && this.JSSComputerName == null)
                { ToReturn += $"{this.AssetTag} - Laptop is a Mac, but no JSS Data Exists\r\n"; }
                if (this.AssetMakeModel.Contains("Mac") || this.AssetMakeModel.Contains("mac"))
                {
                    if (this.SerialNumber != this.JSSSerialNumber)
                    { ToReturn += $"{this.AssetTag} - Jira Serial Number is not the same as JSS Serial Number\r\n"; }
                    if (this.JSSComputerName != this.JSSSerialNumber)
                    { ToReturn += $"{this.AssetTag} - JSS Computer Name is not the same as JSS Serial Number\r\n"; }
                }
                //--------------------
            }
            else
            {
                //If no Jss Data is present:
                
                //--------------------
            }
            return ToReturn;
        }
    }
}
