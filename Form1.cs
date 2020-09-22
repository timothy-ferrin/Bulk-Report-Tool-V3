using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Newtonsoft.Json;
using System.Xml;
using System.Xml.Serialization;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.Web;
using System.Reflection;
using System.Globalization;
using System.Text.RegularExpressions;
using Slack.Webhooks;

namespace Bulk_Report_Tool_V3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private List<string> GenericAssetList = new List<string>();
        private List<string> AssetListType = new List<string>();
        private List<FullCollection> FullCollectionList = new List<FullCollection>();
        public string JssToken = "";
        public string JssAssetData = "";
        public string SlackData = "";


        private void AssetCheck_Click(object sender, EventArgs e)
        {
            AssetWarningBox.Text = "";
            if (JiraUsernameBox.Text != "" && JiraPasswordBox.Text != "")
            {
                //empty lists
                GenericAssetList.Clear();
                AssetListType.Clear();
                AssetOutputBox.Text = "";
                //populate GenericAssetList
                foreach (string s in AssetInputBox.Text.Split('\n'))
                {
                    GenericAssetList.Add(s.Trim());
                }
                //populate AssetListType
                FindAssetType();
                //If asset type is Username, populate GenericAssetList with the assets that belong to the Username (by Asset Tag)
                //Then repopulate GenericAssetList
                if (GenericAssetList.Count == AssetListType.Count)
                {
                    CollectUsernameByFullname();
                }
                else
                { AssetOutputBox.Text += "Cannot run CollectUsernameByFullname - Count of Asset List and Type List is not equal!\r\n"; }
                if (GenericAssetList.Count == AssetListType.Count)
                {
                    CollectAssetTagsByUsername();
                }
                else
                {
                    AssetOutputBox.Text += "Cannot run CollectAssetTagsByUsername - Count of Asset List and Type List is not equal!\r\n";
                }
                //If asset type is Serial Number, populate GenericAssetList with the assets that belong to the Serial Number (by Asset Tag)
                //Then repopulate GenericAssetList
                if (GenericAssetList.Count == AssetListType.Count)
                {
                    CollectAssetTagsBySerialNumber();
                }
                else
                {
                    AssetOutputBox.Text += "Cannot run CollectAssetTagsBySerialNumber - Count of Asset List and Type List is not equal!\r\n";
                }
                //Depending on what type of input, determine the following:
                //Insight information.  Assign all information to an 'FullCollection' class, Create/add to a list<FullCollection>.
                //JSS information, if applicable.  Will also be a part of 'FullCollection' class.
                //AD information for the user.  Add to the 'FullCollection' class.
                if (GenericAssetList.Count == AssetListType.Count)
                {
                    for (int i = 0; i < GenericAssetList.Count; i++)
                    {
                        CreateCollection(i);
                    }
                }
                else
                {
                    AssetOutputBox.Text = "Asset List Error!  Count of Asset List and Type List is not equal!\r\n";
                }
                //Post FullCollectionList collection in a listbox for selection
                //Create an export based on the list<FullCollection>.
                AssetExportBox.Text = "Asset Tag|Serial Number|Location|Owner|Status|Tax Location|Created|Asset Type|Asset Make/Model|";
                if (JssToken != "" && !(JssToken.Contains("(401) Unauthorized")))
                {
                    AssetExportBox.Text += "JSS ID|JSS Computer Name|Managed by JSS|JSS Username|Jss Computer Model|JSS Department|JSS MAC Address|JSS UDID|JSS Serial Number|" +
                        "JSS Report Date\r\n";
                }
                else { AssetExportBox.Text += "\r\n"; }
                foreach (FullCollection s in FullCollectionList)
                {
                    AssetExportBox.Text += $"{s.AssetTag}|{s.SerialNumber}|{s.Location}|{s.Owner}|{s.Status}|{s.TaxLocation}|{s.Created}|{s.AssetType}|{s.AssetMakeModel}|";
                    if (JssToken != "" && !(JssToken.Contains("(401) Unauthorized")))
                    {
                        AssetExportBox.Text += $"{s.JSSID}|{s.JSSComputerName}|{s.JSSManaged}|{s.JSSUsername}|{s.JSSComputerModel}|{s.JSSDepartment}|" +
                        $"{s.JSSMACAddress}|{s.JSSUDID}|{s.JSSSerialNumber}|{s.JSSReportDate}\r\n";
                    }
                    else { AssetExportBox.Text += "\r\n"; }
                }
                //Create a list of 'Flagged' characteristics based on the list<FullCollection>.
                AssetWarningBox.Text += CollectAssetWarnings();
            }
            else
            {
                AssetOutputBox.Text = "Populate Jira Username and Jira Password boxes!";
            }
        }
        private void CollectUsernameByFullname()
        {
            List<string> NewGenericAssetList = new List<string>();
            List<string> NewAssetListType = new List<string>();
            for (int i = 0; i < GenericAssetList.Count(); i++)
            {
                if (AssetListType[i] == "Full Name")
                {
                    string UsernameList = "";
                    if (GenericAssetList[i].Trim() == "")
                    {
                        NewGenericAssetList.Add(GenericAssetList[i]);
                    }
                    else
                    {
                        string QuickToScript = $"$Item = \"{GenericAssetList[i].Trim()}\"";
                        QuickToScript += "\r\n$User = Get-ADUser -Filter{ displayName -like $Item } -Properties SamAccountName";
                        QuickToScript += "\r\n$User.SamAccountName";
                        string QuickUsername = RunPSReturnStr(QuickToScript);
                        //$Item = "Timothy Ferrin"
                        //$User = Get - ADUser - Filter{ displayName - like $Item} -Properties SamAccountName
                        //$User.SamAccountName


                        if (QuickUsername.Trim() == "")
                        {
                            string FirstName = GenericAssetList[i].Trim().Split(' ').First().Trim();
                            string LastName = GenericAssetList[i].Trim().Split(' ').Last().Trim();
                            string ToScript = $"$Item = \"{LastName}*\"";
                            ToScript += $"\r\n$Item2 = \"{FirstName}*\"";
                            ToScript += "\r\n$User = Get-ADUser -Filter{sn -like $Item -and givenname -like $Item2}";
                            ToScript += "\r\n$user.name, $user.SamAccountName";
                            string results = RunPSReturnStr(ToScript);
                            string MyName = results.Split('\n')[0].Trim();
                            string UserName = results.Split('\n')[1].Trim();
                            if (UserName.Contains(" "))
                            {
                                //UsernameList += $"Too many possibilities for {GenericAssetList[i].Trim()}.\r\n";
                                string[] tempFullNames1 = MyName.Split(' ');
                                List<string> tempFullNames = new List<string>();
                                for (int j = 0; j < tempFullNames1.Count(); j++)
                                {
                                    if (j % 2 == 0)
                                    {
                                        tempFullNames.Add($"{tempFullNames1[j]} {tempFullNames1[j + 1]}");
                                    }
                                }
                                string[] tempUsernames = UserName.Split(' ');
                                for (int j = 0; j < tempUsernames.Count(); j++)
                                {
                                    if (MessageBox.Show($"Verify: {GenericAssetList[i].Trim()}: {tempFullNames[j]}?  ({tempUsernames[j]})", "Name Selection", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                    {
                                        UsernameList = $"{tempUsernames[j]}";
                                    }
                                }
                            }
                            else if (MessageBox.Show($"Verify: {GenericAssetList[i].Trim()}: {MyName}?  ({UserName})", "Name Selection", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                UsernameList = $"{UserName}";
                            }
                            else
                            {
                                UsernameList = $"{GenericAssetList[i]}: Username not found.";
                            }
                        }
                        else
                        {
                            UsernameList = QuickUsername;
                        }
                    }
                    NewGenericAssetList.Add(UsernameList.Trim());
                    NewAssetListType.Add("Username");
                }
                else
                {
                    NewGenericAssetList.Add(GenericAssetList[i]);
                    NewAssetListType.Add(AssetListType[i]);
                }
            }
            GenericAssetList = NewGenericAssetList;
            AssetListType = NewAssetListType;
        }
        public string RunPSReturnStr(string MyScript)
        {
            //runs ps script and returns output
            string MyString = "";

            PowerShell ps = PowerShell.Create();
            ps.AddScript(MyScript);
            try
            {
                Collection<PSObject> PSOutput = ps.Invoke();

                string AllInfo = "";
                string TempInfo = "";
                foreach (PSObject outputItem in PSOutput)
                {
                    TempInfo = Convert.ToString(outputItem);
                    AllInfo += TempInfo + "\r\n";
                }

                MyString += AllInfo;
            }
            catch (Exception ex)
            {
                MyString += "Script Failed";
            }

            return MyString;
        }
        private void FindAssetType()
        {
            //Function to determine what type of input each line is, each input is stored
            //All numbers OR 1 A__ or L__ with all numbers = Asset Tag
            //All Letters = Username
            //1 letter followed by all numbers = ask if Asset Tag, otherwise is Computer Name
            //All letters followed by 1 number = ask if Username, otherwise is Computer Name
            //All other cases = Computer Name
            //---Possibly create a case for other things, such as Security Group, SCCM Collection, JSS Policy---//
            foreach (string s in GenericAssetList)
            {
                int LetterCounter = Regex.Matches(s, @"[a-zA-Z]").Count;
                int NumberCounter = Regex.Matches(s, @"[0-9]").Count;
                int DashCounter = Regex.Matches(s, @"-").Count;
                int SpaceCounter = Regex.Matches(s, @" ").Count;
                if (SpaceCounter > 0)
                {
                    AssetListType.Add("Full Name");
                }
                else if (LetterCounter < 1)
                {
                    AssetListType.Add("Asset Tag");
                }
                else if (NumberCounter < 1 && DashCounter < 2)
                {
                    AssetListType.Add("Username");
                }
                else if (NumberCounter < 1 && DashCounter > 1)
                {
                    AssetListType.Add("Serial Number");
                }
                else if (LetterCounter == 1)
                {
                    if (s.IndexOf('L') == 0 || s.IndexOf('A') == 0 || s.IndexOf('l') == 0 || s.IndexOf('a') == 0)
                    {
                        AssetListType.Add("Asset Tag");
                    }
                    else
                    {
                        AssetListType.Add("Serial Number");
                    }
                }
                else if (NumberCounter == 1)
                {
                    if (Regex.IsMatch(s, @"[0-9]" + @"$") && DashCounter < 2)
                    {
                        AssetListType.Add("Username");
                    }
                    else
                    {
                        AssetListType.Add("Serial Number");
                    }
                }
                else
                {
                    AssetListType.Add("Serial Number");
                }
            }
        }
        private void CollectAssetTagsByUsername()
        {
            //remakes the GenericAssetList and AssetListType lists, getting Asset Tags from Usernames.
            List<string> NewGenericAssetList = new List<string>();
            List<string> NewAssetListType = new List<string>();
            for(int i = 0; i < GenericAssetList.Count(); i++)
            {
                string strResponse = "";
                if (AssetListType[i] == "Username")
                {
                    string MyURL = $"https://jira.tlcinternal.com/rest/insight/1.0/iql/objects?objectSchemaId=1&iql=objectobject%20HAVING%20outboundReferences(" + '\u0022' + "Username" + '\u0022' + "%20=%20" + GenericAssetList[i] + ")";
                    strResponse = JiraGetRequest(MyURL);
                    try
                    {
                        var jPerson = JsonConvert.DeserializeObject<dynamic>(strResponse);
                        int count = 0;
                        for (int j = 0; j < Convert.ToInt32(jPerson.toIndex); j++)
                        {
                            try
                            {
                                //AssetOutputBox.Text += jPerson.objectEntries[j].attributes[11].objectAttributeValues[0].displayValue + "\r\n";
                                //var test = jPerson.objectEntries;
                                //if (test.Count > 0)
                                //{
                                //NewGenericAssetList.Add(Convert.ToString(jPerson.objectEntries[j].attributes[11].objectAttributeValues[0].displayValue));
                                if (jPerson.objectEntries[j].objectType.name != "Employee" && jPerson.objectEntries[j].objectType.name != "Phone" && jPerson.objectEntries[j].objectType.name != "Project")
                                {
                                    string temp1 = Convert.ToString(jPerson.objectEntries[j].label);
                                    temp1 = temp1.Split(' ')[0].Trim();
                                    NewGenericAssetList.Add(temp1);
                                    NewAssetListType.Add("Asset Tag");
                                    count++;
                                }
                                //}
                            }
                            catch { };
                        }
                        if (count == 0)
                        {
                            AssetWarningBox.Text += $"Unable to find relevant asset for {GenericAssetList[i]}.\r\n";
                        }
                    }
                    catch (Exception ex)
                    {
                        AssetOutputBox.Text += $"We had a problem procuring asset tags from {GenericAssetList[i]}: " + ex.Message.ToString() + "\r\n";
                    }
                }
                else
                {
                    NewGenericAssetList.Add(GenericAssetList[i]);
                    NewAssetListType.Add(AssetListType[i]);
                }
            }
            GenericAssetList = NewGenericAssetList;
            AssetListType = NewAssetListType;
        }
        private void CollectAssetTagsBySerialNumber()
        {
            //remakes the GenericAssetList and AssetListType lists, getting Asset Tags from Serial Numbers.
            List<string> NewGenericAssetList = new List<string>();
            List<string> NewAssetListType = new List<string>();
            for (int i = 0; i < GenericAssetList.Count(); i++)
            {
                if (AssetListType[i] == "Serial Number")
                {
                    AssetOutputBox.Text = "";
                    string searchAssets = GenericAssetList[i];
                    string searchAssets2 = "";
                    string assetidstring = " " + '\u0022' + "Serial Number" + '\u0022' + " = ";
                    foreach (string s in searchAssets.Split('\n'))
                    {
                        searchAssets2 += assetidstring + s.Trim() + " OR";
                    }
                    searchAssets2 += "DER BY Name ASC";
                    string MyURL = $"https://jira.tlcinternal.com/rest/insight/1.0/iql/objects?objectSchemaId=1&iql=" + searchAssets2;
                    string strResponse = JiraGetRequest(MyURL);
                    var jPerson = JsonConvert.DeserializeObject<dynamic>(strResponse);
                    try { NewGenericAssetList.Add(Convert.ToString(jPerson.objectEntries[0].attributes[11].objectAttributeValues[0].displayValue)); }
                        catch {
                            try
                            {
                                string toAdd = Convert.ToString(jPerson.objectEntries[0].attributes[1].objectAttributeValues[0].displayValue);
                                toAdd = toAdd.Split(' ')[0].Trim();
                                NewGenericAssetList.Add(toAdd);
                            }
                            catch { }
                        }
                    NewAssetListType.Add("Asset Tag");
                }
                else
                {
                    NewGenericAssetList.Add(GenericAssetList[i]);
                    NewAssetListType.Add(AssetListType[i]);
                }
            }
            GenericAssetList = NewGenericAssetList;
            AssetListType = NewAssetListType;
        }
        private void CreateCollection(int indexNum)
        {
            AssetOutputBox.Text += $"{GenericAssetList[indexNum]} - {AssetListType[indexNum]}\r\n";
            //creates the entire collection
            AssetListBox.Items.Add(GenericAssetList[indexNum]);
            //Insight data from Asset Tag
            string strResponse = "";
            string searchAssets = GenericAssetList[indexNum];
            string searchAssets2 = "";
            string assetidstring = " " + '\u0022' + "Asset Id" + '\u0022' + " = ";
            string assetidstring2 = " " + '\u0022' + "Name" + '\u0022' + " LIKE ";
            foreach (string s in searchAssets.Split('\n'))
            { searchAssets2 += assetidstring + s.Trim() + " OR"; }
            searchAssets2 += "DER BY Name ASC";
            RestClient rClient = new RestClient();
            try
            {
                strResponse = JiraGetRequest($"https://jira.tlcinternal.com/rest/insight/1.0/iql/objects?objectSchemaId=1&iql=" + searchAssets2);
                var jPerson = JsonConvert.DeserializeObject<dynamic>(strResponse);
                if (jPerson.objectEntries.Count > 0)
                {
                    FullCollectionList.Add(new FullCollection(strResponse));
                }
                else
                {
                    strResponse = JiraGetRequest($"https://jira.tlcinternal.com/rest/insight/1.0/iql/objects?objectSchemaId=1&iql=" + assetidstring2 + searchAssets);
                    jPerson = JsonConvert.DeserializeObject<dynamic>(strResponse);
                    if (jPerson.objectEntries.Count > 0)
                    {
                        FullCollectionList.Add(new FullCollection(strResponse));
                    }
                    else
                    {

                    }
                }
            }
            catch
            {
            }
            //Check if JSS data is applicable, get JSS data from Serial Number
            if (JssToken == "" || JssToken.Contains("(401) Unauthorized"))
            {
                //No token, no JSS data
            }
            else
            {
                if (FullCollectionList[indexNum].AssetMakeModel.Contains("mac") || FullCollectionList[indexNum].AssetMakeModel.Contains("Mac"))
                {
                    FullCollectionList[indexNum].PopulateJSSData(JssAssetData);
                }
            }
            //AD data by owner of asset
            //Adds collection to FullCollectionList
        }
        public string retrieveAllJSSAssetData()
        {
            string toReturn = "";
            string JSSURL = "https://lendingclub.jamfcloud.com/JSSResource/computers/subset/basic";
            RestClient rClient = new RestClient();
            rClient.endPoint = JSSURL;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JSSUsernameBox.Text;
            rClient.userPassword = JSSPasswordBox.Text;
            string strResponse = string.Empty;
            strResponse = rClient.makeRequest();
            toReturn = strResponse;
            return toReturn;
        }
        public string JiraGetRequest(string URL)
        {
            RestClient rClient = new RestClient();
            rClient.endPoint = URL;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;
            string strResponse = string.Empty;
            strResponse = rClient.makeRequest();
            return strResponse;
        }

        private void ClearAssetInfo_Click(object sender, EventArgs e)
        {
            AssetOutputBox.Text = "";
            AssetWarningBox.Text = "";
            AssetExportBox.Text = "";
            AssetListBox.Items.Clear();
            FullCollectionList.Clear();
        }

        private void CreateJSSToken_Click(object sender, EventArgs e)
        {
            //create bearer token for jssresource use
            string MyToken = "https://lendingclub.jamfcloud.com/uapi/auth/tokens";
            RestClient rClient = new RestClient("POST");
            rClient.endPoint = MyToken;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JSSUsernameBox.Text;
            rClient.userPassword = JSSPasswordBox.Text;
            string strResponse = string.Empty;
            strResponse = rClient.makeRequest();
            JssToken = strResponse;
            AssetOutputBox.Text = "JSS Token Created";
            JssAssetData = retrieveAllJSSAssetData();
        }
        private string CollectAssetWarnings()
        {
            bool JSSData = true;
            string ToReturn = "";
            if (JssToken == "" || JssToken.Contains("(401) Unauthorized"))
            {
                JSSData = false;
                ToReturn += $"JSS Data not acquired.  Disregard JJS-Based Analysis.\r\n\r\n";
            }
            foreach (FullCollection s in FullCollectionList)
            {
                string temp = s.CreateAssetWarnings(JSSData);
                if (temp.Trim() != "")
                {
                    ToReturn += $"{temp}\r\n";
                }
            }
            return ToReturn;
        }

        private void ADQuickExport_Click(object sender, EventArgs e)
        {
            string[] usernames = ADInputBox.Text.Split('\r');
            string QuickExport = "Username|Department|Email Address|Employee ID|Full Name|Locked?|Last Failed Login|Last Login|Manager|" +
                "Office Location|Password Expires|Password Last Changed|Title|OU Path\r\n";
            string temp = "";
            foreach (string username in usernames)
            {
                QuickExport += $"{username.Trim()}|";
                temp = GetDSInfo(username.Trim(), "Department").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "EmailAddress").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "EmployeeID").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "FullName").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "IsAccountLocked").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "LastFailedLogin").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "LastLogin").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "Manager").Trim();
                try { temp = temp.Split(',')[0].Split('=')[1]; } catch { }
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "OfficeLocations").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "PasswordExpirationDate").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "PasswordLastChanged").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "Title").Trim();
                QuickExport += $"{temp}|";
                temp = GetDSInfo(username.Trim(), "AdsPath").Trim();
                QuickExport += $"{temp}|\r\n";
                }
            ADOutputBox.Text = QuickExport;
        }

        private void ADFullExport_Click(object sender, EventArgs e)
        {

        }

        private void ADRetrieveUserGroups_Click(object sender, EventArgs e)
        {
            if (ADInputBox.Text != "")
            {
                string MyScript = "";
                string GroupNames = "";
                string temptext = "";
                string ToText = $"--------------------\r\n\r\n";
                ADUserListBox.Items.Clear();
                foreach (string s in ADInputBox.Text.Split('\n'))
                {
                    ADUserListBox.Items.Add(s.Trim());
                    MyScript = $"(New-Object System.DirectoryServices.DirectorySearcher(\"(&(objectCategory=User)(samAccountName={s.Trim()}))\")).FindOne().GetDirectoryEntry().memberOf";
                    GroupNames = RunPSReturnStr(MyScript);
                    ToText += $"{s.Trim()}'s groups:\r\n\r\n";
                    string[] AllGroups = GroupNames.Split('\n');
                    foreach (string t in AllGroups)
                    {
                        try
                        {
                            if (t.Trim() != "")
                            {
                                temptext += t.Trim().Split(',')[0].Split('=')[1] + "\r\n";
                                if (ADGroupListBox.Items.Contains(t.Trim().Split(',')[0].Split('=')[1]))
                                {

                                }
                                else
                                {
                                    ADGroupListBox.Items.Add(t.Trim().Split(',')[0].Split('=')[1]);
                                }
                            }
                        }
                        catch
                        {
                            temptext += $"{t}\r\n";
                        }
                    }
                    ToText += $"{temptext}\r\n";
                    temptext = "";
                    ToText += $"--------------------\r\n\r\n";
                }
                ADOutputBox.Text = ToText;
            }
        }

        private void ClearADUserListBox_Click(object sender, EventArgs e)
        {
            ADUserListBox.Items.Clear();
        }

        private void ClearADGroupListBox_Click(object sender, EventArgs e)
        {
            ADGroupListBox.Items.Clear();
        }

        private void ADImportUsers_Click(object sender, EventArgs e)
        {
            if (FullCollectionList.Count > 0)
            {
                ADInputBox.Text = "";
                for (int i = 0; i < FullCollectionList.Count; i++)
                {
                    ADInputBox.Text += $"{FullCollectionList[i].Owner}\r\n";
                }
            }
        }

        private void ADImportAssets_Click(object sender, EventArgs e)
        {
            ADInputBox.Text = "";
            if (FullCollectionList.Count > 0)
            {
                for (int i = 0; i < FullCollectionList.Count; i++)
                {
                    ADInputBox.Text += $"{FullCollectionList[i].SerialNumber}\r\n";
                }
            }
        }

        private List<string> GetUsernameFromFullname(string fullName)
        {
            List<string> FinalUsername = new List<string>();
            if (fullName.Trim() == "")
            {
                FinalUsername.Add("");
            }
            else
            {
                string QuickToScript = $"$Item = \"{fullName.Trim()}\"";
                QuickToScript += "\r\n$User = Get-ADUser -Filter{ displayName -like $Item } -Properties SamAccountName";
                QuickToScript += "\r\n$User.SamAccountName";
                string QuickUsername = RunPSReturnStr(QuickToScript);
                //$Item = "Timothy Ferrin"
                //$User = Get-ADUser -Filter{ displayName -like $Item} -Properties SamAccountName
                //$User.SamAccountName
                if (QuickUsername.Trim() == "")
                {
                    string FirstName = fullName.Trim().Split(' ').First().Trim();
                    string LastName = fullName.Trim().Split(' ').Last().Trim();
                    string ToScript = $"$Item = \"{LastName}*\"";
                    ToScript += $"\r\n$Item2 = \"{FirstName}*\"";
                    ToScript += "\r\n$User = Get-ADUser -Filter{sn -like $Item -and givenname -like $Item2}";
                    ToScript += "\r\n$user.name, $user.SamAccountName";
                    string results = RunPSReturnStr(ToScript);
                    string MyName = results.Split('\n')[0].Trim();
                    string UserName = results.Split('\n')[1].Trim();
                    if (UserName.Contains(" "))
                    {
                        //UsernameList += $"Too many possibilities for {GenericAssetList[i].Trim()}.\r\n";
                        string[] tempFullNames1 = MyName.Split(' ');
                        List<string> tempFullNames = new List<string>();
                        for (int j = 0; j < tempFullNames1.Count(); j++)
                        {
                            if (j % 2 == 0)
                            {
                                tempFullNames.Add($"{tempFullNames1[j]} {tempFullNames1[j + 1]}");
                            }
                        }
                        string[] tempUsernames = UserName.Split(' ');
                        for (int j = 0; j < tempUsernames.Count(); j++)
                        {
                            if (MessageBox.Show($"Verify: {fullName.Trim()}: {tempFullNames[j]}?  ({tempUsernames[j]})", "Name Selection", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                FinalUsername.Add($"{tempUsernames[j]}");
                            }
                        }
                    }
                    else if (MessageBox.Show($"Verify: {fullName.Trim()}: {MyName}?  ({UserName})", "Name Selection", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        FinalUsername.Add($"{UserName}");
                    }
                    else
                    {
                        if (MessageBox.Show($"{fullName.Trim()} is not a full name.  Discard?", $"Discard {fullName.Trim()}", MessageBoxButtons.YesNo) == DialogResult.No)
                        {
                            FinalUsername.Add($"{fullName}");
                        }
                        else
                        {
                            FinalUsername.Add("");
                        }
                    }
                }
                else
                {
                    FinalUsername.Add(QuickUsername);
                }
            }
            return FinalUsername;
        }

        private void ADUsernamesFromFull_Click(object sender, EventArgs e)
        {
            string Usernames = "";
            string[] FullNames = ADInputBox.Text.Split('\n');
            for (int i = 0; i < FullNames.Count(); i++)
            {
                List<string> tempUNs = GetUsernameFromFullname(FullNames[i]);
                foreach (string s in tempUNs)
                {
                    Usernames += s.Trim() + "\r\n";
                }
            }
            ADOutputBox.Text = Usernames;
        }
        private string GetDSInfo(string username, string method)
        {
            string MyScript = $"(New-Object System.DirectoryServices.DirectorySearcher(\"(&(objectCategory=User)(samAccountName={username.Trim()}))\")).FindOne().GetDirectoryEntry().{method}";
            string Outcome = RunPSReturnStr(MyScript);
            return Outcome;
        }

        private void ADGetReports_Click(object sender, EventArgs e)
        {
            //gets users who report to listed user

            ADOutputBox.Text = "----------\r\n";
            string[] userList = ADInputBox.Text.Split('\n');
            foreach (string user in userList)
            {
                string tempUser = user.Trim();
                ADOutputBox.Text += $"{user.Trim()}:\r\n\r\n";
                String ManagerInfo = GetDSInfo(tempUser, "directreports").Trim();
                try
                {
                    string[] reports = ManagerInfo.Split('\n');
                    foreach (string report in reports)
                    {
                        string fullname = (report.Split('=')[1]).Split(',')[0].Trim();
                        ADOutputBox.Text += $"{fullname.Trim()}\r\n";
                    }
                }
                catch
                {

                }
                ADOutputBox.Text += "\r\n----------\r\n";
            }
        }

        private void ADQuickInfo_Click(object sender, EventArgs e)
        {
            string[] usernames = ADInputBox.Text.Split('\r');
            //string QuickExport = "Username|Department|Email Address|Employee ID|Full Name|Locked?|Last Failed Login|Last Login|Manager|" +
            //    "Office Location|Password Expires|Password Last Changed|Title|OU Path\r\n";
            string QuickExport = "";
            string temp = "";
            foreach (string username in usernames)
            {
                QuickExport += $"Username: {username.Trim()}\r\n";
                temp = GetDSInfo(username.Trim(), "Department").Trim();
                QuickExport += $"Department: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "EmailAddress").Trim();
                QuickExport += $"Email Address: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "EmployeeID").Trim();
                QuickExport += $"Employee ID: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "FullName").Trim();
                QuickExport += $"Full Name: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "IsAccountLocked").Trim();
                QuickExport += $"Account Locked? {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "LastFailedLogin").Trim();
                QuickExport += $"Last Failed Login: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "LastLogin").Trim();
                QuickExport += $"Last Login: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "Manager").Trim();
                try { temp = temp.Split(',')[0].Split('=')[1]; } catch { }
                QuickExport += $"Manager: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "OfficeLocations").Trim();
                QuickExport += $"Office Location: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "PasswordExpirationDate").Trim();
                QuickExport += $"Password Expiration Date: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "PasswordLastChanged").Trim();
                QuickExport += $"Password Last Changed: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "Title").Trim();
                QuickExport += $"Title: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "AdsPath").Trim();
                QuickExport += $"OU Path: {temp}\r\n\r\n";
            }
            ADOutputBox.Text = QuickExport;
        }

        private void ADFolderPermissions_Click(object sender, EventArgs e)
        {
            ADOutputBox.Text = "";
            if (ADInputBox.Text.Trim() == "")
            {
                FolderBrowserDialog fbd1 = new FolderBrowserDialog();
                fbd1.SelectedPath = "\\\\corp\\securecorp\\";
                DialogResult fbd1Result = fbd1.ShowDialog();
                ADInputBox.Text = fbd1.SelectedPath;
            }
            string[] folderList = ADInputBox.Text.Split('\n');
            string folderListBuilder = "";
            foreach (string folder in folderList)
            {
                string tempFolder = folder.Trim();
                String FolderInfo = RunPSReturnStr("(get-acl \"" + tempFolder + "\").access | select(\"identityreference\", \"filesystemrights\")").Trim();

                folderListBuilder += tempFolder + "\r\n";
                folderListBuilder += FolderInfo;
                folderListBuilder += "\r\n\r\n";
            }
            ADOutputBox.Text = folderListBuilder.Trim();
        }

        private void ADCompareTwoUsers_Click(object sender, EventArgs e)
        {
            //compares groups of 2 different users
            ADOutputBox.Text = "";
            string[] ToSearch = ADInputBox.Text.Split('\n');
            PowerShell ps = PowerShell.Create();
            ps.AddScript("diff ((get-aduser " + ToSearch[0].Trim() + " -properties memberof).memberof) ((get-aduser " + ToSearch[1].Trim() + " -properties memberof).memberof) -includeequal");
            try
            {
                Collection<PSObject> PSOutput = ps.Invoke();
                string AllInfo = "";
                string TempInfo = "";
                foreach (PSObject outputItem in PSOutput)
                {
                    TempInfo = Convert.ToString(outputItem);
                    if (TempInfo.Substring(TempInfo.Length - 3) == "<=}")
                    {
                        AllInfo += ToSearch[0] + " -- ";
                    }
                    else if (TempInfo.Substring(TempInfo.Length - 3) == "=>}")
                    {
                        AllInfo += ToSearch[1] + " -- ";
                    }
                    else
                    {
                        AllInfo += "======= ";
                    }
                    TempInfo = TempInfo.Split('=')[2];
                    TempInfo = TempInfo.Substring(0, TempInfo.Length - 3);
                    AllInfo += TempInfo;
                    AllInfo += "\r\n";
                }

                ADOutputBox.Text += AllInfo;
            }
            catch (Exception ex)
            {
                ADOutputBox.Text += ex;
                ADOutputBox.Text += "Error finding groups.  Try again.";
                ADOutputBox.Text += "\r\n";
            }
        }

        private void ADGetGroupMembership_Click(object sender, EventArgs e)
        {
            ADOutputBox.Text = "";
            string[] groupsToCheck = ADInputBox.Text.Split('\n');
            string output = "----------\r\n";
            foreach (string s in groupsToCheck)
            {
                output += $"{s.Trim()}:\r\n\r\n";
                string forScript = $"(New-Object System.DirectoryServices.DirectorySearcher(\"(&(objectCategory=Group)(name={s.Trim()}))\")).FindAll().GetDirectoryEntry().Member";
                string names = RunPSReturnStr(forScript);
                foreach (string name in names.Split('\n'))
                {
                    if (name.Trim() != "")
                    {
                        try { output += $"{name.Split(',')[0].Split('=')[1].Trim()}\r\n"; } catch { output += $"Problem with {s.Trim()}\r\n"; }
                    }
                }
                output += "\r\n----------\r\n";
            }
            ADOutputBox.Text = output;
        }

        private void ADBulkGroupMembership_Click(object sender, EventArgs e)
        {
            ADOutputBox.Text = "";
            string[] groupsToCheck = ADInputBox.Text.Split('\n');
            string output = "";
            foreach (string s in groupsToCheck)
            {
                output += $"{s.Trim()}|";
                string forScript = $"(New-Object System.DirectoryServices.DirectorySearcher(\"(&(objectCategory=Group)(name={s.Trim()}))\")).FindAll().GetDirectoryEntry().Member";
                string names = RunPSReturnStr(forScript);
                foreach (string name in names.Split('\n'))
                {
                    if (name.Trim() != "")
                    {
                        try { output += $"{name.Split(',')[0].Split('=')[1].Trim()}|"; } catch { output += $"Problem with {s.Trim()}|"; }
                    }
                }
                output += "\r\n";
            }
            ADOutputBox.Text = output;
        }

        private void ADCheckForEnabled_Click(object sender, EventArgs e)
        {
            //checks if listed users are enabled in AD
            ADOutputBox.Text = "";
            int NumTrue = 0;
            int NumFalse = 0;
            string[] ToSearch = ADInputBox.Text.Split('\n');
            foreach (string Username in ToSearch)
            {
                string TempString = GetDSInfo(Username, "useraccountcontrol");
                string TempString2 = "";
                if (TempString.Trim() == "512")
                {
                    TempString2 += Username.Trim() + " -> " + " Enabled\r\n";
                    NumTrue++;
                }
                else if (TempString.Trim() == "514")
                {
                    TempString2 += Username.Trim() + " -> " + " Disabled\r\n";
                    NumFalse++;
                }
                else
                {
                    TempString2 += Username.Trim() + " -> " + " Something went wrong, skipping this name\r\n";
                }

                ADOutputBox.Text += TempString2;
            }
            ADOutputBox.Text += "\r\nNumber of enabled accounts: " + NumTrue + "\r\nNumber of disabled accounts: " + NumFalse + "\r\n";
        }

        private void Unlock_Click(object sender, EventArgs e)
        {
            //unlocks listed user in AD
            string UnlockID = ModifyUsername.Text;
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Unlock-adaccount -identity " + UnlockID);
            Collection<PSObject> PSOutput = ps.Invoke();
        }

        private void VerifyUsernames_Click(object sender, EventArgs e)
        {
            ModOutputBox.Text = "";
            string returnString = "";
            string[] usernames = ModUsernameBox.Text.Split('\n');
            foreach (string username in usernames)
            {
                string toCheck = RunPSReturnStr($"Get-ADUser {username.Trim()}");
                if (toCheck.Contains("Script Failed") || toCheck == "")
                {
                    returnString += $"{username.Trim()} - Invalid\r\n";
                }
                else if (toCheck.Contains("OU=Disabled Users"))
                {
                    returnString += $"{username.Trim()} - Disabled\r\n";
                }
                else
                {
                    returnString += username.Trim() + "\r\n";
                    ModOutputBox.Text += $"Confirmed {username.Trim()}.\r\n";
                }
            }
            ModUsernameBox.Text = returnString.Trim();
        }

        private void VerifyGroups_Click(object sender, EventArgs e)
        {
            ModOutputBox.Text = "";
            string returnString = "";
            string[] usernames = ModGroupBox.Text.Split('\n');
            foreach (string username in usernames)
            {
                string toCheck = RunPSReturnStr($"Get-ADGroup \"{username.Trim()}\"");
                if (toCheck.Contains("Script Failed") || toCheck == "")
                {
                    returnString += $"{username.Trim()} - Invalid\r\n";
                }
                else if (toCheck.Contains("OU=Disabled Users"))
                {
                    returnString += $"{username.Trim()} - Disabled\r\n";
                }
                else
                {
                    returnString += username.Trim() + "\r\n";
                    ModOutputBox.Text += $"Confirmed {username.Trim()}.\r\n";
                }
            }
            ModGroupBox.Text = returnString.Trim();
        }

        private void AddUsersToGroups_Click(object sender, EventArgs e)
        {
            //get-adgroup "zoom accounts" | Add-ADGroupMember -Members tferrin
            ModOutputBox.Text = "";
            string ModOutput = "WARNING: If a username or group name is invalid and not confirmed, the script will still claim that it has been added.\r\n\r\n";
            if (MessageBox.Show("Add these users to the requested groups?  If you haven't verified users and groups, select 'no'.\r\n\r\nWARNING! This is a powerful tool!", "Confirm Group Add", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string[] usernames = ModUsernameBox.Text.Split('\n');
                string[] groupnames = ModGroupBox.Text.Split('\n');
                foreach (string username in usernames)
                {
                    if (username.Trim().Contains(" - Invalid") || username.Trim().Contains(" - Disabled"))
                    {
                        ModOutput += $"Omitting \"{username.Trim()}\", user is Invalid or Disabled.\r\n";
                    }
                    else
                    {
                        foreach (string groupname in groupnames)
                        {
                            if (groupname.Trim().Contains(" - Invalid") || groupname.Trim().Contains(" - Disabled"))
                            {
                                ModOutput += $"     Omitting \"{groupname.Trim()}\", group is Invalid or Disabled.\r\n";
                            }
                            else
                            {
                                ModOutput += $"     Adding {username.Trim()} to {groupname.Trim()}.\r\n";
                                string psOut = RunPSReturnStr($"Get-ADGroup \"{groupname.Trim()}\" | Add-ADGroupMember -Members \"{username.Trim()}\"");
                                if (psOut == "Script Failed")
                                {
                                    ModOutput += $"          Failed to add {groupname.Trim()} to {username.Trim()}.\r\n";
                                }
                                else
                                {
                                    ModOutput += $"          Added {username.Trim()} to {groupname.Trim()}.\r\n";
                                }
                            }
                        }
                    }
                }
            }
            ModOutputBox.Text = ModOutput;
        }

        private void RemoveUsersFromGroups_Click(object sender, EventArgs e)
        {
            //Get-ADGroup "zoom accounts" | Remove-ADGroupMember -Members tferrin -Confirm:$false
            ModOutputBox.Text = "";
            string ModOutput = "WARNING: If a username or group name is invalid and not confirmed, the script will still claim that it has been added.\r\n\r\n";
            if (MessageBox.Show("Remove these users from the requested groups?  If you haven't verified users and groups, select 'no'.\r\n\r\nWARNING! This is a powerful tool!", "Confirm Group Add", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string[] usernames = ModUsernameBox.Text.Split('\n');
                string[] groupnames = ModGroupBox.Text.Split('\n');
                foreach (string username in usernames)
                {
                    if (username.Trim().Contains(" - Invalid") || username.Trim().Contains(" - Disabled"))
                    {
                        ModOutput += $"Omitting \"{username.Trim()}\", user is Invalid or Disabled.\r\n";
                    }
                    else
                    {
                        foreach (string groupname in groupnames)
                        {
                            if (groupname.Trim().Contains(" - Invalid") || groupname.Trim().Contains(" - Disabled"))
                            {
                                ModOutput += $"     Omitting \"{groupname.Trim()}\", group is Invalid or Disabled.\r\n";
                            }
                            else
                            {
                                ModOutput += $"     Removing {username.Trim()} from {groupname.Trim()}.\r\n";
                                string psOut = RunPSReturnStr($"Get-ADGroup \"{groupname.Trim()}\" | Remove-ADGroupMember -Members \"{username.Trim()}\" -Confirm:$false");
                                if (psOut == "Script Failed")
                                {
                                    ModOutput += $"          Failed to remove {groupname.Trim()} from {username.Trim()}.\r\n";
                                }
                                else
                                {
                                    ModOutput += $"          Removed {username.Trim()} from {groupname.Trim()}.\r\n";
                                }
                            }
                        }
                    }
                }
            }
            ModOutputBox.Text = ModOutput;
        }

        private void RemoveNewLines_Click(object sender, EventArgs e)
        {
            //Removes all newlines from a string
            string EditMe = TextEditBox.Text;
            EditMe = EditMe.Replace(System.Environment.NewLine, "");
            TextEditBox.Text = EditMe;
        }

        private void EnabledExport_Click(object sender, EventArgs e)
        {
            // Takes the output from Search For Enabled and creates an exportable string
            string[] outputString = TextEditBox.Text.Split('\n');
            string newString = "";
            string newSubString = "";
            List<string> listEnabled = new List<string>();
            List<string> listDisabled = new List<string>();
            List<string> listError = new List<string>();
            foreach (string s in outputString)
            {
                if (s.Contains("Number of "))
                {
                    newSubString += "";
                }
                else
                {
                    newSubString += s.Trim() + "\r\n";
                }
            }


            outputString = newSubString.Split('\n');
            foreach (string s in outputString)
            {
                if (s.Trim() == "")
                {

                }
                else
                {
                    newString += s.Trim() + "\r\n";
                }
            }

            outputString = newString.Split('\n');
            foreach (string s in outputString)
            {
                if (s.Contains(" ->  Enabled"))
                {
                    listEnabled.Add(s.Split(' ')[0].Trim());
                }
                if (s.Contains(" ->  Disabled"))
                {
                    listDisabled.Add(s.Split(' ')[0].Trim());
                }
                if (s.Contains(" ->  Something went wrong, skipping this name"))
                {
                    listError.Add(s.Split(' ')[0].Trim());
                }
            }

            List<string> FinalList = new List<string>();

            FinalList.Add("Enabled|Disabled|Errors");
            List<int> t1 = new List<int>();
            t1.Add(listEnabled.Count());
            t1.Add(listDisabled.Count());
            t1.Add(listError.Count());
            int totalRows = t1.Max();
            for (int i = 0; i < totalRows; i++)
            {
                listEnabled.Add("");
                listDisabled.Add("");
                listError.Add("");
            }

            for (int i = 0; i < totalRows; i++)
            {
                FinalList.Add(listEnabled[i] + "|" + listDisabled[i] + "|" + listError[i]);
            }

            string MyFinal = "";
            foreach (string s in FinalList)
            {
                MyFinal += s.Trim();
                MyFinal += "\r\n";
            }
            TextEditBox.Text = MyFinal.Trim();
        }

        private void JSONCleaner_Click(object sender, EventArgs e)
        {
            try
            {
                string textToEdit = TextEditBox.Text;
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject(textToEdit);
                var editedText = Newtonsoft.Json.JsonConvert.SerializeObject(obj, Newtonsoft.Json.Formatting.Indented);
                //var editedText = JsonConvert.SerializeObject(textToEdit, Newtonsoft.Json.Formatting.Indented);
                TextEditBox.Text = editedText;
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("JSON Format Failed!");
            }
        }

        private void SearchAndReplace1_Click(object sender, EventArgs e)
        {
            //replaces text from texteditbox based on findbox and replacebox
            string EditMe = TextEditBox.Text;
            EditMe = EditMe.Replace(FindBox.Text, ReplaceBox.Text);
            TextEditBox.Text = EditMe;
        }

        private void SearchAndReplace2_Click(object sender, EventArgs e)
        {
            //replaces selected text from texteditbox based on findbox and replacebox
            try
            {
                string EditMe = "";
                EditMe = TextEditBox.SelectedText;
                if (EditMe != "")
                {
                    EditMe = EditMe.Replace(FindBox.Text, ReplaceBox.Text);
                }
                int Start = TextEditBox.SelectionStart;
                int a = TextEditBox.SelectionLength;
                TextEditBox.Text = TextEditBox.Text.Remove(TextEditBox.SelectionStart, a);
                TextEditBox.Text = TextEditBox.Text.Insert(Start, EditMe);
            }
            catch { }
        }

        private void TicketGetMyTickets_Click(object sender, EventArgs e)
        {
            //gets current user's open tickets across all projects
            //if (MessageBox.Show("Really delete?", "Confirm delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //{
            TicketInputBox.Text = "https://jira.tlcinternal.com/rest/api/2/search?jql=assignee+%3D+currentUser()+AND+resolution+%3D+Unresolved";
            RestClient rClient = new RestClient();
            rClient.endPoint = TicketInputBox.Text;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            TicketOutputBox.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONFilteredTickets(strResponse);
            //}
        }
        private void deserialiseJSONFilteredTickets(string strJSON)
        {
            TicketListOutputBox.Items.Clear();
            try
            {
                var jPerson = JsonConvert.DeserializeObject<dynamic>(strJSON);
                TicketOutputBox.Text = "";

                int TotalCount = Convert.ToInt32(jPerson.total);
                TicketOutputBox.Text = $"Total issue count: {TotalCount}\r\n";
                for (int i = 0; i < TotalCount; i++)
                {
                    TicketListOutputBox.Items.Add($"{jPerson.issues[i].key} -- {jPerson.issues[i].fields.summary}");
                }

                foreach (string s in TicketListOutputBox.Items)
                {
                    TicketOutputBox.Text += s + "\r\n";
                }
            }
            catch (Exception ex)
            {
                TicketOutputBox.Text = "We had a problem: " + ex.Message.ToString();
            }
        }

        private void TicketSearchID_Click(object sender, EventArgs e)
        {
            //searches a single ticket by id
            string myUrl = "https://jira.tlcinternal.com/rest/api/2/issue/" + TicketInputBox.Text;
            RestClient rClient = new RestClient();
            rClient.endPoint = myUrl;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            TicketOutputBox.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONTicket(strResponse);
            SmallTicketBox.Text = TicketInputBox.Text;
        }
        private void deserialiseJSONTicket(string strJSON)
        {
            try
            {
                string tNum;
                string tTitle;
                string tDesc;
                string tAssignee;
                string tReporter;
                string tFullName;
                var jPerson = JsonConvert.DeserializeObject<dynamic>(strJSON);
                try { tNum = jPerson.key; } catch { tNum = ""; }
                try { tTitle = jPerson.fields.summary; } catch { tTitle = ""; }
                try { tDesc = jPerson.fields.description; } catch { tDesc = ""; }
                try { tAssignee = jPerson.fields.assignee.displayName; } catch { tAssignee = ""; }
                try { tReporter = jPerson.fields.reporter.displayName; } catch { tReporter = ""; }
                try { tFullName = jPerson.fields.reporter.displayName; } catch { tFullName = ""; }
                TicketOutputBox.Text = $"Ticket Number: {tNum}\r\nTicket Title: {tTitle}\r\nTicket Description: {tDesc}\r\nTicket Assignee: {tAssignee}\r\nTicket Reporter: {tReporter}\r\n";
                TicketOutputBox.Text += "Desk Location: ";
                RetrieveDeskLocation(tFullName);
                string myComments = RetrieveJSONComments("https://jira.tlcinternal.com/rest/api/2/issue/" + jPerson.key + "/comment/");
                deserialiseJSONComments(myComments);
            }
            catch (Exception ex)
            {
                TicketOutputBox.Text = "We had a problem at deserialiseJSONTicket: " + ex.Message.ToString();
            }
        }
        private string RetrieveDeskLocation(string FullName)
        {
            string loc = "";

            string objectlocstring = " " + '\u0022' + "label" + '\u0022' + " = ";

            RestClient rClient = new RestClient();
            rClient.endPoint = $"https://jira.tlcinternal.com/rest/insight/1.0/iql/objects?objectSchemaId=1&iql=" + objectlocstring + '"' + FullName + '"';
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;
            SmallTicketBox.Text = rClient.endPoint;

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONAssetDeskLocation(strResponse);

            return loc;
        }
        private void deserialiseJSONAssetDeskLocation(string strJSON)
        {
            try
            {
                var jPerson = JsonConvert.DeserializeObject<dynamic>(strJSON);
                //TextOutputBox.Text = "";


                try
                {
                    TicketOutputBox.Text += jPerson.objectEntries[0].attributes[9].objectAttributeValues[0].displayValue + "\r\n";
                }
                catch
                {
                    TicketOutputBox.Text += "Created: Not Listed.\r\n";
                }

            }
            catch (Exception ex)
            {
                TicketOutputBox.Text += "We had a problem: " + ex.Message.ToString();
            }
        }
        private string RetrieveJSONComments(string URL)
        {
            string myJSON = "";

            RestClient rClient = new RestClient();
            rClient.endPoint = URL;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            //textBox4.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            myJSON = strResponse;

            return myJSON;
        }
        private void deserialiseJSONComments(string strJSON)
        {
            try
            {
                var jPerson = JsonConvert.DeserializeObject<dynamic>(strJSON);

                for (int i = 0; i < Convert.ToInt32(jPerson.total); i++)
                {
                    TicketOutputBox.Text += $"\r\n-----\r\n{jPerson.comments[i].created}\r\nComment {i + 1}: {jPerson.comments[i].body}";
                }
            }
            catch (Exception ex)
            {
                TicketOutputBox.Text += "We had a problem: " + ex.Message.ToString();
            }
        }

        private void TicketSearchMultiID_Click(object sender, EventArgs e)
        {
            TicketListOutputBox.Items.Clear();
            string[] Lines1 = TicketInputBox.Text.Split('\r');
            foreach (string line in Lines1)
            {
                TicketListOutputBox.Items.Add(line.Trim());
            }
        }

        private void TicketSearchAssignee_Click(object sender, EventArgs e)
        {
            //checks all open tickets by Assignee
            string s = TicketInputBox.Text;
            TicketInputBox.Text = $"assignee = {s} AND resolution = unresolved";
            RestClient rClient = new RestClient();
            rClient.endPoint = "https://jira.tlcinternal.com/rest/api/2/search?jql=" + TicketInputBox.Text;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            TicketOutputBox.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONFilteredTickets(strResponse);
        }

        private void TicketSearchReporter_Click(object sender, EventArgs e)
        {
            //checks all open tickets by Reporter
            string s = TicketInputBox.Text;
            TicketInputBox.Text = $"reporter = {s} AND resolution = unresolved";
            RestClient rClient = new RestClient();
            rClient.endPoint = "https://jira.tlcinternal.com/rest/api/2/search?jql=" + TicketInputBox.Text;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            TicketOutputBox.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONFilteredTickets(strResponse);
        }

        private void TicketSearchJQL_Click(object sender, EventArgs e)
        {
            //performs a ticket search by JQL

            RestClient rClient = new RestClient();
            rClient.endPoint = "https://jira.tlcinternal.com/rest/api/2/search?jql=" + TicketInputBox.Text;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            TicketOutputBox.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONFilteredTickets(strResponse);
        }

        private void TicketNewHirePull_Click(object sender, EventArgs e)
        {
            //finds New Hire Forms by date
            TicketInputBox.Text = $"https://jira.tlcinternal.com/rest/api/2/search?jql=Issuetype%20=%20%22New%20Hire%22%20AND%20summary%20~%20" + TicketInputBox.Text;
            RestClient rClient = new RestClient();
            rClient.endPoint = TicketInputBox.Text;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            TicketOutputBox.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONFilteredTickets(strResponse);
        }

        private void TicketCreateHDTicket_Click(object sender, EventArgs e)
        {
            //Creates tickets in the HD queue under the "IT Help" issuetype
            if (MessageBox.Show("Create these ticket(s)?", "Confirm Ticket Creation", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                TicketListOutputBox.Items.Clear();
                TicketOutputBox.Text = "";

                string[] Lines = TicketInputBox.Text.Split('\r');

                foreach (string Line in Lines)
                {
                    string addyPOSTIssue = "https://jira.tlcinternal.com/rest/api/2/issue";
                    //11600 is the HD project ID, 10300 is the "IT Help" issue type
                    string TicketForm = "{\"fields\":{\"project\":{\"id\":\"11600\"},\"summary\":\"" + Line.Trim() + "\",\"issuetype\":{\"id\":\"10300\"},\"assignee\":{\"name\":\"" + JiraUsernameBox.Text + "\"}}}";

                    RestClient rClient = new RestClient("POST");
                    rClient.endPoint = addyPOSTIssue;
                    rClient.authTech = autheticationTechnique.RollYourOwn;
                    rClient.authType = authenticationType.Basic;
                    rClient.userName = JiraUsernameBox.Text;
                    rClient.userPassword = JiraPasswordBox.Text;
                    rClient.postJSON = TicketForm;

                    string strResponse = string.Empty;

                    strResponse = rClient.makeRequest();
                    TicketOutputBox.Text += strResponse;

                    var TicketStuff = JsonConvert.DeserializeObject<dynamic>(strResponse);
                    string TicketKey = "";
                    try { TicketKey = TicketStuff.key; } catch { }
                    TicketListOutputBox.Items.Add(TicketKey);
                }
            }
        }

        private void OpenTicketInBrowser_Click(object sender, EventArgs e)
        {
            //Opens currently selected ticket in browser
            System.Diagnostics.Process.Start($"https://jira.tlcinternal.com/browse/{SmallTicketBox.Text}");
        }

        private void OpenAllTicketsFromList_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure?  This will open tab(s) for ALL tickets in the list.", "Confirm Ticket Opening", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (string s in TicketListOutputBox.Items)
                {
                    string search = $"{s.Split('-')[0].Trim()}-{s.Split('-')[1].Trim()}";
                    System.Diagnostics.Process.Start($"https://jira.tlcinternal.com/browse/{search}");
                }
            }
        }

        private void ExportTicketListOutput_Click(object sender, EventArgs e)
        {
            //creates an exportable list in outputbox
            TicketOutputBox.Text = "Key|Summary|Assignee|Reporter|Created|Last Attachment Date|End Date|Description|\r\n";
            foreach (string s in TicketListOutputBox.Items)
            {
                SmallTicketBox.Text = s.Split(' ')[0].Trim();

                string myUrl = "https://jira.tlcinternal.com/rest/api/2/issue/" + SmallTicketBox.Text;
                RestClient rClient = new RestClient();
                rClient.endPoint = myUrl;
                rClient.authTech = autheticationTechnique.RollYourOwn;
                rClient.authType = authenticationType.Basic;
                rClient.userName = JiraUsernameBox.Text;
                rClient.userPassword = JiraPasswordBox.Text;

                string strResponse = string.Empty;

                strResponse = rClient.makeRequest();

                TicketExportBuilder(strResponse);
            }
        }
        private void TicketExportBuilder(string TicketInfo)
        {
            try
            {
                string tNum;
                string tTitle;
                string tCreated;
                string tAssignee;
                string tReporter;
                string tAttachmentDate;
                string tEndDate;
                int tEndDateIndex = 0;
                string tDescription;
                var jPerson = JsonConvert.DeserializeObject<dynamic>(TicketInfo);
                try { tNum = jPerson.key; } catch { tNum = "No Info"; }
                try { tTitle = jPerson.fields.summary; } catch { tTitle = "No Info"; }
                try { tCreated = jPerson.fields.created; } catch { tCreated = "No Info"; }
                try { tAssignee = jPerson.fields.assignee.displayName; } catch { tAssignee = "No Info"; }
                try { tReporter = jPerson.fields.reporter.displayName; } catch { tReporter = "No Info"; }
                try { tAttachmentDate = jPerson.fields.attachment.Last.created; } catch { tAttachmentDate = "No Info"; }
                try { tEndDate = jPerson.fields.customfield_10305.First; } catch { tEndDate = "No Info"; }
                try { tEndDateIndex = tEndDate.IndexOf("endDate="); } catch { }
                try { tEndDate = tEndDate.Substring(tEndDateIndex + 8).Split('T')[0]; } catch { }
                try { tDescription = jPerson.fields.description; } catch { tDescription = "No Info"; }
                try { tDescription = tDescription.Replace("\r", " "); } catch { }
                try { tDescription = tDescription.Replace("\n", " "); } catch { }
                try { tDescription = tDescription.Replace("|", " - "); } catch { }
                TicketOutputBox.Text += $"{tNum}|{tTitle}|{tAssignee}|{tReporter}|{tCreated}|{tAttachmentDate}|{tEndDate}|{tDescription}\r\n";
                //string myComments = RetrieveJSONComments("https://jira.tlcinternal.com/rest/api/2/issue/" + jPerson.key + "/comment/");
                //deserialiseJSONComments(myComments);
            }
            catch (Exception ex)
            {
                TicketOutputBox.Text = "We had a problem at DeserialiseObject: " + ex.Message.ToString();
            }
        }

        private void ClearTicketFields_Click(object sender, EventArgs e)
        {
            TicketInputBox.Text = "";
            TicketOutputBox.Text = "";
            TicketListOutputBox.Items.Clear();
            SmallTicketBox.Text = "";
        }

        private void TicketListOutputBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Gives info for the selected ticket

            SmallTicketBox.Text = Convert.ToString(TicketListOutputBox.SelectedItem).Split(' ')[0];
            string myUrl = "https://jira.tlcinternal.com/rest/api/2/issue/" + SmallTicketBox.Text;
            RestClient rClient = new RestClient();
            rClient.endPoint = myUrl;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JiraUsernameBox.Text;
            rClient.userPassword = JiraPasswordBox.Text;

            TicketOutputBox.Text = ("Rest Client Created");

            string strResponse = string.Empty;

            strResponse = rClient.makeRequest();

            deserialiseJSONTicket(strResponse);
            SmallTicketBox.Text = Convert.ToString(TicketListOutputBox.SelectedItem).Split(' ')[0].Trim();
        }

        private void PullHDSlack_Click(object sender, EventArgs e)
        {
            //successfully pushed message to slack channel via slack api integration using this:
            //--------------------
            //var slackClient = new SlackClient("https://hooks.slack.com/services/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");
            //var slackMessage = new SlackMessage
            //{
            //    Text = "Testing this from a C# program."
            //};
            //slackClient.Post(slackMessage);
            //--------------------
            //https://slack.com/api/conversations.history?token=xoxb-101335543696-1378130353059-voL4FWpjoVt9DOkt968dbfoN&channel=C01B12DGMUM
            string tempText = "";
            string JSSURL = "https://slack.com/api/conversations.history?token=xoxb-101335543696-1378130353059-voL4FWpjoVt9DOkt968dbfoN&channel=CMESA6V2R&limit=10";
            RestClient rClient = new RestClient();
            rClient.endPoint = JSSURL;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JSSUsernameBox.Text;
            rClient.userPassword = JSSPasswordBox.Text;
            string strResponse = string.Empty;
            strResponse = rClient.makeRequest();
            SlackData = strResponse;
            var jPerson = JsonConvert.DeserializeObject<dynamic>(strResponse);
            try { tempText = Convert.ToString(jPerson.messages[0].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg1.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[0].text); } catch { tempText = ""; }
            SlackMsg1.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[1].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg2.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[1].text); } catch { tempText = ""; }
            SlackMsg2.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[2].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg3.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[2].text); } catch { tempText = ""; }
            SlackMsg3.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[3].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg4.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[3].text); } catch { tempText = ""; }
            SlackMsg4.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[4].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg5.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[4].text); } catch { tempText = ""; }
            SlackMsg5.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[5].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg6.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[5].text); } catch { tempText = ""; }
            SlackMsg6.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[6].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg7.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[6].text); } catch { tempText = ""; }
            SlackMsg7.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[7].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg8.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[7].text); } catch { tempText = ""; }
            SlackMsg8.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[8].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg9.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[8].text); } catch { tempText = ""; }
            SlackMsg9.Text += $"{tempText}";
            try { tempText = Convert.ToString(jPerson.messages[9].user); } catch { tempText = ""; }
            tempText = CollectUsernameFromSlackID(tempText);
            SlackMsg10.Text = $"{tempText}:\r\n";
            try { tempText = Convert.ToString(jPerson.messages[9].text); } catch { tempText = ""; }
            SlackMsg10.Text += $"{tempText}";
        }
        private string CollectUsernameFromSlackID(string SlackID)
        {
            //To get the user from user id:
            //https://slack.com/api/users.profile.get?token=xoxb-101335543696-1378130353059-XCnpHMXK2mG4kgVJuWENfJJa&user=WHWJ3B9ED

            string ToReturn = "";
            string JSSURL = "https://slack.com/api/users.profile.get?token=xoxb-101335543696-1378130353059-voL4FWpjoVt9DOkt968dbfoN&user=" + SlackID;
            RestClient rClient = new RestClient();
            rClient.endPoint = JSSURL;
            rClient.authTech = autheticationTechnique.RollYourOwn;
            rClient.authType = authenticationType.Basic;
            rClient.userName = JSSUsernameBox.Text;
            rClient.userPassword = JSSPasswordBox.Text;
            string strResponse = string.Empty;
            strResponse = rClient.makeRequest();
            var jPerson = JsonConvert.DeserializeObject<dynamic>(strResponse);
            try { ToReturn = Convert.ToString(jPerson.profile.real_name); } catch { }
            return ToReturn;
        }

        private void FixSlack1_Click(object sender, EventArgs e)
        {
            PushSlackOutput(1);
        }
        private void PushSlackOutput(int ButtonNumber)
        {
            string MyUser = "";
            try
            {
                var jPerson = JsonConvert.DeserializeObject<dynamic>(SlackData);
                string user = Convert.ToString(jPerson.messages[ButtonNumber - 1].user);
                MyUser = CollectUsernameFromSlackID(user);
            }
            catch { }
            List<string> tempUN = GetUsernameFromFullname(MyUser);
            SlackRemediationBox.Text = $"{MyUser}\r\n{tempUN[0]}\r\n";
            string QuickExport = "";
            string temp = "";
            foreach (string username in tempUN)
            {
                QuickExport += $"Username: {username.Trim()}\r\n";
                temp = GetDSInfo(username.Trim(), "Department").Trim();
                QuickExport += $"Department: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "EmailAddress").Trim();
                QuickExport += $"Email Address: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "EmployeeID").Trim();
                QuickExport += $"Employee ID: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "FullName").Trim();
                QuickExport += $"Full Name: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "IsAccountLocked").Trim();
                QuickExport += $"Account Locked? {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "LastFailedLogin").Trim();
                QuickExport += $"Last Failed Login: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "LastLogin").Trim();
                QuickExport += $"Last Login: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "Manager").Trim();
                try { temp = temp.Split(',')[0].Split('=')[1]; } catch { }
                QuickExport += $"Manager: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "OfficeLocations").Trim();
                QuickExport += $"Office Location: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "PasswordExpirationDate").Trim();
                QuickExport += $"Password Expiration Date: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "PasswordLastChanged").Trim();
                QuickExport += $"Password Last Changed: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "Title").Trim();
                QuickExport += $"Title: {temp}\r\n";
                temp = GetDSInfo(username.Trim(), "AdsPath").Trim();
                QuickExport += $"OU Path: {temp}\r\n\r\n";
            }
            SlackRemediationBox.Text += QuickExport;
        }

        private void FixSlack2_Click(object sender, EventArgs e)
        {
            PushSlackOutput(2);
        }

        private void FixSlack3_Click(object sender, EventArgs e)
        {
            PushSlackOutput(3);
        }

        private void FixSlack4_Click(object sender, EventArgs e)
        {
            PushSlackOutput(4);
        }

        private void FixSlack5_Click(object sender, EventArgs e)
        {
            PushSlackOutput(5);
        }

        private void FixSlack6_Click(object sender, EventArgs e)
        {
            PushSlackOutput(6);
        }

        private void FixSlack7_Click(object sender, EventArgs e)
        {
            PushSlackOutput(7);
        }

        private void FixSlack8_Click(object sender, EventArgs e)
        {
            PushSlackOutput(8);
        }

        private void FixSlack9_Click(object sender, EventArgs e)
        {
            PushSlackOutput(9);
        }

        private void FixSlack10_Click(object sender, EventArgs e)
        {
            PushSlackOutput(10);
        }
    }
}
