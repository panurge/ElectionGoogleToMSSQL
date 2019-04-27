using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Data;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Net;
using System.Diagnostics;
using System.Xml;

namespace SheetsQuickstart
{

    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        static string connectionString = "Data Source=WLSNEWS-NAS04;Initial Catalog=Elections2017 ;User ID=sa;Password=armalyte";
        static SqlConnection conn = new SqlConnection(connectionString);
        static SqlCommand cmd = new SqlCommand("", conn);

        static void Main(string[] args)
        {
            //webScrape();
            // WebClient webClient = new WebClient();
            // webClient.Credentials = new System.Net.NetworkCredential("UserName", "Password", "Domain");
            //// string pageHTML = webClient.DownloadString("https://www.bbc.co.uk/news/politics/wales-constituencies/W09000022");
            // string pageHTML = webClient.DownloadString("http://10.176.4.109:8080/Default.aspx");
            var s = @"<span class=""exchangeable"" data-value=' *Segno Preliminare Quarto Orario 20190412* ' data-names='[""Titolo""]' data-detail=""Titolo"" data-detail-value=""Segno Preliminare Quarto Orario 20190412"">Segno Preliminare Quarto Orario 20190412</span>";
            var doc = HtmlNode.CreateNode(s);
            Debug.WriteLine(doc.GetAttributeValue("data-names", "default"));
            Debug.WriteLine(doc.InnerText);
            // var doch = new HtmlDocument();
            // //doch.LoadHtml(@"<html><body><div id='foo'>text</div></body></html>");
            // //var div = doch.GetElementbyId("foo");
            // doch.LoadHtml(pageHTML);
            // //var doc = HtmlNode.CreateNode(doch.ToString());
            // XmlNamespaceManager oManager = new XmlNamespaceManager(new NameTable());
            // oManager.AddNamespace("ns", "http://www.w2.org/1999/xhtml");
            // //*[@id="Repeater1_lblRegion2"]
            // //*[@id='Repeater1_lblRegion2']";
            // //*[@id="myTable"]/tbody/tr[2]/td[1]/a
            // string nodestr = "//*[@id='myTable']/tbody/tr[2]/td[1]/a";
            // nodestr= "//*[@id='myTable']/tbody/tr[2]/td[1]/a";
            // nodestr = "//div[@class='menu']//text()[normalize-space()])[last()]";
            //*[@id="myTable"]/tbody/tr[2]/td[2]
            //*[@id="myTable"]/tbody/tr[2]/td[1]/a
            //*[@id="myTable"]/tbody/tr[2]/td[1]/a
            ;            //"*[@id=\"wales_constituency_result_table\"]/div/table/tbody/tr[1]/td[2]/span[2]";
                         //@"(//*[@id='page']/div[2]/div/div[1]/h1"
                         // HtmlNode node = doch.DocumentNode.SelectSingleNode(nodestr);
                         //HtmlNode node = doch.DocumentNode.SelectSingleNode(nodestr);
                         //Debug.WriteLine(node.XPath.ToString());
                         // Show info
                         //System.Console.WriteLine(div.OuterHtml);

            //// Show info
            //FiddleHelper.WriteTable(new List<string>() { div.OuterHtml });

            //// Show info
            //FiddleHelper.WriteTable(new List<HtmlAgilityPack.HtmlNode>() { div });

            //Console.WriteLine("Awaiting Input");
            //Console.Read();
            //loadOfficialDataRotated();
        }
        static void webScrape()
        {
            HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create("http://10.176.4.109:8080/Default.aspx");
            HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
            var stream = myHttpWebResponse.GetResponseStream();
            var reader = new StreamReader(stream);
            string html = reader.ReadToEnd();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(html);
            XmlNamespaceManager oManager = new XmlNamespaceManager(new NameTable());
            oManager.AddNamespace("ns", "http://www.w2.org/1999/xhtml");
            doc.SelectSingleNode("/ns:uclassify/ns:readCalls/ns:classify/ns:classification/ns:class[1]/@p", oManager);
            Debug.WriteLine(doc.ToString());
        }
        static void loadOfficialDataRotated()
        {
            cmd.CommandText = @"insert into OfficialConstituencyDataRotated (ConstituencyID,Constituency,PartyID,Result)values( @ConstituencyID,@Constituency,@PartyID,@Result)";
            cmd.Parameters.Add("@ConstituencyID", SqlDbType.Int);
            cmd.Parameters.Add("@Constituency", SqlDbType.VarChar);
            cmd.Parameters.Add("@PartyID", SqlDbType.Int);
            cmd.Parameters.Add("@Result", SqlDbType.Int);

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = "1Pgi8O5N9JnzXzN9TIU7VQAWZi-ApmFcsF3eEGt2j8Ww";
            // String range = "Class Data!A2:AQ";
            String range = "NAW2016 CONST Transposed!A1:AQ";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            conn.Open();
            if (values != null && values.Count > 0)
            {
                for (int i = 3; i < values[0].Count; i++)
                {
                    Console.Write(values[0][i] + ", ");
                    cmd.Parameters["@ConstituencyID"].Value = values[0][i];
                    Console.WriteLine(values[3][i] + ", ");
                    cmd.Parameters["@Constituency"].Value = values[3][i];
                    for (int j = 4; j < 12; j++) {
                        Console.Write(values[j][0] + ", ");
                        cmd.Parameters["@PartyID"].Value = values[j][0];
                        Console.WriteLine(values[j][i] + ", ");
                        cmd.Parameters["@Result"].Value = values[j][i];
                        cmd.ExecuteNonQuery();
                    }
                }
                //foreach (var row in values)
                //{
                //    Console.WriteLine(row[0]);

                //}
            }
            Console.WriteLine("Awaiting Input");
            Console.Read();

        }
        static void loadOfficialData()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            conn.Open();
            //cmd.CommandText = @"Update Constituencies set ONS = @ONS, PA = @PA where ConstituencyID = @ConstituencyID";
            cmd.CommandText = @"insert into OfficialConstituencyData (ONS,PA,Constituency,CON,LAB,LIB,PC,UKIP,GRN,IND,OTH) values (@ONS,@PA,@Constituency,@CON,@LAB,@LIB,@PC,@UKIP,@GRN,@IND,@OTH)"; 
            cmd.Parameters.Add("@ConstituencyID", SqlDbType.Int);
            cmd.Parameters.Add("@ONS", SqlDbType.VarChar);
            cmd.Parameters.Add("@PA", SqlDbType.Int);
            cmd.Parameters.Add("@Constituency", SqlDbType.VarChar);
            cmd.Parameters.Add("@CON", SqlDbType.Int);
            cmd.Parameters.Add("@LAB", SqlDbType.Int);
            cmd.Parameters.Add("@LIB", SqlDbType.Int);
            cmd.Parameters.Add("@PC", SqlDbType.Int);
            cmd.Parameters.Add("@UKIP", SqlDbType.Int);
            cmd.Parameters.Add("@GRN", SqlDbType.Int);
            cmd.Parameters.Add("@IND", SqlDbType.Int);
            cmd.Parameters.Add("@OTH", SqlDbType.Int);
            //            cmd.Parameters.Add("@PA", SqlDbType.Int);
            //cmd.Parameters.Add("@Forenames", SqlDbType.NVarChar);

            // Define request parameters.
            // String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String spreadsheetId = "1Pgi8O5N9JnzXzN9TIU7VQAWZi-ApmFcsF3eEGt2j8Ww";
            // String range = "Class Data!A2:E";
            String range = "NAW2016 CONST!A3:L";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    try
                    {
                        cmd.Parameters["@ConstituencyID"].Value = row[0];
                        cmd.Parameters["@ONS"].Value = row[1];
                        cmd.Parameters["@PA"].Value = row[2];
                        cmd.Parameters["@Constituency"].Value = row[3];
                        cmd.Parameters["@CON"].Value = row[4];
                        cmd.Parameters["@LAB"].Value = row[5];
                        cmd.Parameters["@LIB"].Value = row[6];
                        cmd.Parameters["@PC"].Value = row[7];
                        if (row.Count > 8)
                            cmd.Parameters["@UKIP"].Value = StrToIntDef((string)row[8]);
                        else
                            cmd.Parameters["@UKIP"].Value = 0;
                        if (row.Count > 9)
                            cmd.Parameters["@GRN"].Value = StrToIntDef((string)row[9]);
                        else
                            cmd.Parameters["@GRN"].Value = 0;
                        if (row.Count > 10)
                           cmd.Parameters["@IND"].Value = StrToIntDef((string)row[10]);
                        else
                            cmd.Parameters["@IND"].Value = 0;
                        if (row.Count > 11)
                            cmd.Parameters["@OTH"].Value = StrToIntDef((string)row[11]);
                        else
                            cmd.Parameters["@OTH"].Value = 0;
                        //else
                        //cmd.Parameters["@IND"].Value = row[10];
                        //cmd.Parameters["@OTH"].Value = row[11];
                        //cmd.Parameters["@PA"].Value = row[2];
                        //cmd.Parameters["@Forenames"].Value = row[3];
                        cmd.ExecuteNonQuery();
                        // Print columns A and E, which correspond to indices 0 and 4.
                        for (int i = 0; i< row.Count; i++)
                        Console.Write("{0}, ", row[i] );
                        Console.WriteLine();
                    }
                    catch (SqlException e) { Console.WriteLine("CONSTITUENCY DUPLICATE " + e.Message); }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            conn.Close();
            Console.WriteLine("Awaiting Input");
            Console.Read();
        }
        public static int StrToIntDef(string s)
        {
            int number;
            if (int.TryParse(s, out number))
                return number;
            return 0;
        }
        public void updateCandidates()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            // String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String spreadsheetId = "1Pgi8O5N9JnzXzN9TIU7VQAWZi-ApmFcsF3eEGt2j8Ww";
            // String range = "Class Data!A2:E";
            String range = "Local2021!A2:D";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            //https://docs.google.com/spreadsheets/d/1Pgi8O5N9JnzXzN9TIU7VQAWZi-ApmFcsF3eEGt2j8Ww/edit?usp=sharing
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit

            conn.Open();
            cmd.CommandText = @"insert into localcandidates2021 (Constituency, Party, Surname, Forenames) 
                                        Values(@Constituency, @Party, @Surname, @Forenames)";
            cmd.Parameters.Add("@Constituency", SqlDbType.NVarChar);
            cmd.Parameters.Add("@Party", SqlDbType.NVarChar);
            cmd.Parameters.Add("@Surname", SqlDbType.NVarChar);
            cmd.Parameters.Add("@Forenames", SqlDbType.NVarChar);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    try
                    {
                        cmd.Parameters["@Constituency"].Value = row[0];
                        cmd.Parameters["@Party"].Value = row[1];
                        cmd.Parameters["@Surname"].Value = row[2];
                        cmd.Parameters["@Forenames"].Value = row[3];
                        cmd.ExecuteNonQuery();
                        // Print columns A and E, which correspond to indices 0 and 4.
                        Console.WriteLine("{0}, {1}, {2}, {3}", row[0], row[1], row[2], row[3]);
                    }
                    catch(SqlException e) { Console.WriteLine("CONSTITUENCY DUPLICATE " + e.Message); }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }

            cmd.Parameters.RemoveAt("@Constituency");
            cmd.Parameters.Add("@Region", SqlDbType.NVarChar);
            cmd.Parameters.Add("@GroupID", SqlDbType.NVarChar);
            cmd.CommandText = @"insert into regionalcandidates2021 (Region, GroupID, Party, Surname, Forenames) 
                                        Values(@Region, @GroupID, @Party, @Surname, @Forenames)";

             range = "Regional2021!A2:E";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            response = request.Execute();
            values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    try
                    {
                        cmd.Parameters["@Region"].Value = row[0];
                        cmd.Parameters["@Party"].Value = row[1];
                        cmd.Parameters["@GroupID"].Value = row[2];
                        cmd.Parameters["@Surname"].Value = row[3];
                        cmd.Parameters["@Forenames"].Value = row[4];
                        cmd.ExecuteNonQuery();
                        Console.WriteLine("{0}, {1}, {2}, {3}, {4}", row[0], row[1], row[2], row[3], row[4]);
                    }
                    catch(SqlException e) { Console.WriteLine("REGION DUPLICATE " + e.Message); }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }

            //cmd.Parameters.RemoveAt("@Constituency");
            cmd.Parameters.RemoveAt("@Region");
            cmd.Parameters.RemoveAt("@GroupID");
            //cmd.Parameters.RemoveAt("@Constituency");
            cmd.Parameters.RemoveAt("@Party");
            cmd.Parameters.RemoveAt("@Surname");
            cmd.Parameters.RemoveAt("@Forenames");

            Console.WriteLine("Starting Fixing Tables");
            cmd.CommandText = @"USE Elections2017
                                update regionalCandidates2021 set RegionID = (select  RegionID from Elections2016.dbo.Regions where region = regionalCandidates2021.Region)
                                update regionalCandidates2021 set PartyID = (select PartyID from Elections2016.dbo.Parties where party = regionalCandidates2021.Party)
                                update regionalCandidates2021 set Candidate = Forenames + ' ' + Surname
                                update regionalCandidates2021 set DisplayName  = UPPER(Surname) +', '+Forenames
                                update localCandidates2021 set ConstituencyID = (select ConstituencyID from Elections2016.dbo.Constituencies where constituency = localCandidates2021.Constituency)
                                update localCandidates2021 set PartyID = (select PartyID from Elections2016.dbo.Parties where party = localCandidates2021.Party)
                                update localCandidates2021 set Candidate = Forenames + ' ' + Surname
                                update localCandidates2021 set DisplayName  = UPPER(Surname) +', '+Forenames

                                DECLARE @RegionalCandidateId int
                                DECLARE @LocalCandidateId int
                                DECLARE @PartyID int
                                DECLARE @ConstituencyID int
                                DECLARE @RegionID int
                                DECLARE @Win nvarchar
                                DECLARE @Perc float
                                DECLARE @Result float

                                DECLARE MY_CURSOR CURSOR 
                                    LOCAL STATIC READ_ONLY FORWARD_ONLY
                                FOR 
                                SELECT DISTINCT LocalCandidateID 
                                FROM LocalCandidates2021

                                OPEN MY_CURSOR
                                FETCH NEXT FROM MY_CURSOR INTO @LocalCandidateId
                                WHILE @@FETCH_STATUS = 0
                                BEGIN 
                                    --Do something with Id here
	                                set @PartyID = (select PartyID from LocalCandidates2021 where LocalCandidateid = @LocalCandidateId)
	                                set @ConstituencyID = (select ConstituencyID from LocalCandidates2021 where LocalCandidateid = @LocalCandidateId)
	                                set @Win = (select Win from Elections2016.dbo.LocalCandidates where PartyID = @PartyID and ConstituencyID = @ConstituencyID)
	                                set @Result = (select result from Elections2016.dbo.LocalCandidates where PartyID = @PartyID and ConstituencyID = @ConstituencyID)
	                                set @Perc = (select perc from Elections2016.dbo.LocalCandidates where PartyID = @PartyID and ConstituencyID = @ConstituencyID)
	                                update LocalCandidates2021 set WinLast = @Win, PercLast = @Perc, ResultLast = @Result
	                                where LocalCandidateId = @LocalCandidateID
                                    PRINT CAST( @LocalCandidateId AS nvarchar) + ', '+ CAST( @ConstituencyID AS nvarchar) + ', '+ CAST(@PartyID as nvarchar)  + ', '+  @Win
                                    FETCH NEXT FROM MY_CURSOR INTO @LocalCandidateId
                                END
                                CLOSE MY_CURSOR
                                DEALLOCATE MY_CURSOR

                                DECLARE MY_CURSOR CURSOR 
                                    LOCAL STATIC READ_ONLY FORWARD_ONLY
                                FOR 
                                SELECT DISTINCT RegionalCandidateID 
                                FROM RegionalCandidates2021

                                OPEN MY_CURSOR
                                FETCH NEXT FROM MY_CURSOR INTO @RegionalCandidateId
                                WHILE @@FETCH_STATUS = 0
                                BEGIN 
                                    --Do something with Id here
	                                set @PartyID = (select PartyID from RegionalCandidates2021 where RegionalCandidateid = @RegionalCandidateId)
	                                set @RegionID = (select RegionID from RegionalCandidates2021 where RegionalCandidateid = @RegionalCandidateId)
	                                --set @Win = (select Win from Elections2016.dbo.RegionalCandidateswhere PartyID = @PartyID and RegionID = @RegionID)
	                                set @Result = (select distinct result from Elections2016.dbo.RegionalCandidates where PartyID = @PartyID and RegionID = @RegionID)
	                                set @Perc = (select distinct perc from Elections2016.dbo.RegionalCandidates where PartyID = @PartyID and RegionID = @RegionID)
	                                update RegionalCandidates2021 set  PercLast = @Perc, ResultLast = @Result
	                                where RegionalCandidateId = @RegionalCandidateID
                                    PRINT CAST( @RegionalCandidateId AS nvarchar) + ', '+ CAST( @RegionID AS nvarchar) + ', '+ CAST(@PartyID as nvarchar)  + ', '+  @Win
                                    FETCH NEXT FROM MY_CURSOR INTO @RegionalCandidateId
                                END
                                CLOSE MY_CURSOR
                                DEALLOCATE MY_CURSOR";
            cmd.ExecuteNonQuery();
            conn.Close();
            Console.WriteLine("Finished Fixing Tables");
            Console.WriteLine("Awaiting Input");
            Console.Read();
        }
    }
}
//Copare Official with ITV 2016
//DECLARE @RegionalCandidateId int
//DECLARE @LocalCandidateId int
//DECLARE @PartyID int
//DECLARE @Party varchar(6)
//DECLARE @ConstituencyID int
//DECLARE @CandidateID int
//DECLARE @Constituency varchar(6)
//DECLARE @RegionID int
//DECLARE @Win nvarchar
//DECLARE @Perc float
//DECLARE @Result float
//DECLARE @Result1 float
//DECLARE @Result2 float
//Declare @Database1 nvarchar
//Declare @Database2 nvarchar

//DECLARE MY_CURSOR CURSOR
//LOCAL STATIC READ_ONLY FORWARD_ONLY
//FOR
//SELECT DISTINCT CandidateID
//FROM OfficialConstituencyDataRotated

//OPEN MY_CURSOR
//FETCH NEXT FROM MY_CURSOR INTO @CandidateId
//WHILE @@FETCH_STATUS = 0
//BEGIN
//--Do something with Id here
//--set @Database1 = "LocalCandidates2021"
//set @Result1 = (select result from Elections2017.dbo.OfficialConstituencyDataRotated where CandidateId=@CandidateId)
//--set @Party = (select Party from Elections2016.dbo.LocalCandidates where LocalCandidateid = @LocalCandidateId)
//set @PartyID = (select PartyID from Elections2017.dbo.OfficialConstituencyDataRotated where Candidateid = @CandidateId)
//--set @Constituency = (select CONSTITUENCY from LocalCandidates where LocalCandidateid = @LocalCandidateId)
//set @ConstituencyID = (select CONSTITUENCYID from  OfficialConstituencyDataRotated where Candidateid = @CandidateId)
//-----set @Win = (select Win from Elections2016.dbo.LocalCandidates where PartyID = @PartyID and ConstituencyID = @ConstituencyID)
//set @Result2 = (select result from Elections2017.dbo.LocalCandidates where PartyID = @PartyID and ConstituencyID = @ConstituencyID)
//--set @Perc = (select perc from Elections2016.dbo.LocalCandidates where PartyID = @PartyID and ConstituencyID = @ConstituencyID)
//--update LocalCandidates2021 set WinLast = @Win, PercLast = @Perc, ResultLast = @Result
//--where LocalCandidateId = @LocalCandidateID
//--IF @Result1 != @Result2
//--PRINT CAST(@LocalCandidateId AS nvarchar) + ', '+ @CONSTITUENCY + ', '+ @Party
//--		 + ', ' +CAST(@Result1 as nvarchar)  + ', ' + CAST(@Result2 as nvarchar);
//IF @Result1 != @Result2
//PRINT CAST(@ConstituencyID as nvarchar)  + ', '+  CAST(@Result1 as nvarchar)  + ', '+  CAST(@Result2 as nvarchar);
//FETCH NEXT FROM MY_CURSOR INTO @CandidateId
//END
//CLOSE MY_CURSOR
//DEALLOCATE MY_CURSOR