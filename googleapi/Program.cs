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
            UserCredential credential;

            cmd.CommandText = "update blah";
            
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
                    catch { Console.WriteLine("CONSTITUENCY DUPLICATE"); }
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
                    catch { Console.WriteLine("REGION DUPLICATE"); }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
           
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
            conn.Close();
            Console.Read();
        }
    }
}