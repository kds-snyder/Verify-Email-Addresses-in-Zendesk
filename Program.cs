using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.IO;
using VerifyZdUserEmailAddresses.ZdModels;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net;
using System.Text;
using System.Diagnostics;

namespace VerifyUserEmailAddresses
{
    class Program
    {
        private static long _numEmailsVerified = 0;
        private static long _numEmailsAlreadyVerified = 0;
        private static long _numEmailsNotVerified = 0;
        private static long _numEmailsProcessed = 0;
        static async Task Main(string[] args)
        {
            // Get input variables
            Console.Write("Enter API base for reading user data (e.g. https://xyz.zendesk.com): ");
            var userApiBase = Console.ReadLine();
            Console.Write("Enter email address for API access: ");
            var emailAddress = Console.ReadLine();
            Console.Write("Enter API token: ");
            var apiToken = Console.ReadLine();
            var apiCredentials = Convert.ToBase64String(Encoding.Default.GetBytes($"{emailAddress}/token:{apiToken}"));
            Console.Write("Enter Excel file name: ");
            var fileName = Console.ReadLine();
            Console.Write("Enter column containing the email addresses to verify: ");
            var emailAddrColIndex = Console.ReadLine();

            // Start processing of file
            var stopwatch = Stopwatch.StartNew();
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (SpreadsheetDocument xlsDocument = SpreadsheetDocument.Open(fs, false))
                    {
                        // Get sheet data
                        WorkbookPart wbPart = xlsDocument.WorkbookPart;
                        WorksheetPart worksheetPart = wbPart.WorksheetParts.First();
                        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                        // Get shared string table
                        SharedStringTablePart stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                        foreach (Row r in sheetData.Elements<Row>())
                        {
                            // Cell reference is column index + row index
                            var rowIndex = r.RowIndex;
                            var cellRef = emailAddrColIndex + rowIndex;

                            // Get the cell that contains the email address to verify
                            var cell = r.Elements<Cell>().FirstOrDefault(x => String.Equals(x.CellReference, cellRef, StringComparison.InvariantCultureIgnoreCase));

                            // Get the cell string value from the shared strings table                            
                            var cellString = stringTable.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;

                            //Validate email if cell contains email address (to avoid processing header row or empty rows)
                            if (cellString.Contains("@"))
                            {
                                await ZdVerifyEmail(userApiBase, apiCredentials, cellString, rowIndex);
                            }
                        }
                        stopwatch.Stop();
                        TimeSpan ts = stopwatch.Elapsed;
                        Console.WriteLine("");
                        Console.WriteLine($"Total # email addresses verified: {_numEmailsVerified}");
                        Console.WriteLine($"Total # email addresses already verified: {_numEmailsAlreadyVerified}");
                        Console.WriteLine($"Total # email addresses couldn't be verified: {_numEmailsNotVerified}");
                        Console.WriteLine($"Total # email addresses processed: {_numEmailsProcessed}");
                        Console.WriteLine($"Program duration time: {ts.Hours} hours, {ts.Minutes} minutes, {ts.Seconds} seconds, {ts.Milliseconds} milliseconds");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening file: {fileName}, Exception Message: {ex.Message}\nInnerException: {ex.InnerException}\nStack trace: {ex.StackTrace}");
            }
        }

        // Verify email by getting user ID, getting user identity corresponding to email, then verifying identity 
        static async Task ZdVerifyEmail(string userApiBase, string apiCredentials, string userEmailAddress, long rowIndex)
        {
            var VerifyInfo = new StringBuilder($"Row: {rowIndex}, user email address: {userEmailAddress}");
            ++_numEmailsProcessed;
            long userId = await ZdGetUserId(userApiBase, apiCredentials, userEmailAddress);
            if (userId > 0)
            {
                VerifyInfo.Append($", user ID: {userId}");

                var userIdentity = await ZdGetUserIdentity(userApiBase, apiCredentials, userId, userEmailAddress);
                if (!(userIdentity is null))
                {
                    VerifyInfo.Append($", user identity ID: {userIdentity.id}, verified setting: {userIdentity.verified}");
                    if (!userIdentity.verified)
                    {
                        bool verified = await ZdVerifyIdentity(userApiBase, apiCredentials, userIdentity);
                        if (verified)
                        {
                            VerifyInfo.Append($", email is now verified");
                            Console.WriteLine(VerifyInfo.ToString());
                            ++_numEmailsVerified;
                        }
                        else
                        {;
                            VerifyInfo.Append($", couldn't verify");
                            Console.WriteLine(VerifyInfo.ToString());
                            ++_numEmailsNotVerified;
                        }
                    }
                    else
                    {
                        VerifyInfo.Append($", email is already verified");
                        Console.WriteLine(VerifyInfo.ToString());
                        ++_numEmailsAlreadyVerified;
                    }
                }
                else
                {
                    VerifyInfo.Append($", couldn't get user identity");
                    Console.WriteLine(VerifyInfo.ToString());
                    ++_numEmailsNotVerified;
                }
            }
            else
            {
                VerifyInfo.Append($", couldn't get user ID");
                Console.WriteLine(VerifyInfo.ToString());
                ++_numEmailsNotVerified;
            }
        }

        // Get user ID corresponding to email
        static async Task<long> ZdGetUserId(string userApiBase, string apiCredentials, string userEmailAddress)
        {
            long userId = 0;
            try
            {
                using (var httpClient = new HttpClient())
                {
                    var requestUrl = $"{userApiBase}/api/v2/search.json?query=email:{userEmailAddress}";
                    var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("Authorization", $"Basic {apiCredentials}");

                    var response = await httpClient.SendAsync(request);
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        var contentString = await response.Content.ReadAsStringAsync();
                        var result = JsonConvert.DeserializeObject<ZdUsers>(contentString);

                        if (result.results.Any())
                        {
                            userId = result.results.First().id;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"GET failed to {requestUrl}, status code: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ZdGetUserId Exception Message: {ex.Message}\nInnerException: {ex.InnerException}\nStack trace: {ex.StackTrace}");
            }
            return userId;
        }

        // Get user identity corresponding to user ID and email address
        static async Task<ZdUserIdentity> ZdGetUserIdentity(string userApiBase, string apiCredentials, long userId, string userEmailAddress)
        {
            var userIdentity = new ZdUserIdentity();
            try
            {
                using (var httpClient = new HttpClient())
                {
                    var requestUrl = $"{userApiBase}/api/v2/users/{userId}/identities.json";
                    var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("Authorization", $"Basic {apiCredentials}");

                    var response = await httpClient.SendAsync(request);
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        var contentString = await response.Content.ReadAsStringAsync();
                        var result = JsonConvert.DeserializeObject<ZdUserIdentities>(contentString);

                        if (result.identities.Any())
                        {
                            userIdentity = result.identities.FirstOrDefault(x => String.Equals(x.value, userEmailAddress, StringComparison.InvariantCultureIgnoreCase));
                        }
                    }
                    else
                    {
                        Console.WriteLine($"GET failed to {requestUrl}, status code: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ZdGetUserIdentity Exception Message: {ex.Message}\nInnerException: {ex.InnerException}\nStack trace: {ex.StackTrace}");
            }
            return userIdentity;
        }


        // Verify user identity
        static async Task<bool> ZdVerifyIdentity(string userApiBase, string apiCredentials, ZdUserIdentity userIdentity)
        {
            bool verifySuccess = false;
            try
            {
                using (var httpClient = new HttpClient())
                {
                    var requestUrl = $"{userApiBase}/api/v2/users/{userIdentity.user_id}/identities/{userIdentity.id}.json";
                    var request = new HttpRequestMessage(HttpMethod.Put, requestUrl);
                    request.Headers.Add("Authorization", $"Basic {apiCredentials}");

                    var zdUpdateUserIdentity = new ZdUserIdentityUpdate
                    {
                        identity = new ZdVerified
                        {
                            verified = true
                        }
                    };
                    var jsonRequest = JsonConvert.SerializeObject(zdUpdateUserIdentity, Formatting.None);
                    request.Content = new StringContent(jsonRequest, System.Text.Encoding.UTF8, "application/json");

                    var response = await httpClient.SendAsync(request);
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        verifySuccess = true;
                    }
                    else
                    {
                        Console.WriteLine($"PUT failed to {requestUrl}, status code: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ZdVerifyIdentity Exception Message: {ex.Message}\nInnerException: {ex.InnerException}\nStack trace: {ex.StackTrace}");
            }
            return verifySuccess;
        }
    }
}
