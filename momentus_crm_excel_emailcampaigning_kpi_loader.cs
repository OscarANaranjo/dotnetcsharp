// Script: momentus_crm_excel_emailcampaign_kpi_loader.cs
// Purpose: Load email campaign KPIs (opens, clicks) from Excel CSV into Momentus CRM via Ungerboeck API
// Technologies: C#, .NET Framework, CsvHelper, Ungerboeck SDK
// Business Impact: Automated ingestion of email engagement metrics into CRM campaigns to identify hot leads and reduce manual data entry
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using Ungerboeck.Api.Sdk;
using Ungerboeck.Api.Models.Authorization;
using Ungerboeck.Api.Models.Subjects;

// Represents a single KPI record from the CSV file
public class CsvKpiRecord
{
    public string? Account { get; set; }       // CRM Account Code
    public string? Name { get; set; }          // Contact Name
    public string? Email { get; set; }         // Contact Email
    public string? Company { get; set; }       // Company Name
    public string? Title { get; set; }         // Job Title
    public int? Opens { get; set; }            // Email Opens
    public int? Clicks { get; set; }           // Email Clicks
}

public class Program
{
    // Authentication credentials (masked for public repo)
    private static string apiUserId = "YOUR_API_USER_ID";
    private static string apiKey = "YOUR_API_KEY";
    private static string apiSecret = "YOUR_API_SECRET";
    private static string ungerboeckUri = "https://your-instance.ungerboeck.com";
    private static string csvFilePath = "ExcelKPILoad.csv"; // Path to CSV file

    public static void Main(string[] args)
    {
        // Authenticate with Ungerboeck API
        var auth = GetApiClientAuth();
        var apiClient = new ApiClient(auth);

        // Read KPI records from CSV
        var kpis = ReadKpisFromCsv(csvFilePath);

        // Update campaign metadata (PUT)
        var campaign = new CampaignsModel
        {
            ID = "999",                      // Campaign ID in CRM
            OrganizationCode = "00",         // Org code in Momentus
            Designation = "C",               // Campaign type
            Description = "CAMPAIGN_NAME"       // Campaign name
        };

        apiClient.Endpoints.Campaigns.Update(campaign);
        Console.WriteLine($"Updated Campaign ID: {campaign.ID}");

        // Loop through each KPI record and add it to the campaign
        foreach (var kpi in kpis)
        {
            var campaignDetail = new CampaignDetailsModel
            {
                OrganizationCode = "00",
                CampaignDesignation = campaign.Designation,
                Campaign = campaign.ID,
                Account = kpi.Account,
                OutgoingCalls = kpi.Opens ?? 0,   // Map opens to CRM field
                Mailings = kpi.Clicks ?? 0        // Map clicks to CRM field
            };

            apiClient.Endpoints.CampaignDetails.Add(campaignDetail); // POST new record
        }
    }

    // Builds the JWT authentication object
    private static Jwt GetApiClientAuth()
    {
        return new Jwt
        {
            APIUserID = apiUserId,
            Key = apiKey,
            Secret = apiSecret,
            UngerboeckURI = ungerboeckUri,
            Expiration = DateTime.UtcNow.AddMinutes(5)
        };
    }

    // Reads KPI records from the specified CSV file
    private static List<CsvKpiRecord> ReadKpisFromCsv(string filePath)
    {
        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HeaderValidated = null,
            MissingFieldFound = null
        });

        return csv.GetRecords<CsvKpiRecord>().ToList();
    }
}
