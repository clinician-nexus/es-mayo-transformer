namespace MayoExcelTransformer;

using System;
using OfficeOpenXml;
using System.IO;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

class Program
{
  static void Main(string[] args)
  {
    try
    {
      string sourceFilePath = args.Length > 0 ? args[0] : "input.xlsx";
      string destinationFilePath = sourceFilePath.Replace(".xlsx", $"-transformed-{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

      Console.WriteLine($"Reading input file {sourceFilePath}");
      using (var package = new ExcelPackage(new FileInfo(sourceFilePath)))
      {
        var worksheet = package.Workbook.Worksheets[0]; // Assuming you want to work with the first worksheet

        // read column indexes from first row
        var columnIndexes = new Dictionary<string, int>();
        for (int column = 1; column <= worksheet.Dimension.End.Column; column++)
        {
          try
          {
            var columnName = worksheet.Cells[1, column].Value?.ToString().ToLower();
            if (columnName != null)
            {
              columnIndexes.Add(columnName, column);
            }
          }
          catch (Exception e)
          {
            Console.WriteLine($"Error reading column name at column {column}: {e.Message}");
          }
        }

        // create a function to get the value of a cell in a row by column name
        Func<int, string, string> getCellValue = (row, columnName) =>
        {
          var regex = new Regex(columnName, RegexOptions.IgnoreCase);
          var value = columnIndexes
            .Where(x => regex.IsMatch(x.Key))
            .Select(x => worksheet.Cells[row, x.Value].Value?.ToString())
            .FirstOrDefault(x => x != null);
          return value;
        };

        // read all rows in the first worksheet into a list of InputRow objects
        var inputRows = new List<InputRow>();
        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
          var inputRow = new InputRow
          {
            StepName = getCellValue(row, "Step name"),
            MatchId = getCellValue(row, "Match id"),
            UserName = getCellValue(row, "User name"),
            UserEmail = getCellValue(row, "User email"),
            CompletedDate = getCellValue(row, "Completed date"),
            CoordinatorNonEmployeeCategory = getCellValue(row, "form Coordinator Step .+ Non-Employee Category"),
            CoordinatorTypeOfNonEmployee = getCellValue(row, "form Coordinator Step .+ Type of Non-Employee"),
            StudentLegalFirstName = getCellValue(row, "form Student Demographic Data .+ Legal First Name"),
            StudentLegalLastName = getCellValue(row, "form Student Demographic Data .+ Legal Last Name"),
            StudentPreferredName = getCellValue(row, "form Student Demographic Data .+ Preferred Name"),
            StudentMiddleName = getCellValue(row, "form Student Demographic Data .+ Middle Name"),
            StudentSuffix = getCellValue(row, "form Student Demographic Data .+ Suffix"),
            StudentGender = getCellValue(row, "form Student Demographic Data .+ Gender"),
            StudentDateOfBirth = getCellValue(row, "form Student Demographic Data .+ Date of Birth"),
            StudentHomeAddressLine1 = getCellValue(row, "form Student Demographic Data .+ Home Address Line 1"),
            StudentHomeAddressLine2 = getCellValue(row, "form Student Demographic Data .+ Home Address Line 2"),
            StudentCity = getCellValue(row, "form Student Demographic Data .+ City"),
            StudentState = getCellValue(row, "form Student Demographic Data .+ State"),
            StudentCountryOfResidence = getCellValue(row, "form Student Demographic Data .+ Country of Residential Address"),
            StudentZipOrPostalCode = getCellValue(row, "form Student Demographic Data .+ Zip or Postal Code"),
            StudentCountryOfCitizenship = getCellValue(row, "form Student Demographic Data .+ What is your country of citizenship?"),
            StudentCountryOfOrigin = getCellValue(row, "form Student Demographic Data .+ Country of Origin"),
            StudentVisaType = getCellValue(row, "form Student Demographic Data .+ Visa Type"),
            CoordinatorWorkState = getCellValue(row, "form Coordinator Step .+ Work State"),
            StudentSchoolEmailAddress = getCellValue(row, "form Student Demographic Data .+ School Email Address"),
            StudentPrimaryPhoneNumber = getCellValue(row, "form Student Demographic Data .+ Primary Phone Number"),
            MatchStartDate = getCellValue(row, "Match start date"),
            MatchEndDate = getCellValue(row, "Match end date"),
            CoordinatorCampus = getCellValue(row, "form Coordinator Step .+ Campus"),
            CoordinatorLengthOfStay = getCellValue(row, "form Coordinator Step .+ Length of Stay"),
            CoordinatorSupervisorCode = getCellValue(row, "form Coordinator Step .+ Supervisor Code/Reporting Unit"),
            CoordinatorSupervisorPerID = getCellValue(row, "form Coordinator Step .+ Supervisor Per ID"),
          };
          inputRows.Add(inputRow);
        }

        // group by matchid and filter by:
        // - has a row with stepname "Student Demographic Data" with non-null completed date
        // - has a row with stepname "Coordinator Step" with non-null completed date
        var matchGroups = inputRows.GroupBy(x => x.MatchId);
        var matchGroupsWithStepsCompleted = matchGroups.Where(x =>
              x.Any(y => y.StepName == "Student Demographic Data" && y.CompletedDate != null)
          && x.Any(y => y.StepName == "Coordinator Step" && y.CompletedDate != null)
        );

        Console.WriteLine($"Found {inputRows.Count} input rows");
        Console.WriteLine($"Found {matchGroups.Count()} distinct matches");
        Console.WriteLine($"Found {matchGroupsWithStepsCompleted.Count()} matches with completed Student Demographic Data and Coordinator Step");

        // Generate 1 output row for each matchGroup
        var outputRows = new List<Dictionary<string, string?>>();
        foreach (var matchGroup in matchGroupsWithStepsCompleted)
        {
          var coordinatorStepRow = matchGroup.First(x => x.StepName == "Coordinator Step" && x.CompletedDate != null);
          var studentDemographicDataRow = matchGroup.First(x => x.StepName == "Student Demographic Data" && x.CompletedDate != null);
          var outputRow = new Dictionary<string, string?> {
            { "Non-Employee Category", coordinatorStepRow.CoordinatorNonEmployeeCategory },
            { "Type of Non-Employee", coordinatorStepRow.CoordinatorTypeOfNonEmployee },
            { "Legal First Name", studentDemographicDataRow.StudentLegalFirstName },
            { "Legal Last Name", studentDemographicDataRow.StudentLegalLastName },
            { "Gender", studentDemographicDataRow.StudentGender },
            { "Date of Birth", studentDemographicDataRow.StudentDateOfBirth },
            { "Country of Residential Address", studentDemographicDataRow.StudentCountryOfResidence },
            { "Street Address Line 1", studentDemographicDataRow.StudentHomeAddressLine1 },
            { "City", studentDemographicDataRow.StudentCity },
            { "Zip or Postal Code", studentDemographicDataRow.StudentZipOrPostalCode },
            { "Country of Citizenship", studentDemographicDataRow.StudentCountryOfCitizenship },
            { "Country of Origin", studentDemographicDataRow.StudentCountryOfOrigin },
            { "Visa Type", studentDemographicDataRow.StudentVisaType },
            { "Current Employer", "" },
            { "Work State", coordinatorStepRow.CoordinatorWorkState },
            { "Personal Email", studentDemographicDataRow.StudentSchoolEmailAddress ?? studentDemographicDataRow.UserEmail },
            { "Personal Phone Number", studentDemographicDataRow.StudentPrimaryPhoneNumber },
            { "Start Date", studentDemographicDataRow.MatchStartDate },
            { "Campus", coordinatorStepRow.CoordinatorCampus },
            { "Length of Stay", coordinatorStepRow.CoordinatorLengthOfStay },
            { "Supervisor Code / Reporting Unit", coordinatorStepRow.CoordinatorSupervisorCode },
            { "Administrator Per ID", coordinatorStepRow.CoordinatorSupervisorPerID },
            { "Person ID", "" },
            { "Preferred Name", studentDemographicDataRow.StudentPreferredName },
            { "Middle Name", studentDemographicDataRow.StudentMiddleName },
            { "Suffix", studentDemographicDataRow.StudentSuffix },
            { "Street Address Line 2", studentDemographicDataRow.StudentHomeAddressLine2 },
            { "State", studentDemographicDataRow.StudentState },
            { "Province", "" },
            { "Secondary Phone Number", "" },
            { "Country of Dual Citizenship", "" },
            { "Other Visa Type", "" },
            { "Other Current Employer", "" },
            { "Non-Mayo Clinic School Name", "" },
            { "Other Non-Mayo Clinic School Name", "" },
            { "Building", "" },
            { "Floor", "" },
            { "Licensed Provider", "" },
            { "Credentials", "" },
            { "Credentials verified by", "" },
            { "Non-US Government Official Status", "" },
            { "Job Title", "" },
            { "Primary Work Location", "" },
            { "End Date", studentDemographicDataRow.MatchEndDate },
            { "Non-US Address 1", "" },
            { "Non-US Address 2", "" },
            { "Non-US Address 3", "" },
            { "Non-US Address 4", "" },
            { "Non-US Address 5", "" },
            { "EAM ONLY FIELD: Group 1", "" },
            { "EAM ONLY FIELD: Group 2", "" },
            { "EAM ONLY FIELD: Group 3", "" },
            { "EAM ONLY FIELD: Group 4", "" },
            { "EAM ONLY FIELD: Group 5", "" }
          };

          outputRows.Add(outputRow);
        }

        // Generate output file
        Console.WriteLine($"Generating output file {destinationFilePath} with {outputRows.Count} rows");

        using var newPackage = new ExcelPackage();
        var ws = newPackage.Workbook.Worksheets.Add("Sheet 1");
        var headers = new string[] {
          "Non-Employee Category",
          "Type of Non-Employee",
          "Legal First Name",
          "Legal Last Name",
          "Gender",
          "Date of Birth",
          "Country of Residential Address",
          "Street Address Line 1",
          "City",
          "Zip or Postal Code",
          "Country of Citizenship",
          "Country of Origin",
          "Visa Type",
          "Current Employer",
          "Work State",
          "Personal Email",
          "Personal Phone Number",
          "Start Date",
          "Campus",
          "Length of Stay",
          "Supervisor Code / Reporting Unit",
          "Administrator Per ID",
          "Person ID",
          "Preferred Name",
          "Middle Name",
          "Suffix",
          "Street Address Line 2",
          "State",
          "Province",
          "Secondary Phone Number",
          "Country of Dual Citizenship",
          "Other Visa Type",
          "Other Current Employer",
          "Non-Mayo Clinic School Name",
          "Other Non-Mayo Clinic School Name",
          "Building",
          "Floor",
          "Licensed Provider",
          "Credentials",
          "Credentials verified by",
          "Non-US Government Official Status",
          "Job Title",
          "Primary Work Location",
          "End Date",
          "Non-US Address 1",
          "Non-US Address 2",
          "Non-US Address 3",
          "Non-US Address 4",
          "Non-US Address 5",
          "EAM ONLY FIELD: Group 1",
          "EAM ONLY FIELD: Group 2",
          "EAM ONLY FIELD: Group 3",
          "EAM ONLY FIELD: Group 4",
          "EAM ONLY FIELD: Group 5"
        };

        // write header row
        for (int i = 0; i < headers.Length; i++)
        {
          ws.Cells[1, i + 1].Value = headers[i];
        }

        // write output rows
        for (int i = 0; i < outputRows.Count; i++)
        {
          var outputRow = outputRows[i];
          for (int j = 0; j < headers.Length; j++)
          {
            ws.Cells[i + 2, j + 1].Value = outputRow[headers[j]];
          }
        }

        // Save the new Excel file
        var fileInfo = new FileInfo(destinationFilePath);
        newPackage.SaveAs(fileInfo);
      }
    }
    catch (Exception e)
    {
      Console.WriteLine($"Error: {e.Message}");
    }

    // readline to wait for user to do something before exiting
    Console.WriteLine("Press any key to exit...");
    Console.ReadKey();
  }
}
