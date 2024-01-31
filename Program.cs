namespace MayoExcelTransformer;

using System;
using OfficeOpenXml;
using System.IO;
using System.Text.Json;

class Program
{
  static void Main(string[] args)
  {
    string sourceFilePath = args.Length > 0 ? args[0] : "input.xlsx"; //"path/to/your/source/excel.xlsx";
    string destinationFilePath = sourceFilePath.Replace(".xlsx", $"-transformed-{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

    using (var package = new ExcelPackage(new FileInfo(sourceFilePath)))
    {
      Console.WriteLine("worksheets: ", package.Workbook.Worksheets.Count);
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
      Func<ExcelWorksheet, int, string, string> getCellValue = (worksheet, row, columnName) =>
      {
        var columnIndex = columnIndexes[columnName.ToLower()];
        try
        {
          return worksheet.Cells[row, columnIndex].Value?.ToString();
        }
        catch (Exception)
        {
          return "";
        }
      };

      // read all rows in the first worksheet into a list of InputRow objects
      var inputRows = new List<InputRow>();
      for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
      {
        var inputRow = new InputRow
        {
          StepName = getCellValue(worksheet, row, "Step name"),
          MatchId = getCellValue(worksheet, row, "match id"),
          UserName = getCellValue(worksheet, row, "User name"),
          UserEmail = getCellValue(worksheet, row, "User email"),
          CompletedDate = getCellValue(worksheet, row, "Completed date"),
          CoordinatorNonEmployeeCategory = getCellValue(worksheet, row, "form Coordinator Step 1) Non-Employee Category"),
          CoordinatorTypeOfNonEmployee = getCellValue(worksheet, row, "form Coordinator Step 3) Type of Non-Employee"),
          StudentLegalFirstName = getCellValue(worksheet, row, "form Student Demographic Data 1) Legal First Name"),
          StudentLegalLastName = getCellValue(worksheet, row, "form Student Demographic Data 2) Legal Last Name"),
          StudentPreferredName = getCellValue(worksheet, row, "form Student Demographic Data 3) Preferred Name"),
          StudentGender = getCellValue(worksheet, row, "form Student Demographic Data 4) Gender"),
          StudentDateOfBirth = getCellValue(worksheet, row, "form Student Demographic Data 5) Date of Birth"),
          StudentHomeAddress = getCellValue(worksheet, row, "form Student Demographic Data 6) Home Address"),
          StudentCity = getCellValue(worksheet, row, "form Student Demographic Data 7) City"),
          StudentCountryOfResidence = getCellValue(worksheet, row, "form Student Demographic Data 8) Country of Residential Address"),
          StudentZipOrPostalCode = getCellValue(worksheet, row, "form Student Demographic Data 9) Zip or Postal Code"),
          StudentCountryOfCitizenship = getCellValue(worksheet, row, "form Student Demographic Data 18) What is your country of citizenship?"),
          StudentCountryOfOrigin = getCellValue(worksheet, row, "form Student Demographic Data 16) Country of Origin"),
          StudentVisaType = getCellValue(worksheet, row, "form Student Demographic Data 16) Visa Type"),
          CoordinatorWorkState = getCellValue(worksheet, row, "form Coordinator Step 4) Work State"),
          StudentSchoolEmailAddress = getCellValue(worksheet, row, "form Student Demographic Data 10) School Email Address"),
          StudentPrimaryPhoneNumber = getCellValue(worksheet, row, "form Student Demographic Data 11) Primary Phone Number"),
          MatchStartDate = getCellValue(worksheet, row, "Match start date"),
          MatchEndDate = getCellValue(worksheet, row, "Match end date"),
          CoordinatorCampus = getCellValue(worksheet, row, "form Coordinator Step 2) Campus"),
          CoordinatorLengthOfStay = getCellValue(worksheet, row, "form Coordinator Step 5) Length of Stay"),
          CoordinatorSupervisorCode = getCellValue(worksheet, row, "form Coordinator Step 6) Supervisor Code/Reporting Unit"),
          CoordinatorSupervisorPerID = getCellValue(worksheet, row, "form Coordinator Step 7) Supervisor Per ID"),

          // PersonID = getCellValue(worksheet, row, ""),
          // MiddleName = getCellValue(worksheet, row, ""),
          // Suffix = getCellValue(worksheet, row, ""),
          // StreetAddressLine2 = getCellValue(worksheet, row, ""),
          // State = getCellValue(worksheet, row, ""), 
          // Province = getCellValue(worksheet, row, ""),
          // SecondaryPhoneNumber = getCellValue(worksheet, row, ""),
          // CountryOfDualCitizenship = getCellValue(worksheet, row, ""),
          // OtherVisaType = getCellValue(worksheet, row, ""),
          // OtherCurrentEmployer = getCellValue(worksheet, row, ""),
          // NonMayoClinicSchoolName = getCellValue(worksheet, row, ""),
          // OtherNonMayoClinicSchoolName = getCellValue(worksheet, row, ""),
          // Building = getCellValue(worksheet, row, ""),
          // Floor = getCellValue(worksheet, row, ""),
          // LicensedProvider = getCellValue(worksheet, row, ""),
          // Credentials = getCellValue(worksheet, row, ""),
          // CredentialsVerifiedBy = getCellValue(worksheet, row, ""),
          // NonUSGovernmentOfficialStatus = getCellValue(worksheet, row, ""),
          // JobTitle = getCellValue(worksheet, row, ""),
          // PrimaryWorkLocation = getCellValue(worksheet, row, ""),
          // NonUSAddress1 = getCellValue(worksheet, row, ""),
          // NonUSAddress2 = getCellValue(worksheet, row, ""),
          // NonUSAddress3 = getCellValue(worksheet, row, ""),
          // NonUSAddress4 = getCellValue(worksheet, row, ""),
          // NonUSAddress5 = getCellValue(worksheet, row, ""),
          // EAMOnlyFieldGroup1 = getCellValue(worksheet, row, ""),
          // EAMOnlyFieldGroup2 = getCellValue(worksheet, row, ""),
          // EAMOnlyFieldGroup3 = getCellValue(worksheet, row, ""),
          // EAMOnlyFieldGroup4 = getCellValue(worksheet, row, ""),
          // EAMOnlyFieldGroup5 = getCellValue(worksheet, row, ""),
        };

        inputRows.Add(inputRow);
      }

      // group by matchid
      // and filter by:
      // - has a row with stepname "Student Demographic Data" with non-null completed date
      // - has a row with stepname "Coordinator Step" with non-null completed date
      var matchGroups = inputRows.GroupBy(x => x.MatchId);
      var matchGroupsWithStepsCompleted = matchGroups.Where(x =>
           x.Any(y => y.StepName == "Student Demographic Data" && y.CompletedDate != null)
        && x.Any(y => y.StepName == "Coordinator Step" && y.CompletedDate != null)
      );

      Console.WriteLine($"Found {inputRows.Count()} input rows");
      Console.WriteLine($"Found {matchGroups.Count()} distinct matches");
      Console.WriteLine($"Found {matchGroupsWithStepsCompleted.Count()} matches with completed Student Demographic Data and Coordinator Step");

      var outputRows = new List<Dictionary<string, string?>>();

      // generate 1 output row for each matchGroup
      foreach (var matchGroup in matchGroupsWithStepsCompleted)
      {
        var coordinatorStepRow = matchGroup.First(x => x.StepName == "Coordinator Step" && x.CompletedDate != null);
        var studentDemographicDataRow = matchGroup.First(x => x.StepName == "Student Demographic Data" && x.CompletedDate != null);

        // TODO: fill in the missing columns as well as possible
        var outputRow = new Dictionary<string, string?> {
          { "Non-Employee Category", coordinatorStepRow.CoordinatorNonEmployeeCategory },
          { "Type of Non-Employee", coordinatorStepRow.CoordinatorTypeOfNonEmployee },
          { "Legal First Name", studentDemographicDataRow.StudentLegalFirstName },
          { "Legal Last Name", studentDemographicDataRow.StudentLegalLastName },
          { "Gender", studentDemographicDataRow.StudentGender },
          { "Date of Birth", studentDemographicDataRow.StudentDateOfBirth },
          { "Country of Residential Address", studentDemographicDataRow.StudentCountryOfResidence },
          { "Street Address Line 1", studentDemographicDataRow.StudentHomeAddress },
          { "City", studentDemographicDataRow.StudentCity },
          { "Zip or Postal Code", studentDemographicDataRow.StudentZipOrPostalCode },
          { "Country of Citizenship", studentDemographicDataRow.StudentCountryOfCitizenship },
          { "Country of Origin", studentDemographicDataRow.StudentCountryOfOrigin },
          { "Visa Type", studentDemographicDataRow.StudentVisaType },
          { "Current Employer", coordinatorStepRow.CoordinatorWorkState },
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
          { "Middle Name", "" },
          { "Suffix", "" },
          { "Street Address Line 2", "" },
          { "State", "" },
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

      // output as json
      Console.WriteLine($"Output rows:");
      Console.WriteLine(JsonSerializer.Serialize(outputRows, new JsonSerializerOptions
      {
        WriteIndented = true
      }));

      // Generate output file
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
      for (int i = 0; i < outputRows.Count; i++) {
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
}
