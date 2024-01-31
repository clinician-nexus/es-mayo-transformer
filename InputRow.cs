namespace MayoExcelTransformer;

public class InputRow
{
  public string? StepName { get; set; }
  public string? MatchId { get; set; }
  public string? UserName { get; set; }
  public string? UserEmail { get; set; }
  public string? TargetUserName { get; set; }
  public string? TargetUserEmail { get; set; }
  public string? CompletedUserName { get; set; }
  public string? CompletedUserEmail { get; set; }
  public string? CompletedDate { get; set; }
  public string? WaivedUserName { get; set; }
  public string? WaivedUserEmail { get; set; }
  public string? WaivedDate { get; set; }
  public string? SkippedUserName { get; set; }
  public string? SkippedUserEmail { get; set; }
  public string? SkippedDate { get; set; }
  public string? VerifiedUserName { get; set; }
  public string? VerifiedUserEmail { get; set; }
  public string? VerifiedDate { get; set; }
  public string? VerificationExpirationDate { get; set; }
  public string? VerificationExpirationDateSuggestion { get; set; }
  public string? RejectedReason { get; set; }
  public string? SchoolRelativeName { get; set; }
  public string? CompleteFromDate { get; set; }
  public string? CompleteToDate { get; set; }
  public string? MatchGuestOrgRelativeName { get; set; }
  public string? MatchGuestOrgName { get; set; }
  public string? MatchHostOrgRelativeName { get; set; }
  public string? MatchHostOrgName { get; set; }
  public string? MatchCapacityName { get; set; }
  public string? MatchServiceName { get; set; }
  public string? MatchStatus { get; set; }
  public string? MatchStartDate { get; set; }
  public string? MatchEndDate { get; set; }
  public string? Files { get; set; }

  // form Electronic Authentication Security Agreement Statement 1) Name/Signature
  public string? ElectronicAuthAgreementNameSignature { get; set; }

  // form Student Demographic Data 1) Legal First Name
  public string? StudentLegalFirstName { get; set; }
  public string? StudentLegalLastName { get; set; }
  public string? StudentPreferredName { get; set; }
  public string? StudentMiddleName { get; set; }
  public string? StudentSuffix { get; set; }
  public string? StudentGender { get; set; }
  public string? StudentDateOfBirth { get; set; }
  public string? StudentHomeAddressLine1 { get; set; }
  public string? StudentHomeAddressLine2 { get; set; }
  public string? StudentCity { get; set; }
  public string? StudentState { get; set; }
  public string? StudentCountryOfResidence { get; set; }
  public string? StudentZipOrPostalCode { get; set; }
  public string? StudentSchoolEmailAddress { get; set; }
  public string? StudentPrimaryPhoneNumber { get; set; }
  public string? StudentMayoClinicEmployeeStatus { get; set; }
  public string? StudentPreviousRotation { get; set; }
  public string? StudentLANID { get; set; }
  public string? StudentPreviousRotationDetails { get; set; }
  public string? StudentCountryOfOrigin { get; set; }
  public string? StudentCitizenshipStatus { get; set; }
  public string? StudentCountryOfCitizenship { get; set; }
  public string? StudentValidVisaStatus { get; set; }
  public string? StudentVisaType { get; set; }
  public string? StudentGenderInclusivityCommitment { get; set; }
  public string? StudentEthnicity { get; set; }
  public string? StudentCrimeConvictionStatus { get; set; }

  // Additional properties for flu reporting form fields...
  public string? FluVaccineStatus { get; set; }
  public string? FluVaccineDateReceived { get; set; }
  public string? FluVaccineFacilityObtained { get; set; }

  // Additional properties for emergency contact form fields...
  public string? EmergencyContactFirstName { get; set; }
  public string? EmergencyContactLastName { get; set; }
  public string? EmergencyContactPhoneNumber { get; set; }

  // Additional properties for coordinator step form fields...
  public string? CoordinatorNonEmployeeCategory { get; set; }
  public string? CoordinatorCampus { get; set; }
  public string? CoordinatorTypeOfNonEmployee { get; set; }
  public string? CoordinatorWorkState { get; set; }
  public string? CoordinatorLengthOfStay { get; set; }
  public string? CoordinatorSupervisorCode { get; set; }
  public string? CoordinatorSupervisorPerID { get; set; }

  // Additional properties for student demographic data form fields...
  // Include other properties as needed, following the pattern above.
}
