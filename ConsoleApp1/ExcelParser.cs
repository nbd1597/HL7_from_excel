using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Globalization;
using OfficeOpenXml;
using NHapi.Model.V23.Message;


namespace ExceltoHl7
{
    internal class ExcelParser
    {
        private ADT_A01 adt01;
        static public String NewFile;
        static public String ExcelPathFile = Directory.GetCurrentDirectory() + "\\excelpath.txt";
        static public StreamReader DefaultExcelPath = new StreamReader(ExcelPathFile);
        static public String ExcelPath = DefaultExcelPath.ReadLine();
        
        public ADT_A01 ReadTemplte()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // enable license for epplus api
            //Console.WriteLine(ExcelPath);
            
            // make new file base on template
            Console.WriteLine("new file name: ");
            NewFile = Console.ReadLine();
            if (!Directory.Exists(String.Format("{0}\\HL7TestOutputs\\excel", Directory.GetCurrentDirectory())))    
                    Directory.CreateDirectory(String.Format("{0}\\HL7TestOutputs\\excel", Directory.GetCurrentDirectory()));
            String filepath = String.Format("{0}\\HL7TestOutputs\\excel\\{1}.xlsx", Directory.GetCurrentDirectory(),NewFile);   // where to save the file
            //var adt01_template = new FileInfo(@"d:\Project\ConsoleApp1\ConsoleApp1\adt01.xlsx");
            String templatepath;
            Console.WriteLine("Use basic form excel ? ");
            Console.WriteLine("'enter' for short, type 'long' for longer form");
            if(Console.ReadLine() == "long")
            {
                templatepath = Directory.GetCurrentDirectory() + "\\adt01.xlsx";     // template path
            }
            else
            {
                templatepath = Directory.GetCurrentDirectory() + "\\adt01_short.xlsx";     // template path
            }
            
            var adt01_template = new FileInfo(templatepath);
            using (var NewExcel = new ExcelPackage(adt01_template))
            {
                NewExcel.SaveAs(new FileInfo(@filepath));       //save new file
            }
            
            try
            {
                System.Diagnostics.Process.Start(ExcelPath, filepath);  // open file with excel
            }
            catch (Exception e)
            {
                Console.WriteLine($"********Error {e.Message}************");
                Console.WriteLine("************Please restart and select correct path*************");
            }
            //System.Diagnostics.Process.Start(ExcelPath,filepath);
            Console.WriteLine("Please close and save file after complete");
            Console.WriteLine("press 'yes' to confirm finished");
            while (Console.ReadLine() != "yes")
            {
                continue;
            }

            var currentDateTimeString = DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            adt01 = new ADT_A01();
            var WorkingFile = new FileInfo(@filepath);
            //CreateMshSegment(currentDateTimeString, WorkingFile);
            if (templatepath == Directory.GetCurrentDirectory() + "\\adt01.xlsx")
            {
                CreateMshSegment(currentDateTimeString, WorkingFile, filepath);
                CreateEvnSegment(currentDateTimeString, WorkingFile, filepath);
                CreatePidSegment(WorkingFile);
                CreateNk1Segment(WorkingFile);
                CreatePv1Segment(WorkingFile);
            }
            else
            {
                CreateShortFormSegment(currentDateTimeString, WorkingFile, filepath);
            }


            return adt01;
        }
        public void CreateShortFormSegment(string currentDateTimeString, FileInfo file, String filepath)
        {
            

            using (var p = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
                adt01.MSH.FieldSeparator.Value = "|";
                adt01.MSH.EncodingCharacters.Value = "^~\\&";
                adt01.MSH.SendingApplication.NamespaceID.Value = "our app";
                adt01.MSH.SendingFacility.NamespaceID.Value = "laptop";
                adt01.MSH.DateTimeOfMessage.TimeOfAnEvent.Value = currentDateTimeString;
                //Console.WriteLine(adt01.MSH.DateTimeOfMessage.TimeOfAnEvent.Value.ToString());
                worksheet.Cells["B2"].Value = currentDateTimeString.ToString();
                adt01.MSH.MessageType.MessageType.Value = "ADT_A01";
                adt01.MSH.VersionID.Value = "v23";

                adt01.EVN.EventTypeCode.Value = "A01";
                adt01.EVN.RecordedDateTime.TimeOfAnEvent.Value = currentDateTimeString;

                adt01.PID.SetIDPatientID.Value = worksheet.Cells["B5"].Value.ToString();
                adt01.PID.GetPatientName(0).FamilyName.Value = worksheet.Cells["B4"].Value.ToString();
                adt01.PID.GetPatientName(0).GivenName.Value = worksheet.Cells["D4"].Value.ToString();
                adt01.PID.GetPatientAddress(0).StreetAddress.Value = worksheet.Cells["B6"].Value.ToString();
                adt01.PID.GetPatientAddress(0).City.Value = worksheet.Cells["D6"].Value.ToString();
                adt01.PID.GetPatientAddress(0).StateOrProvince.Value = worksheet.Cells["B7"].Value.ToString();
                adt01.PID.GetPhoneNumberHome(0).PhoneNumber.Value = worksheet.Cells["D7"].Value.ToString();
                adt01.PID.Sex.Value = worksheet.Cells["B8"].Value.ToString();
                adt01.PID.DateOfBirth.TimeOfAnEvent.Value = worksheet.Cells["D8"].Value.ToString();
                
                adt01.PV1.AssignedPatientLocation.PointOfCare.Value = worksheet.Cells["B12"].Value.ToString();
                adt01.PV1.AssignedPatientLocation.Room.Value = worksheet.Cells["D12"].Value.ToString();
                adt01.PV1.AssignedPatientLocation.Bed.Value = worksheet.Cells["B13"].Value.ToString();
                adt01.PV1.GetAdmittingDoctor(0).IDNumber.Value = worksheet.Cells["B19"].Value.ToString();
                adt01.PV1.GetAdmittingDoctor(0).FamilyName.Value = worksheet.Cells["B18"].Value.ToString();
                adt01.PV1.GetAdmittingDoctor(0).GivenName.Value = worksheet.Cells["D18"].Value.ToString();
                adt01.PV1.AdmitDateTime.TimeOfAnEvent.Value = worksheet.Cells["B14"].Value.ToString();

                adt01.PV2.AdmitReason.Identifier.Value = worksheet.Cells["D14"].Value.ToString();

                adt01.GetINSURANCE(0).IN1.InsurancePlanID.Identifier.Value = worksheet.Cells["B9"].Value.ToString();
                adt01.GetINSURANCE(0).IN1.PlanExpirationDate.Value = worksheet.Cells["D9"].Value.ToString();

                p.SaveAs(new FileInfo(@filepath));
            }
        }
        public void CreateMshSegment(string currentDateTimeString, FileInfo file, String filepath)
        {
            var mshSegment = adt01.MSH;
            using (var p = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
                mshSegment.FieldSeparator.Value = worksheet.Cells["B3"].Value.ToString();
                mshSegment.EncodingCharacters.Value = worksheet.Cells["B4"].Value.ToString();
                mshSegment.SendingApplication.NamespaceID.Value = worksheet.Cells["B5"].Value.ToString();
                mshSegment.SendingFacility.NamespaceID.Value = worksheet.Cells["B6"].Value.ToString();
                mshSegment.ReceivingApplication.NamespaceID.Value = worksheet.Cells["B7"].Value.ToString();
                mshSegment.ReceivingFacility.NamespaceID.Value = worksheet.Cells["B8"].Value.ToString();
                mshSegment.DateTimeOfMessage.TimeOfAnEvent.Value = currentDateTimeString;
                worksheet.Cells["B9"].Value = mshSegment.DateTimeOfMessage.TimeOfAnEvent.Value.ToString();
                mshSegment.Security.Value = worksheet.Cells["B10"].Value.ToString();
                mshSegment.MessageType.MessageType.Value = worksheet.Cells["B12"].Value.ToString();
                mshSegment.MessageType.TriggerEvent.Value = worksheet.Cells["B12"].Value.ToString();
                //mshSegment.MessageType.MessageStructure.Value = worksheet.Cells["B14"].Value.ToString();
                mshSegment.MessageControlID.Value = worksheet.Cells["B14"].Value.ToString();
                mshSegment.ProcessingID.ProcessingID.Value = worksheet.Cells["B15"].Value.ToString();
                mshSegment.VersionID.Value = worksheet.Cells["B16"].Value.ToString();
                mshSegment.SequenceNumber.Value = worksheet.Cells["B17"].Value.ToString();
                mshSegment.ContinuationPointer.Value = worksheet.Cells["B18"].Value.ToString();

                p.SaveAs(new FileInfo(@filepath));

            }

        }
        public void CreateEvnSegment(string currentDateTimeString, FileInfo file, String filepath)
        {
            var evn = adt01.EVN;
            using (var p = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
                
                evn.EventTypeCode.Value = "A01";
                evn.RecordedDateTime.TimeOfAnEvent.Value = currentDateTimeString;
                worksheet.Cells["B23"].Value = evn.RecordedDateTime.TimeOfAnEvent.Value.ToString();
                evn.EventReasonCode.Value = worksheet.Cells["B25"].Value.ToString();
                evn.DateTimePlannedEvent.TimeOfAnEvent.Value = worksheet.Cells["B24"].Value.ToString();

                p.SaveAs(new FileInfo(@filepath));
            }

        }
        public void CreatePidSegment(FileInfo file)
        {
            var pid = adt01.PID;
            using (var p = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
                
                pid.SetIDPatientID.Value = worksheet.Cells["B29"].Value.ToString();
                pid.PatientIDExternalID.ID.Value = worksheet.Cells["B31"].Value.ToString();
                pid.PatientIDExternalID.CheckDigit.Value = worksheet.Cells["B32"].Value.ToString();
                pid.PatientIDExternalID.CodeIdentifyingTheCheckDigitSchemeEmployed.Value = worksheet.Cells["B33"].Value.ToString();
                pid.PatientIDExternalID.AssigningAuthority.NamespaceID.Value = worksheet.Cells["B34"].Value.ToString();
                pid.PatientIDExternalID.IdentifierTypeCode.Value = worksheet.Cells["B35"].Value.ToString();
                pid.PatientIDExternalID.AssigningFacility.NamespaceID.Value = worksheet.Cells["B36"].Value.ToString();

                pid.GetPatientIDInternalID(0).ID.Value = worksheet.Cells["B37"].Value.ToString();
                pid.GetAlternatePatientID(0).ID.Value = worksheet.Cells["B38"].Value.ToString();


                pid.GetPatientName(0).FamilyName.Value = worksheet.Cells["B40"].Value.ToString();
                pid.GetPatientName(0).GivenName.Value = worksheet.Cells["B41"].Value.ToString();
                pid.MotherSMaidenName.FamilyName.Value = worksheet.Cells["B42"].Value.ToString();
                pid.DateOfBirth.TimeOfAnEvent.Value = worksheet.Cells["B43"].Value.ToString();
                pid.Sex.Value = worksheet.Cells["B44"].Value.ToString();
                pid.GetPatientAlias(0).GivenName.Value = worksheet.Cells["B45"].Value.ToString();
                pid.Race.Value = worksheet.Cells["B46"].Value.ToString();
                //var PatientAddress = pid.GetPatientAddress(0);
                pid.GetPatientAddress(0).StreetAddress.Value = worksheet.Cells["B48"].Value.ToString();
                pid.GetPatientAddress(0).OtherDesignation.Value = worksheet.Cells["B49"].Value.ToString();
                pid.GetPatientAddress(0).City.Value = worksheet.Cells["B50"].Value.ToString();
                pid.GetPatientAddress(0).StateOrProvince.Value = worksheet.Cells["B51"].Value.ToString();
                pid.GetPatientAddress(0).ZipOrPostalCode.Value = worksheet.Cells["B52"].Value.ToString();
                //pid.SetIDPID.Value = worksheet.Cells["B30"].Value.ToString();
                pid.CountyCode.Value = worksheet.Cells["B53"].Value.ToString();
                pid.GetPhoneNumberHome(0).PhoneNumber.Value = worksheet.Cells["B54"].Value.ToString();
                pid.GetPhoneNumberBusiness(0).PhoneNumber.Value = worksheet.Cells["B55"].Value.ToString();
                pid.PrimaryLanguage.Identifier.Value = worksheet.Cells["B56"].Value.ToString();
                pid.MaritalStatus.Value = worksheet.Cells["B57"].Value.ToString();
                pid.Religion.Value = worksheet.Cells["B58"].Value.ToString();
                pid.PatientAccountNumber.ID.Value = worksheet.Cells["B60"].Value.ToString();
                pid.PatientAccountNumber.CheckDigit.Value = worksheet.Cells["B61"].Value.ToString();
                pid.PatientAccountNumber.CodeIdentifyingTheCheckDigitSchemeEmployed.Value = worksheet.Cells["B62"].Value.ToString();
                pid.PatientAccountNumber.AssigningAuthority.NamespaceID.Value = worksheet.Cells["B63"].Value.ToString();
                pid.PatientAccountNumber.IdentifierTypeCode.Value = worksheet.Cells["B64"].Value.ToString();
                pid.PatientAccountNumber.AssigningFacility.NamespaceID.Value = worksheet.Cells["B65"].Value.ToString();
                pid.SSNNumberPatient.Value = worksheet.Cells["B66"].Value.ToString();
                pid.DriverSLicenseNumber.DriverSLicenseNumber.Value = worksheet.Cells["B67"].Value.ToString();
                pid.GetMotherSIdentifier(0).ID.Value = worksheet.Cells["B68"].Value.ToString();
            }
        }
        public void CreateNk1Segment(FileInfo file)
        {
            //adt01.AddNK1();
            var nk1 = adt01.GetNK1(0);
            using (var p = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
                nk1.SetIDNextOfKin.Value = worksheet.Cells["B72"].Value.ToString();
                nk1.GetName(0).FamilyName.Value = worksheet.Cells["B74"].Value.ToString();
                nk1.GetName(0).GivenName.Value = worksheet.Cells["B75"].Value.ToString();
                nk1.GetName(0).MiddleInitialOrName.Value = worksheet.Cells["B76"].Value.ToString();
                nk1.Relationship.Identifier.Value = worksheet.Cells["B77"].Value.ToString();
                nk1.GetAddress(0).StreetAddress.Value = worksheet.Cells["B79"].Value.ToString();
                nk1.GetAddress(0).OtherDesignation.Value = worksheet.Cells["B80"].Value.ToString();
                nk1.GetAddress(0).City.Value = worksheet.Cells["B81"].Value.ToString();
                nk1.GetAddress(0).StateOrProvince.Value = worksheet.Cells["B82"].Value.ToString();
                nk1.GetAddress(0).ZipOrPostalCode.Value = worksheet.Cells["B83"].Value.ToString();
                nk1.GetPhoneNumber(0).PhoneNumber.Value = worksheet.Cells["B84"].Value.ToString();
                nk1.GetBusinessPhoneNumber(0).PhoneNumber.Value = worksheet.Cells["B85"].Value.ToString();
                nk1.ContactRole.Identifier.Value = worksheet.Cells["B86"].Value.ToString();
            }
        }
        public void CreatePv1Segment(FileInfo file)
        {
            var pv1 = adt01.PV1;
            using (var p = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = p.Workbook.Worksheets[0];
                pv1.SetIDPatientVisit.Value = worksheet.Cells["B90"].Value.ToString();
                pv1.PatientClass.Value = worksheet.Cells["B91"].Value.ToString();
                pv1.AssignedPatientLocation.PointOfCare.Value = worksheet.Cells["B93"].Value.ToString();
                pv1.AssignedPatientLocation.Room.Value = worksheet.Cells["B94"].Value.ToString();
                pv1.AssignedPatientLocation.Bed.Value = worksheet.Cells["B95"].Value.ToString();
                pv1.AdmissionType.Value = worksheet.Cells["B96"].Value.ToString();
                pv1.PreadmitNumber.ID.Value = worksheet.Cells["B97"].Value.ToString();
                pv1.PriorPatientLocation.PointOfCare.Value = worksheet.Cells["B98"].Value.ToString();
                pv1.GetAttendingDoctor(0).IDNumber.Value = worksheet.Cells["B100"].Value.ToString();
                pv1.GetAttendingDoctor(0).FamilyName.Value = worksheet.Cells["B101"].Value.ToString();
                pv1.GetAttendingDoctor(0).GivenName.Value = worksheet.Cells["B102"].Value.ToString();
                pv1.GetAttendingDoctor(0).MiddleInitialOrName.Value = worksheet.Cells["B103"].Value.ToString();
                pv1.GetReferringDoctor(0).IDNumber.Value = worksheet.Cells["B105"].Value.ToString();
                pv1.GetReferringDoctor(0).FamilyName.Value = worksheet.Cells["B106"].Value.ToString();
                pv1.GetReferringDoctor(0).GivenName.Value = worksheet.Cells["B107"].Value.ToString();
                pv1.GetReferringDoctor(0).MiddleInitialOrName.Value = worksheet.Cells["B108"].Value.ToString();
                pv1.GetConsultingDoctor(0).IDNumber.Value = worksheet.Cells["B110"].Value.ToString();
                pv1.GetConsultingDoctor(0).FamilyName.Value = worksheet.Cells["B111"].Value.ToString();
                pv1.GetConsultingDoctor(0).GivenName.Value = worksheet.Cells["B112"].Value.ToString();
                pv1.GetConsultingDoctor(0).MiddleInitialOrName.Value = worksheet.Cells["B113"].Value.ToString();
                pv1.HospitalService.Value = worksheet.Cells["B114"].Value.ToString();
                pv1.TemporaryLocation.PointOfCare.Value = worksheet.Cells["B115"].Value.ToString();
                pv1.PreadmitTestIndicator.Value = worksheet.Cells["B116"].Value.ToString();
                pv1.ReadmissionIndicator.Value = worksheet.Cells["B117"].Value.ToString();
                pv1.AdmitSource.Value = worksheet.Cells["B118"].Value.ToString();
                pv1.GetAmbulatoryStatus(0).Value = worksheet.Cells["B119"].Value.ToString();
                pv1.VIPIndicator.Value = worksheet.Cells["B120"].Value.ToString();

            }
        }


    }
}
