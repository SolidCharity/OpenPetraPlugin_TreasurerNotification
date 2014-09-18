//
// DO NOT REMOVE COPYRIGHT NOTICES OR THIS FILE HEADER.
//
// @Authors:
//       timop
//
// Copyright 2004-2014 by OM International
//
// This file is part of OpenPetra.org.
//
// OpenPetra.org is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// OpenPetra.org is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with OpenPetra.org.  If not, see <http://www.gnu.org/licenses/>.
//
using System;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net.Mail;
using Ict.Common.DB;
using Ict.Common;
using Ict.Common.IO;
using Ict.Common.Data;
using Ict.Common.Printing;
using Ict.Common.Remoting.Server;
using Ict.Petra.Server.App.Core;
using Ict.Petra.Server.App.Core.Security;
using Ict.Petra.Server.MPartner.Common;
using Ict.Petra.Shared.MPartner;
using Ict.Petra.Shared.MPartner.Partner.Data;
using Ict.Petra.Server.MPartner.Partner.Data.Access;
using Ict.Petra.Shared.MCommon.Data;
using Ict.Petra.Server.MCommon.Data.Access;
using Ict.Petra.Shared.MFinance.Account.Data;
using Ict.Petra.Server.MFinance.Account.Data.Access;
using Ict.Petra.Plugins.TreasurerNotification.Data;
using Ict.Petra.Shared.MPersonnel.Personnel.Data;

namespace  Ict.Petra.Plugins.TreasurerNotification.Server.WebConnectors
{
/// <summary>
/// web connector for calculating the emails and letters for the treasurers
/// </summary>
    public class TTreasurerNotificationWebConnector
    {
        /// retrieve all sums of donations for each treasurer for the last x months
        private static TreasurerNotificationTDS GetTreasurerData(
            Int32 ALedgerNumber,
            string AMotivationGroup,
            string AMotivationDetail,
            DateTime ALastDonationDate, Int16 ANumberMonths)
        {
            // establish connection to database
            TAppSettingsManager settings = new TAppSettingsManager();

            // calculate the first and the last days of the months range
            // last donation date covers the full month
            DateTime EndDate =
                new DateTime(ALastDonationDate.Year, ALastDonationDate.Month, DateTime.DaysInMonth(ALastDonationDate.Year, ALastDonationDate.Month));

            DateTime StartDate = EndDate.AddMonths(-1 * (ANumberMonths - 1));

            StartDate = new DateTime(StartDate.Year, StartDate.Month, 1);

            TreasurerNotificationTDS ResultDataset = new TreasurerNotificationTDS();
            GetAllGiftsForRecipientPerMonthByMotivation(ref ResultDataset, ALedgerNumber, AMotivationGroup, AMotivationDetail, StartDate, EndDate);

            // add the last date of the month to the table
            AddMonthDate(ref ResultDataset, ALedgerNumber);

            // get the treasurer(s) for each recipient; get their name and partner key
            AddTreasurer(ref ResultDataset, EndDate, ALedgerNumber);

            // get the name of each recipient
            AddRecipientName(ref ResultDataset);

            // get the country code of the ledger, which is important for printing the address labels
            AddLedgerDetails(ref ResultDataset, ALedgerNumber);

            // use GetBestAddress to get the email address of the treasurer
            AddTreasurerEmailOrPostalAddress(ref ResultDataset);

            return ResultDataset;
        }

        /// <summary>
        /// Get the sum of all gifts per recipient per month, by specified motivation and time span
        /// </summary>
        /// <param name="AMainDS"></param>
        /// <param name="ALedgerNumber"></param>
        /// <param name="AMotivationGroup"></param>
        /// <param name="AMotivationDetail"></param>
        /// <param name="AStartDate"></param>
        /// <param name="AEndDate"></param>
        private static void GetAllGiftsForRecipientPerMonthByMotivation(
            ref TreasurerNotificationTDS AMainDS,
            Int32 ALedgerNumber,
            string AMotivationGroup,
            string AMotivationDetail,
            DateTime AStartDate,
            DateTime AEndDate)
        {
            TDBTransaction transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadUncommitted);

            string stmt = TDataBase.ReadSqlFile("TreasurerGetAllGiftsForRecipientPerMonthByMotivation.sql");

            OdbcParameter[] parameters = new OdbcParameter[6];
            parameters[0] = new OdbcParameter("Ledger", OdbcType.Int);
            parameters[0].Value = ALedgerNumber;
            parameters[1] = new OdbcParameter("MotivationGroup", OdbcType.VarChar);
            parameters[1].Value = AMotivationGroup;
            parameters[2] = new OdbcParameter("MotivationDetail", OdbcType.VarChar);
            parameters[2].Value = AMotivationDetail;
            parameters[3] = new OdbcParameter("StartDate", OdbcType.Date);
            parameters[3].Value = AStartDate;
            parameters[4] = new OdbcParameter("EndDate", OdbcType.Date);
            parameters[4].Value = AEndDate;
            parameters[5] = new OdbcParameter("HomeOffice", OdbcType.BigInt);
            parameters[5].Value = ALedgerNumber * 1000000L;
            DBAccess.GDBAccessObj.Select(AMainDS, stmt, AMainDS.DonationsPerMonth.TableName, transaction, parameters);
            DBAccess.GDBAccessObj.RollbackTransaction();
        }

        /// <summary>
        /// The previous query only retrieves the financial year number and period number,
        /// but we want the last date of that period to be added to the result table
        /// </summary>
        /// <param name="AMainDS"></param>
        /// <param name="ALedgerNumber"></param>
        private static void AddMonthDate(ref TreasurerNotificationTDS AMainDS, Int32 ALedgerNumber)
        {
            TDBTransaction transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadUncommitted);

            // get the accounting period table to know the month
            AAccountingPeriodTable periods = AAccountingPeriodAccess.LoadAll(transaction);

            OdbcParameter[] parameters = new OdbcParameter[1];
            parameters[0] = new OdbcParameter("Ledger", ALedgerNumber);
            Int32 currentFinancialYear =
                Convert.ToInt32(DBAccess.GDBAccessObj.ExecuteScalar("SELECT a_current_financial_year_i FROM PUB_a_ledger WHERE a_ledger_number_i = ?",
                        transaction, parameters));

            foreach (TreasurerNotificationTDSDonationsPerMonthRow row in AMainDS.DonationsPerMonth.Rows)
            {
                // somehow in Petra 2.x there is a problem that some batches have period and batch = 0, but the gl date effective is correct
                if (row.FinancialPeriod == 0)
                {
                    throw new Exception("Problem with gift batch, no valid period; please run a fix program on the Petra database");
                }

                DateTime monthDate = DateTime.MinValue;

                foreach (AAccountingPeriodRow period in periods.Rows)
                {
                    if ((period.LedgerNumber == ALedgerNumber) && (period.AccountingPeriodNumber == row.FinancialPeriod))
                    {
                        monthDate = period.PeriodEndDate;
                    }
                }

                if (row.FinancialYear != currentFinancialYear)
                {
                    // substract the years to get the right date
                    monthDate = monthDate.AddYears(-1 * (currentFinancialYear - row.FinancialYear));
                }

                row.MonthDate = monthDate;
            }

            DBAccess.GDBAccessObj.RollbackTransaction();
        }

        /// <summary>
        /// get the name of the recipient and add to the table
        /// </summary>
        /// <param name="AMainDS"></param>
        private static void AddRecipientName(ref TreasurerNotificationTDS AMainDS)
        {
            TDBTransaction transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadUncommitted);

            foreach (TreasurerNotificationTDSTreasurerRow row in AMainDS.Treasurer.Rows)
            {
                row.RecipientName = GetPartnerShortName(row.RecipientKey, transaction);
            }

            DBAccess.GDBAccessObj.RollbackTransaction();
        }

        private static void AddLedgerDetails(ref TreasurerNotificationTDS AMainDS, Int32 ALedgerNumber)
        {
            TDBTransaction transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadUncommitted);

            AMainDS.ALedger.Merge(ALedgerAccess.LoadByPrimaryKey(ALedgerNumber, transaction));

            DBAccess.GDBAccessObj.RollbackTransaction();
        }

        /// <summary>
        /// get the treasurer(s) for each recipient;
        /// get their name and partner key
        /// </summary>
        private static void AddTreasurer(ref TreasurerNotificationTDS AMainDS, DateTime AEndDate, Int32 ALedgerNumber)
        {
            TDBTransaction transaction = DBAccess.GDBAccessObj.BeginTransaction(IsolationLevel.ReadUncommitted);

            // to check if the recipient has been processed already
            List <Int64>recipients = new List <Int64>();

            foreach (TreasurerNotificationTDSDonationsPerMonthRow row in AMainDS.DonationsPerMonth.Rows)
            {
                Int64 recipientKey = row.RecipientKey;

                if (recipients.Contains(recipientKey))
                {
                    continue;
                }

                recipients.Add(recipientKey);

                // first check if the worker has a valid commitment period, or is in TRANSITION
                bool bTransition = false;
                bool bExWorker = false;
                bool bFutureWorker = false;
                bool bNonNationalWorker = false;
                bool bCurrentCommitment = false;

                // commitments of worker by family key
                OdbcParameter[] parameters = new OdbcParameter[1];
                parameters[0] = new OdbcParameter("PartnerKey", recipientKey);

                string stmt = TDataBase.ReadSqlFile("TreasurerCommitmentsOfWorker.sql");

                PmStaffDataTable CommitmentsTable = new PmStaffDataTable();
                DBAccess.GDBAccessObj.SelectDT(CommitmentsTable, stmt,
                    transaction,
                    parameters, 0, 0);

                foreach (PmStaffDataRow commitment in CommitmentsTable.Rows)
                {
                    if ((commitment.IsEndOfCommitmentNull() && (AEndDate >= commitment.StartOfCommitment))
                        || (AEndDate >= commitment.StartOfCommitment) && (AEndDate <= commitment.EndOfCommitment))
                    {
                        // check if our office is the home office
                        if (commitment.HomeOffice == ALedgerNumber * 1000000L)
                        {
                            // currently valid commitment
                            bCurrentCommitment = true;

                            if (commitment.StatusCode == "TRANSITION")
                            {
                                bTransition = true;
                            }
                        }
                        else
                        {
                            bNonNationalWorker = true;
                        }
                    }
                    else if (commitment.StartOfCommitment > AEndDate)
                    {
                        // check if our office is the home office
                        if (commitment.HomeOffice == ALedgerNumber * 1000000L)
                        {
                            // will join soon
                            bFutureWorker = true;
                        }
                        else
                        {
                            bNonNationalWorker = true;
                        }
                    }
                    else if (!commitment.IsEndOfCommitmentNull() && (AEndDate > commitment.EndOfCommitment))
                    {
                        // check if our office is the home office
                        if (commitment.HomeOffice == ALedgerNumber * 1000000L)
                        {
                            // has already left
                            bExWorker = true;
                        }
                        else
                        {
                            bNonNationalWorker = true;
                        }
                    }
                }

                // if there is a current commitment, that overrides all other commitments
                if (bCurrentCommitment)
                {
                    bNonNationalWorker = false;
                    bExWorker = false;
                    bFutureWorker = false;
                }
                else if (bFutureWorker)
                {
                    bExWorker = false;
                    bNonNationalWorker = false;
                }
                else if (bNonNationalWorker)
                {
                    bExWorker = false;
                }

                parameters = new OdbcParameter[1];
                parameters[0] = new OdbcParameter("PartnerKey", recipientKey);

                stmt = TDataBase.ReadSqlFile("TreasurerOfWorker.sql");

                TreasurerNotificationTDSTreasurerTable TreasurerTable =
                    new TreasurerNotificationTDSTreasurerTable();

                DBAccess.GDBAccessObj.SelectDT(TreasurerTable, stmt, transaction, parameters, 0, 0);

                if (TreasurerTable.Rows.Count >= 1)
                {
                    foreach (TreasurerNotificationTDSTreasurerRow r in TreasurerTable.Rows)
                    {
                        r.Transition = bTransition;

                        if (bNonNationalWorker)
                        {
                            r.ErrorMessage = "NONNATIONALWORKER";
                        }

                        if (bExWorker)
                        {
                            r.ErrorMessage = "EXWORKER";
                        }

                        r.TreasurerName = GetPartnerShortName(r.TreasurerKey, transaction);
                    }

                    AMainDS.Treasurer.Merge(TreasurerTable);
                }
                else
                {
                    TreasurerNotificationTDSTreasurerRow InvalidTreasurer = AMainDS.Treasurer.NewRowTyped();
                    InvalidTreasurer.RecipientKey = row.RecipientKey;
                    InvalidTreasurer.Transition = bTransition;

                    if (bNonNationalWorker)
                    {
                        InvalidTreasurer.ErrorMessage = "NONNATIONALWORKER";
                    }

                    if (bExWorker)
                    {
                        InvalidTreasurer.ErrorMessage = "EXWORKER";
                    }

                    AMainDS.Treasurer.Rows.Add(InvalidTreasurer);
                }
            }

            DBAccess.GDBAccessObj.RollbackTransaction();
        }

        /// <summary>
        /// get the email address or the postal address of the treasurer and add to the table
        /// </summary>
        /// <param name="AMainDS"></param>
        private static void AddTreasurerEmailOrPostalAddress(ref TreasurerNotificationTDS AMainDS)
        {
            foreach (TreasurerNotificationTDSTreasurerRow row in AMainDS.Treasurer.Rows)
            {
                if (!row.IsTreasurerKeyNull())
                {
                    PLocationTable Address;
                    string CountryNameLocal;
                    string emailAddress = TMailing.GetBestEmailAddressWithDetails(row.TreasurerKey, out Address, out CountryNameLocal);

                    if (emailAddress.Length > 0)
                    {
                        row.TreasurerEmail = emailAddress;
                    }

                    row.ValidAddress = (Address != null);

                    if (Address == null)
                    {
                        // no best address; only report if emailAddress is empty as well???
                        continue;
                    }

                    if (!Address[0].IsLocalityNull())
                    {
                        row.TreasurerLocality = Address[0].Locality;
                    }

                    if (!Address[0].IsStreetNameNull())
                    {
                        row.TreasurerStreetName = Address[0].StreetName;
                    }

                    if (!Address[0].IsBuilding1Null())
                    {
                        row.TreasurerBuilding1 = Address[0].Building1;
                    }

                    if (!Address[0].IsBuilding2Null())
                    {
                        row.TreasurerBuilding2 = Address[0].Building2;
                    }

                    if (!Address[0].IsAddress3Null())
                    {
                        row.TreasurerAddress3 = Address[0].Address3;
                    }

                    if (!Address[0].IsCountryCodeNull())
                    {
                        row.TreasurerCountryCode = Address[0].CountryCode;
                    }

                    row.TreasurerCountryName = CountryNameLocal;

                    if (!Address[0].IsPostalCodeNull())
                    {
                        row.TreasurerPostalCode = Address[0].PostalCode;
                    }

                    if (!Address[0].IsCityNull())
                    {
                        row.TreasurerCity = Address[0].City;
                    }
                }
            }
        }

        /// <summary>
        /// generate messages, both emails and letters
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static TreasurerNotificationTDSMessageTable GenerateMessages(
            Int32 ALedgerNumber,
            string AHTMLTemplate,
            string AMotivationGroupCode,
            string AMotivationDetailCode,
            bool AForceLetters,
            DateTime ALastDonationDate,
            Int16 ANumberMonths)
        {
            string MyClientID = DomainManager.GClientID.ToString();

            TProgressTracker.InitProgressTracker(MyClientID,
                Catalog.GetString("Generating Letters and Emails"),
                4);

            TProgressTracker.SetCurrentState(MyClientID, "Processing Treasurers...", 1);

            TreasurerNotificationTDS MainDS = GetTreasurerData(
                ALedgerNumber,
                AMotivationGroupCode,
                AMotivationDetailCode,
                ALastDonationDate,
                ANumberMonths);

            TreasurerNotificationTDSMessageTable messages = new TreasurerNotificationTDSMessageTable();

            DataView view = MainDS.Treasurer.DefaultView;

            view.Sort = "TreasurerName ASC";

            string LedgerCountryCode = ((ALedgerTable)MainDS.Tables[ALedgerTable.GetTableName()])[0].CountryCode;

            Int64 PreviousTreasurerKey = -1;

            TProgressTracker.SetCurrentState(MyClientID, "Processing Donations...", 3);

            try
            {
                foreach (DataRowView rowview in view)
                {
                    TreasurerNotificationTDSTreasurerRow row = (TreasurerNotificationTDSTreasurerRow)rowview.Row;

                    string treasurerName = "";
                    Int64 treasurerKey = -1;
                    string errorMessage = "NOTREASURER";
                    bool SeveralLettersForSameTreasurer = false;

                    if (!row.IsTreasurerKeyNull())
                    {
                        treasurerName = row.TreasurerName;
                        treasurerKey = row.TreasurerKey;
                        SeveralLettersForSameTreasurer = (treasurerKey == PreviousTreasurerKey);
                        PreviousTreasurerKey = treasurerKey;
                        errorMessage = String.Empty;
                    }

                    if (!row.IsErrorMessageNull())
                    {
                        errorMessage = row.ErrorMessage;

                        bool bRecentGift = false;

                        // check if there has been a gift in the last month of the reporting period
                        DataRow[] rowGifts = MainDS.DonationsPerMonth.Select("RecipientKey = " + row.RecipientKey.ToString(), "MonthDate");

                        if (rowGifts.Length > 0)
                        {
                            DateTime month = Convert.ToDateTime(rowGifts[rowGifts.Length - 1]["MonthDate"]);
                            bRecentGift = (month.Month == ALastDonationDate.Month);
                        }

                        if (!bRecentGift)
                        {
                            continue;
                        }
                    }

                    TreasurerNotificationTDSMessageRow letter = messages.NewRowTyped();
                    letter.Subject = String.Format(Catalog.GetString("Gifts for {0}"), row.RecipientName);
                    letter.SimpleMessageText = GenerateSimpleDebugString(MainDS, row);
                    letter.TreasurerName = treasurerName;
                    letter.TreasurerKey = treasurerKey;
                    letter.RecipientName = row.RecipientName;
                    letter.RecipientKey = row.RecipientKey;

                    if (AForceLetters
                        || row.IsTreasurerEmailNull()
                        || row.IsTreasurerKeyNull())
                    {
                        if (!row.IsTreasurerKeyNull()
                            && row.IsTreasurerCityNull()
                            && (errorMessage.Length == 0))
                        {
                            errorMessage = "NOADDRESS";
                        }

                        string subject = letter.Subject;
                        letter.HTMLMessage = GenerateLetterText(MainDS,
                            row,
                            AHTMLTemplate,
                            LedgerCountryCode,
                            "letter",
                            SeveralLettersForSameTreasurer,
                            out subject);
                        letter.Subject = subject;
                    }
                    else
                    {
                        string subject = letter.Subject;
                        letter.HTMLMessage =
                            GenerateLetterText(MainDS, row, AHTMLTemplate, "", "email", false, out subject);
                        letter.Subject = subject;
                        letter.EmailAddress = row.TreasurerEmail;
                    }

                    letter.ErrorMessage = errorMessage;
                    letter.Transition = row.Transition;
                    messages.Rows.Add(letter);
                }

                TProgressTracker.FinishJob(MyClientID);
            }
            catch (Exception ex)
            {
                TLogging.Log("Problem during generation of messages for treasurers: ");
                TLogging.Log(ex.ToString());
                TProgressTracker.FinishJob(MyClientID);
                return new TreasurerNotificationTDSMessageTable();
            }

            return messages;
        }

        /// return the short name for a partner;
        /// the short name is a comma separated list of title, familyname, firstname
        /// p_partner.p_partner_short_name_c is not always useful and reliable (too long names have been cut off in old databases? B...mmer)
        /// TODO: this should not be necessary, clean up your p_partner.p_partner_short_name_c!!!
        private static string GetPartnerShortName(Int64 APartnerKey, TDBTransaction ATransaction)
        {
            OdbcParameter[] parameters = new OdbcParameter[1];
            parameters[0] = new OdbcParameter("PartnerKey", APartnerKey);
            string shortname = DBAccess.GDBAccessObj.ExecuteScalar(
                "SELECT p_partner_short_name_c FROM PUB_p_partner WHERE p_partner_key_n = ?",
                ATransaction, parameters).ToString();

            // p_partner.p_partner_short_name_c is not always useful and reliable (too long names have been cut off in old databases? B...mmer)
            parameters = new OdbcParameter[4];
            parameters[0] = new OdbcParameter("PartnerKey", APartnerKey);
            parameters[1] = new OdbcParameter("PartnerKey", APartnerKey);
            parameters[2] = new OdbcParameter("PartnerKey", APartnerKey);
            parameters[3] = new OdbcParameter("PartnerKey", APartnerKey);

            // TODO: deal with different family names etc
            DataTable NameTable = DBAccess.GDBAccessObj.SelectDT(
                "SELECT p_title_c, p_first_name_c, p_family_name_c FROM PUB_p_person WHERE p_partner_key_n = ? " +
                "UNION SELECT p_title_c, p_first_name_c, p_family_name_c FROM PUB_p_family WHERE p_partner_key_n = ? " +
                "UNION SELECT '', '', p_organisation_name_c FROM PUB_p_organisation WHERE p_partner_key_n = ? " +
                "UNION SELECT '', '', p_church_name_c FROM PUB_p_church WHERE p_partner_key_n = ?",
                "names",
                ATransaction, parameters);

            if (NameTable.Rows.Count > 0)
            {
                shortname = NameTable.Rows[0][2].ToString() + ", " +
                            NameTable.Rows[0][1].ToString() + ", " +
                            NameTable.Rows[0][0].ToString();
            }

            return shortname;
        }

        private static string GetStringOrEmpty(object obj)
        {
            if (obj == System.DBNull.Value)
            {
                return "";
            }

            return obj.ToString();
        }

        /// <summary>
        /// generate the printed letter for one treasurer, one worker
        /// </summary>
        private static string GenerateLetterText(TreasurerNotificationTDS AMainDS,
            TreasurerNotificationTDSTreasurerRow row,
            string AHTMLTemplate,
            string ALedgerCountryCode,
            string ALetterOrEmail,
            bool APreviousLetterSameTreasurer,
            out string ASubject)
        {
            // make sure that the Euro character is printed correctly
            Catalog.Init("de-DE", "de-DE");

            string treasurerName = row.TreasurerName;
            Int64 recipientKey = row.RecipientKey;

            string msg = AHTMLTemplate;

            msg = msg.Replace("#MARKMULTIPLELETTERS", APreviousLetterSameTreasurer ? "*" : "");
            msg = msg.Replace("#RECIPIENTNAME", Calculations.FormatShortName(row.RecipientName.ToString(), eShortNameFormat.eReverseWithoutTitle));
            msg =
                msg.Replace("#RECIPIENTINITIALS",
                    Calculations.FormatShortName(row.RecipientName.ToString(), eShortNameFormat.eReverseLastnameInitialsOnly));
            msg = msg.Replace("#RECIPIENTKEY", row.RecipientKey.ToString());
            msg = msg.Replace("#TREASUREREMAIL", GetStringOrEmpty(row["TreasurerEmail"]));
            msg = msg.Replace("#TREASURERTITLE", Calculations.FormatShortName(treasurerName, eShortNameFormat.eOnlyTitle));
            msg = msg.Replace("#TREASURERNAME", Calculations.FormatShortName(treasurerName, eShortNameFormat.eReverseWithoutTitle));
            msg = msg.Replace("#STREETNAME", GetStringOrEmpty(row["TreasurerStreetName"]));
            msg = msg.Replace("#LOCATION", GetStringOrEmpty(row["TreasurerLocality"]));
            msg = msg.Replace("#ADDRESS3", GetStringOrEmpty(row["TreasurerAddress3"]));
            msg = msg.Replace("#BUILDING1", GetStringOrEmpty(row["TreasurerBuilding1"]));
            msg = msg.Replace("#BUILDING2", GetStringOrEmpty(row["TreasurerBuilding2"]));
            msg = msg.Replace("#CITY", GetStringOrEmpty(row["TreasurerCity"]));
            msg = msg.Replace("#POSTALCODE", GetStringOrEmpty(row["TreasurerPostalCode"]));
            msg = msg.Replace("#DATE", DateTime.Now.ToString("d. MMMM yyyy"));

            // according to German Post, there is no country code in front of the post code
            // if country code is same for the address of the recipient and this office, then COUNTRYNAME is cleared
            if (GetStringOrEmpty(row["TreasurerCountryCode"]) != ALedgerCountryCode)
            {
                msg = msg.Replace("#COUNTRYNAME", GetStringOrEmpty(row["TreasurerCountryName"]));
            }
            else
            {
                msg = msg.Replace("#COUNTRYNAME", "");
            }

            bool bTransition = row.Transition;

            if (bTransition)
            {
                msg = TPrinterHtml.RemoveDivWithClass(msg, "normal");
            }
            else
            {
                msg = TPrinterHtml.RemoveDivWithClass(msg, "transition");
            }

            if (ALetterOrEmail == "letter")
            {
                msg = TPrinterHtml.RemoveDivWithClass(msg, "email");
            }
            else
            {
                msg = TPrinterHtml.RemoveDivWithClass(msg, "letter");
            }

            // recognise detail lines automatically
            string RowTemplate;
            msg = TPrinterHtml.GetTableRow(msg, "#MONTH", out RowTemplate);
            string rowTexts = "";
            DataRow[] rows = AMainDS.DonationsPerMonth.Select("RecipientKey = " + recipientKey.ToString(), "MonthDate");

            foreach (DataRow rowGifts in rows)
            {
                DateTime month = Convert.ToDateTime(rowGifts["MonthDate"]);

                rowTexts += RowTemplate.
                            Replace("#MONTH", month.ToString("MMMM yyyy")).
                            Replace("#AMOUNT", String.Format("{0:C}", Convert.ToDouble(rowGifts["MonthAmount"]))).
                            Replace("#NUMBERGIFTS", rowGifts["MonthCount"].ToString());
            }

            // subject comes from HTML title tag
            ASubject = TPrinterHtml.GetTitle(msg);

            return msg.Replace("#ROWTEMPLATE", rowTexts);
        }

        /// <summary>
        /// generates a simple string with the list of donations per month, for one treasuer and one worker;
        /// this is useful for debugging, and looking at data of exworkers etc
        /// </summary>
        /// <param name="AMainDS"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static string GenerateSimpleDebugString(TreasurerNotificationTDS AMainDS, TreasurerNotificationTDSTreasurerRow row)
        {
            string result = String.Empty;

            result += "Treasurer: " + row.TreasurerName + Environment.NewLine;

            if (row.IsTreasurerKeyNull())
            {
                result += "Treasurer Key: None" + Environment.NewLine;
            }
            else
            {
                result += "Treasurer Key: " + row.TreasurerKey.ToString() + Environment.NewLine;
            }

            result += "Recipient: " + row.RecipientName + Environment.NewLine;
            result += "Recipient Key: " + row.RecipientKey.ToString() + Environment.NewLine;
            result += Environment.NewLine;

            DataRow[] rows = AMainDS.DonationsPerMonth.Select("RecipientKey = " + row.RecipientKey.ToString(), "MonthDate");

            foreach (DataRow rowGifts in rows)
            {
                DateTime month = Convert.ToDateTime(rowGifts["MonthDate"]);
                result += month.ToString("MMMM yyyy") + ": ";
                result += String.Format("{0:C}", Convert.ToDouble(rowGifts["MonthAmount"])) + "   ";
                result += String.Format(", Number of Gifts: {0}", Convert.ToDouble(rowGifts["MonthCount"]));
                result += Environment.NewLine;
            }

            return result;
        }

        private static bool SendAsEmail(TreasurerNotificationTDSMessageRow AMsg)
        {
            return !AMsg.IsEmailAddressNull() && !AMsg.IsHTMLMessageNull() && AMsg.ErrorMessage == String.Empty;
        }

        private static string GetEmailID(TreasurerNotificationTDSMessageRow AMsg)
        {
            return AMsg.TreasurerKey.ToString() + "-" + AMsg.RecipientKey.ToString();
        }

        /// <summary>
        /// generate the email text for one treasurer, one worker
        /// </summary>
        private static MailMessage CreateEmail(TreasurerNotificationTDSMessageRow AMsg, string ASenderEmailAddress)
        {
            try
            {
                MailMessage msg = new MailMessage(ASenderEmailAddress,
                    AMsg.EmailAddress,
                    AMsg.Subject,
                    AMsg.HTMLMessage);

                msg.Bcc.Add(ASenderEmailAddress);

                return msg;
            }
            catch (Exception e)
            {
                TLogging.Log(e.Message);
                TLogging.Log(AMsg.EmailAddress);
                throw e;
            }
        }

        /// <summary>
        /// send the emails
        /// </summary>
        [RequireModulePermission("FINANCE-1")]
        public static bool SendEmails(string ASendingEmailAddress, string AUserName, string AEmailPassword,
            ref TreasurerNotificationTDSMessageTable ALetters,
            out int ANumberOfEmailsSent,
            out string AErrorMessage)
        {
            ANumberOfEmailsSent = 0;
            AErrorMessage = string.Empty;

            string MyClientID = DomainManager.GClientID.ToString();

            TProgressTracker.InitProgressTracker(MyClientID,
                Catalog.GetString("Sending Emails"),
                ALetters.Rows.Count + 1);

            TProgressTracker.SetCurrentState(MyClientID, "Preparing the emails...", 0.0m);

            TSmtpSender smtp = new TSmtpSender(
                TAppSettingsManager.GetValue("SmtpHostUser"),
                TAppSettingsManager.GetInt16("SmtpPortUser", 25),
                TAppSettingsManager.GetBoolean("SmtpEnableSslUser", false),
                AUserName,
                AEmailPassword,
                string.Empty);

            int MessagesProcessed = 0;
            try
            {
                foreach (TreasurerNotificationTDSMessageRow email in ALetters.Rows)
                {
                    if (TProgressTracker.GetCurrentState(MyClientID).CancelJob == true)
                    {
                        TProgressTracker.FinishJob(MyClientID);
                        return false;
                    }

                    if (SendAsEmail(email) && email.IsDateSentNull())
                    {
                        MailMessage m = CreateEmail(email, ASendingEmailAddress);

                        try
                        {
                            if (!smtp.SendMessage(m))
                            {
                                string id = GetEmailID(email);
                                TProgressTracker.FinishJob(MyClientID);
                                AErrorMessage = "failure sending email " + id + " " + m.Subject;
                                TLogging.Log(AErrorMessage);
                                return false;
                            }
                        }
                        catch (Exception ex)
                        {
                            // already printed exception to logfile inside SendMessage
                            TProgressTracker.FinishJob(MyClientID);
                            AErrorMessage = "email server error: " + ex.Message;
                            TLogging.Log(AErrorMessage);
                            return false;
                        }

                        email.DateSent = m.Headers.Get("Date-Sent");
                        ANumberOfEmailsSent++;
                        TProgressTracker.SetCurrentState(MyClientID, "email to " + email.EmailAddress, MessagesProcessed);
                        // TODO: add email to p_partner_contact
                        // TODO: add email to sent box???
                    }

                    MessagesProcessed++;
                }
            }
            catch (Exception ex)
            {
                AErrorMessage = "There was a problem sending an email";
                TLogging.Log(AErrorMessage);
                TLogging.Log(ex.ToString());
                TProgressTracker.FinishJob(MyClientID);
                return false;
            }

            TProgressTracker.FinishJob(MyClientID);

            return true;
        }
    }
}