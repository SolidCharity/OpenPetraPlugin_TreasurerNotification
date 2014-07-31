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
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Drawing.Printing;
using Ict.Common.IO;
using Ict.Common;
using Ict.Common.Printing;
using Ict.Common.Controls;
using Ict.Petra.Client.CommonControls;
using Ict.Petra.Client.CommonControls.Logic;
using Ict.Petra.Plugins.TreasurerNotification.RemoteObjects;
using Ict.Petra.Plugins.TreasurerNotification.Data;

namespace Ict.Petra.Plugins.TreasurerNotification.Client
{
    /// manual methods for the generated window
    public partial class TFrmTreasurerNotification
    {
        private TreasurerNotificationTDSMessageTable FLetters = new TreasurerNotificationTDSMessageTable();
        private SortedList <string, MailMessage>FEmails = new SortedList <string, MailMessage>();
        private Int32 FLedgerNumber;

        /// <summary>
        /// use this ledger
        /// </summary>
        public Int32 LedgerNumber
        {
            set
            {
                FLedgerNumber = value;

                grdAllWorkers.Columns[0].Width = 30;
                grdAllWorkers.Columns[1].Width = 80;
                grdAllWorkers.Columns[2].Width = 150;
                grdAllWorkers.Columns[3].Width = 100;
                grdAllWorkers.Columns[4].Width = 150;
                grdAllWorkers.Columns[5].Width = 100;
                grdAllWorkers.Columns[6].Width = 30;
                grdAllWorkers.Columns[7].Width = 400;

                grdEmails.Columns[0].Width = 30;
                grdEmails.Columns[1].Width = 100;
                grdEmails.Columns[2].Width = 200;
                grdEmails.Columns[3].Width = 200;
                grdEmails.Columns[4].Width = 400;

                grdLetters.Columns[0].Width = 30;
                grdLetters.Columns[1].Width = 200;
                grdLetters.Columns[2].Width = 400;

                DateTime LastDayOfPreviousMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1);
                dtpLastMonth.Date = LastDayOfPreviousMonth;
            }
        }

        void RefreshStatistics()
        {
            List <string>EmailRecipientsUnique = new List <string>();
            List <Int64>LetterRecipientsUnique = new List <Int64>();
            List <Int64>WorkersAlltogether = new List <Int64>();
            List <Int64>WorkersInTransition = new List <Int64>();
            List <Int64>ExWorkersWithGifts = new List <Int64>();
            Int32 CountInvalidAddressTreasurer = 0;
            Int32 CountMissingTreasurer = 0;
            Int32 CountPagesSent = 0;

            foreach (TreasurerNotificationTDSMessageRow msg in FLetters.Rows)
            {
                if (!WorkersAlltogether.Contains(msg.RecipientKey))
                {
                    WorkersAlltogether.Add(msg.RecipientKey);
                }

                if (msg.Transition && !WorkersInTransition.Contains(msg.RecipientKey))
                {
                    WorkersInTransition.Add(msg.RecipientKey);
                }

                if (msg.ErrorMessage == "NOADDRESS")
                {
                    CountInvalidAddressTreasurer++;
                }
                else if (msg.ErrorMessage == "NOTREASURER")
                {
                    CountMissingTreasurer++;
                }
                else if (msg.ErrorMessage == "EXWORKER")
                {
                    if (!ExWorkersWithGifts.Contains(msg.RecipientKey))
                    {
                        ExWorkersWithGifts.Add(msg.RecipientKey);
                    }
                }
                else
                {
                    if (SendAsEmail(msg) && !EmailRecipientsUnique.Contains(msg.TreasurerName))
                    {
                        EmailRecipientsUnique.Add(msg.TreasurerName);
                    }

                    if (SendAsLetter(msg) && !LetterRecipientsUnique.Contains(msg.TreasurerKey))
                    {
                        LetterRecipientsUnique.Add(msg.TreasurerKey);
                    }

                    if (SendAsLetter(msg) || SendAsEmail(msg))
                    {
                        CountPagesSent++;
                    }
                }
            }

            txtTreasurersEmail.Text = EmailRecipientsUnique.Count.ToString();
            txtTreasurersLetter.Text = LetterRecipientsUnique.Count.ToString();
            txtNumberOfWorkersReceivingDonations.Text = WorkersAlltogether.Count.ToString();
            txtWorkersWithTreasurer.Text = (WorkersAlltogether.Count - CountMissingTreasurer).ToString();
            txtNumberOfUniqueTreasurers.Text = (EmailRecipientsUnique.Count + LetterRecipientsUnique.Count).ToString();
            txtWorkersWithoutTreasurer.Text = CountMissingTreasurer.ToString();
            txtTreasurerInvalidAddress.Text = CountInvalidAddressTreasurer.ToString();
            txtWorkersInTransition.Text = WorkersInTransition.Count.ToString();
            txtPagesSent.Text = CountPagesSent.ToString();
            txtExWorkersWithGifts.Text = ExWorkersWithGifts.Count.ToString();
        }

        void RefreshGridEmails()
        {
            TreasurerNotificationTDSMessageTable data = new TreasurerNotificationTDSMessageTable();

            data.DefaultView.Sort = TreasurerNotificationTDSMessageTable.GetIdDBName();
            data.DefaultView.AllowNew = false;

            foreach (TreasurerNotificationTDSMessageRow email in FLetters.Rows)
            {
                if (SendAsEmail(email))
                {
                    string id = GetEmailID(email);

                    if (!FEmails.ContainsKey(id))
                    {
                        FEmails.Add(id, CreateEmail(email, txtEmailUser.Text));
                    }

                    TreasurerNotificationTDSMessageRow row = data.NewRowTyped();
                    row.Id = data.Rows.Count + 1;
                    row.DateSent = FEmails[id].Headers.Get("Date-Sent");
                    row.EmailAddress = FEmails[id].To.ToString();
                    row.TreasurerName = email.TreasurerName;
                    row.Subject = FEmails[id].Subject;
                    row.SimpleMessageText = email.SimpleMessageText;
                    row.RecipientKey = email.RecipientKey;
                    row.TreasurerKey = email.TreasurerKey;
                    data.Rows.Add(row);
                }
            }

            grdEmails.DataSource = new DevAge.ComponentModel.BoundDataView(data.DefaultView);
        }

        void RefreshGridLetters()
        {
            TreasurerNotificationTDSMessageTable data = new TreasurerNotificationTDSMessageTable();

            data.DefaultView.Sort = TreasurerNotificationTDSMessageTable.GetIdDBName();
            data.DefaultView.AllowNew = false;

            foreach (TreasurerNotificationTDSMessageRow letter in FLetters.Rows)
            {
                if (SendAsLetter(letter))
                {
                    TreasurerNotificationTDSMessageRow row = data.NewRowTyped();
                    row.Id = data.Rows.Count + 1;
                    row.TreasurerName = letter.TreasurerName;
                    row.SimpleMessageText = letter.SimpleMessageText;
                    row.RecipientKey = letter.RecipientKey;
                    row.TreasurerKey = letter.TreasurerKey;
                    data.Rows.Add(row);
                }
            }

            grdLetters.DataSource = new DevAge.ComponentModel.BoundDataView(data.DefaultView);
        }

        void RefreshWorkers()
        {
            TreasurerNotificationTDSMessageTable data = new TreasurerNotificationTDSMessageTable();

            data.DefaultView.Sort = TreasurerNotificationTDSMessageTable.GetIdDBName();
            data.DefaultView.AllowNew = false;
            List <Int64>workers = new List <Int64>();

            foreach (TreasurerNotificationTDSMessageRow m in FLetters.Rows)
            {
                string localisedErrormessage = m.ErrorMessage;

                if (localisedErrormessage == "NOTREASURER")
                {
                    localisedErrormessage = Catalog.GetString("No treasurer assigned to this worker");
                }
                else if (localisedErrormessage == "NOADDRESS")
                {
                    localisedErrormessage = Catalog.GetString("There is no valid address for the treasurer");
                }
                else if (localisedErrormessage == "EXWORKER")
                {
                    localisedErrormessage = Catalog.GetString("The Worker has left the organisation and is not in TRANSITION anymore");
                }

                TreasurerNotificationTDSMessageRow row = data.NewRowTyped();
                row.Id = data.Rows.Count + 1;
                row.Action = m.ErrorMessage.Length > 0 ? "Nothing" : (!m.IsEmailAddressNull() ? "Email" : "Letter");
                row.TreasurerName = m.TreasurerName;
                row.TreasurerKey = m.TreasurerKey;
                row.RecipientName = m.RecipientName;
                row.RecipientKey = m.RecipientKey;
                row.Transition = m.Transition;
                row.ErrorMessage = localisedErrormessage;
                row.Subject = m.Subject;
                row.SimpleMessageText = m.SimpleMessageText;
                data.Rows.Add(row);
            }

            grdAllWorkers.DataSource = new DevAge.ComponentModel.BoundDataView(data.DefaultView);
        }

        /// <summary>
        /// display the email in the web browser control below the list
        /// </summary>
        void DisplayEmail(MailMessage AEmailMessage)
        {
            MailMessage selectedMail = AEmailMessage;

            if (selectedMail == null)
            {
                // should not get here
                return;
            }

            string header = "<html><body>";
            header += String.Format("{0}: {1}<br/>",
                Catalog.GetString("From"),
                selectedMail.From.ToString());
            header += String.Format("{0}: {1}<br/>",
                Catalog.GetString("To"),
                selectedMail.To);

            if (selectedMail.CC.Count > 0)
            {
                header += String.Format("{0}: {1}<br/>",
                    Catalog.GetString("Cc"),
                    selectedMail.CC);
            }

            if (selectedMail.Bcc.Count > 0)
            {
                header += String.Format("{0}: {1}<br/>",
                    Catalog.GetString("Bcc"),
                    selectedMail.Bcc);
            }

            header += String.Format("<b>{0}: {1}</b><br/>",
                Catalog.GetString("Subject"),
                selectedMail.Subject);
            header += "<hr></body></html>";

            if (selectedMail.IsBodyHtml)
            {
                brwEmailPreview.DocumentText = header + selectedMail.Body;
            }
            else
            {
                brwEmailPreview.DocumentText = header +
                                               "<html><body>" + selectedMail.Body +
                                               "</body></html>";
            }
        }

        private Int32 FNumberOfPages = 0;
        private TGfxPrinter FGfxPrinter = null;

        /// <summary>
        /// display all letters in the web browser control below the list
        /// </summary>
        void PreparePrintLetters()
        {
            if (FLetters == null)
            {
                return;
            }

            System.Drawing.Printing.PrintDocument printDocument = new System.Drawing.Printing.PrintDocument();
            bool printerInstalled = printDocument.PrinterSettings.IsValid;

            if (printerInstalled)
            {
                string AllLetters = String.Empty;

                foreach (TreasurerNotificationTDSMessageRow letter in FLetters.Rows)
                {
                    if (SendAsLetter(letter))
                    {
                        if (AllLetters.Length > 0)
                        {
                            // AllLetters += "<div style=\"page-break-before: always;\"/>";
                            string body = letter.HTMLMessage.Substring(letter.HTMLMessage.IndexOf("<body"));
                            body = body.Substring(0, body.IndexOf("</html"));
                            AllLetters += body;
                        }
                        else
                        {
                            // without closing html
                            AllLetters += letter.HTMLMessage.Substring(0, letter.HTMLMessage.IndexOf("</html"));
                        }
                    }
                }

                if (AllLetters.Length > 0)
                {
                    AllLetters += "</html>";
                }

                FGfxPrinter = new TGfxPrinter(printDocument, TGfxPrinter.ePrinterBehaviour.eFormLetter);
                try
                {
                    TPrinterHtml htmlPrinter = new TPrinterHtml(AllLetters,
                        System.IO.Path.GetDirectoryName(txtPathHTMLTemplate.Text),
                        FGfxPrinter);
                    FGfxPrinter.Init(eOrientation.ePortrait, htmlPrinter, eMarginType.ePrintableArea);
                    this.ppvLetters.InvalidatePreview();
                    this.ppvLetters.Document = FGfxPrinter.Document;
                    this.ppvLetters.Zoom = 1;
                    FGfxPrinter.Document.EndPrint += new PrintEventHandler(this.EndPrint);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }

        private void EndPrint(object ASender, PrintEventArgs AEv)
        {
            FNumberOfPages = FGfxPrinter.NumberOfPages;
            RefreshPagePosition();
        }

        void SendEmails(object sender, EventArgs e)
        {
            if (txtEmailPassword.Text.Trim().Length == 0)
            {
                MessageBox.Show("please enter Email password", "error");
                return;
            }

            if (MessageBox.Show(Catalog.GetString("Do you really want to send the emails?"),
                    Catalog.GetString("Confirm sending emails"),
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2) == DialogResult.Cancel)
            {
                return;
            }

            // TODO fortschrittsbalken
            // TODO: dialog that allows to cancel???
            // TODO send from the server???
            TSmtpSender smtp = new TSmtpSender(
                TAppSettingsManager.GetValue("SmtpHost"),
                TAppSettingsManager.GetInt16("SmtpPort", 25),
                TAppSettingsManager.GetBoolean("SmtpEnableSsl", false),
                txtEmailUser.Text,
                txtEmailPassword.Text,
                string.Empty);

            // send the emails
            this.UseWaitCursor = true;
            int count = 0;

            foreach (TreasurerNotificationTDSMessageRow email in FLetters.Rows)
            {
                if (SendAsEmail(email))
                {
                    string id = GetEmailID(email);

                    if (!smtp.SendMessage(FEmails[id]))
                    {
                        RefreshGridEmails();
                        return;
                    }

                    count++;
                    // TODO: add email to p_partner_contact
                    // TODO: add email to sent box???
                }
            }

            RefreshGridEmails();

            this.UseWaitCursor = false;
            MessageBox.Show(count.ToString() + " Emails have been sent successfully!");
        }

        void GenerateLetters(object sender, EventArgs e)
        {
            if (grdEmails.Rows.Count > 0)
            {
                MessageBox.Show("will not reload because emails might have been sent already");
                return;
            }

            Cursor = Cursors.WaitCursor;

            string HTMLTemplate;
            try
            {
                using (StreamReader sr = new StreamReader(txtPathHTMLTemplate.Text))
                {
                    HTMLTemplate = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Cannot read file " + txtPathHTMLTemplate.Text, "Failure");
                return;
            }

            TMTeasurerNotificationNamespace TRemote = new TMTeasurerNotificationNamespace();
            FLetters = TRemote.Server.WebConnectors.GenerateMessages(
                FLedgerNumber,
                HTMLTemplate,
                // TODO do not hardcode motivation group and detail
                "GIFT",
                "SUPPORT",
                chkLettersOnly.Checked,
                dtpLastMonth.Date.Value,
                Convert.ToInt16(nudNumberMonths.Value));

            Cursor = Cursors.Default;

            if (FLetters.Count == 0)
            {
                MessageBox.Show(Catalog.GetString("There are no gifts in the given period of time"));
                return;
            }

            RefreshWorkers();
            RefreshGridLetters();
            RefreshGridEmails();
            PreparePrintLetters();
            RefreshStatistics();
        }

        void GrdAllWorkersDoubleClick(object sender, System.EventArgs e)
        {
            TreasurerNotificationTDSMessageRow r =
                (TreasurerNotificationTDSMessageRow)((DataRowView)grdEmails.SelectedDataRows[0]).Row;

            MessageBox.Show(r.SimpleMessageText);
        }

        private void EmailFocusedRowChanged(object sender, EventArgs e)
        {
            TreasurerNotificationTDSMessageRow r =
                (TreasurerNotificationTDSMessageRow)((DataRowView)grdEmails.SelectedDataRows[0]).Row;
            string id = GetEmailID(r);

            DisplayEmail(FEmails[id]);
        }

        /// get the currently selected row, the appropriate ID, and scroll to the printed page
        void LetterFocusedRowChanged(object sender, System.EventArgs e)
        {
            TreasurerNotificationTDSMessageRow r =
                (TreasurerNotificationTDSMessageRow)((DataRowView)grdEmails.SelectedDataRows[0]).Row;

            ppvLetters.StartPage = r.Id - 1;
            RefreshPagePosition();
        }

        private bool SendAsEmail(TreasurerNotificationTDSMessageRow AMsg)
        {
            return !AMsg.IsEmailAddressNull() && !AMsg.IsHTMLMessageNull() && AMsg.ErrorMessage == String.Empty;
        }

        private bool SendAsLetter(TreasurerNotificationTDSMessageRow AMsg)
        {
            return AMsg.IsEmailAddressNull() && !AMsg.IsHTMLMessageNull() && AMsg.ErrorMessage == String.Empty;
        }

        private string GetEmailID(TreasurerNotificationTDSMessageRow AMsg)
        {
            return AMsg.TreasurerKey.ToString() + "-" + AMsg.RecipientKey.ToString();
        }

        /// <summary>
        /// generate the email text for one treasurer, one worker
        /// </summary>
        private MailMessage CreateEmail(TreasurerNotificationTDSMessageRow AMsg, string ASenderEmailAddress)
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
    }
}