using System;
using System.IO;
using MsgReader;
using OpenMcdf;
using System.Management.Automation;

namespace MsgFiles
{
    /// <summary>
    /// <para type="synopsis">Reads emails saved in Microsoft Outlook's .msg files.</para>
    /// <para type="description">The Cmdlet Read-MsgFile automates the extraction of key email data like from, to, cc, date, subject and body from emails saved in Microsoft Outlook's .msg files.</para>
    /// </summary>
    /// <para type="link" uri="https://www.datenteiler.de/">Homepage of datenteiler</para>   
    /// <example>
    ///   <para>Read a .msg file:</para>
    ///   <code>Read-MsgFile -File sample.msg</code>
    /// </example>
    [Cmdlet(VerbsCommunications.Read, "MsgFile")]
    public class ReadMsgFile : PSCmdlet
    {
        /// <summary>
        /// <para type="description">Path to or the .msg file.</para>
        /// </summary>
        [Parameter(
            Position = 0,
            Mandatory = true,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public string File { get; set; } = string.Empty;

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            string fileName = this.File;

            // Get the current working directory using the core command API
            using (var msg = new MsgReader.Outlook.Storage.Message(System.IO.Path.Combine(SessionState.Path.CurrentLocation.Path.ToString(), fileName)))
            {
                string from = msg.GetEmailSender(false, false); // msg.Sender;
                string sentTo = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, false, false);
                string sentCc = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, false, false);
                var sentOn = msg.SentOn;
                string subject = msg.Subject;
                int CountAttachments = msg.Attachments.Count;
                
                // Create PowerShell object
                PSObject responseObject = new PSObject();

                // Add members to the object
                if (!String.IsNullOrEmpty(from)) 
                { 
                    responseObject.Members.Add(new PSNoteProperty("From", from)); 
                }
                else
                {
                    responseObject.Members.Add(new PSNoteProperty("From", string.Empty)); 
                }

                if (!String.IsNullOrEmpty(sentTo)) 
                { 
                    responseObject.Members.Add(new PSNoteProperty("To", sentTo)); 
                }
                else
                {
                    responseObject.Members.Add(new PSNoteProperty("To", string.Empty));
                }
                if (!String.IsNullOrEmpty(sentCc)) 
                { 
                    responseObject.Members.Add(new PSNoteProperty("CC", sentCc)); 
                }
                else
                {
                    responseObject.Members.Add(new PSNoteProperty("CC", string.Empty));
                }
                if (sentOn.HasValue) 
                { 
                    responseObject.Members.Add(new PSNoteProperty("Sent", sentOn.Value)); 
                }
                else
                {
                    responseObject.Members.Add(new PSNoteProperty("Sent", string.Empty));
                }
                responseObject.Members.Add(new PSNoteProperty("Attachments", CountAttachments));
                if (!String.IsNullOrEmpty(subject)) 
                { 
                    responseObject.Members.Add(new PSNoteProperty("Subject", subject)); 
                }
                else
                {
                    responseObject.Members.Add(new PSNoteProperty("Subject", string.Empty));
                }
                if (String.IsNullOrEmpty(msg.BodyText))
                {
                    responseObject.Members.Add(new PSNoteProperty("Body", msg.BodyHtml));
                }
                else
                {
                    responseObject.Members.Add(new PSNoteProperty("Body", msg.BodyText));
                }

                // Write PSObject
                this.WriteObject(responseObject);               
            }
        }
    }

    /// <summary>
    /// <para type="synopsis">Extracts attachments saved in Microsoft Outlook's .msg files.</para>
    /// <para type="description">Extracts the email's attachments saved in a Microsoft Outlook's .msg file and saves it to your path.</para>
    /// </summary>
    /// <example>
    ///   <para>Extracts attachments from a .msg file:</para>
    ///   <code>Get-MsgAttachment -File sample.msg -Path /path/to/extract -Verbose</code>
    /// </example>
    [Cmdlet(VerbsCommon.Get, "MsgAttachment")]
    public class GetMsgAttachment : PSCmdlet
    {
        /// <summary>
        /// <para type="description">Path to or the .msg file.</para>
        /// </summary>
        [Parameter(
           Position = 0,
            Mandatory = true,
           ValueFromPipeline = true,
           ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public string File { get; set; } = string.Empty;

        /// <summary>
        /// <para type="description">Where to extract the attachment(s).</para>
        /// </summary>
        [Parameter(
           Position = 1,
            Mandatory = true,
           ValueFromPipeline = true,
           ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public string Path { get; set; } = string.Empty;

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            string fileName = this.File;
            string path = this.Path;
            // Get the current working directory using the core command API
            using (var msg = new MsgReader.Outlook.Storage.Message(System.IO.Path.Combine(SessionState.Path.CurrentLocation.Path.ToString(), fileName)))
            {
                foreach (MsgReader.Outlook.Storage.Attachment attach in msg.Attachments)
                {
                    string combined = System.IO.Path.Combine(path, System.IO.Path.GetFileName(attach.FileName));
                    this.WriteVerbose("Write Attachment " + combined);
                    byte[] attachBytes = attach.Data;
                    FileStream attachStream = System.IO.File.Create(combined);
                    attachStream.Write(attachBytes, 0, attachBytes.Length);
                    attachStream.Close();
                }
            }
        }
    }
}
