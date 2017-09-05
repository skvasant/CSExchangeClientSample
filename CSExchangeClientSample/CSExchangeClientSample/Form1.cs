using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CSExchangeClientSample
{
    public partial class Form1 : Form
    {
        private ExchangeService _exchange;

        public Form1()
        {
            InitializeComponent();
            button1.Location = new Point(0, 240);
            button1.Height = 25;
            button1.Text = "Reply";
            button2.Location = new Point(50, 240);
            button2.Height = 25;
            button2.Text = "Cancel";
            button3.Location = new Point(725, 240);
            button3.Height = 25;
            button3.Width = 55;
            button3.Text = "Refresh";
            textBox1.Location = new Point(0, 265);
            textBox1.Multiline = true;
            textBox1.Height = 335;
            textBox1.Width = 784;
            textBox2.Location = new Point(102, 242);
            textBox2.Width = 600;
            Height = 600;
            Width = 800;
        }

        private void InitializeView()
        {
            listView1.Clear();
            listView1.View = View.Details;
            listView1.Columns.Add("From", 150);
            listView1.Columns.Add("Subject", 334);
            listView1.Columns.Add("Date", 150);
            listView1.Columns.Add("Size", 145);
            listView1.Columns.Add("Id", 0);
            listView1.FullRowSelect = true;
            listView1.Height = 240;
            listView1.Dock = DockStyle.Top;
            listView1.MultiSelect = false;
            textBox1.ResetText();
            textBox2.ResetText();
            button1.Text = "Reply";
        }

        public void ConnectToExchangeServer()
        {
            var username = ConfigurationManager.AppSettings["username"];
            // TBD: Configuration value is the password encoded in base64, I know not security by obscurity threw this in for sample
            // further development can be done to implement encryption/decryption or have the user enter the password everytime, whatever
            // works!
            var password = Encoding.ASCII.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings["password"]));
            var domain = ConfigurationManager.AppSettings["domain"];
            var emailAddress = ConfigurationManager.AppSettings["emailAddress"];
            Text = @"Connecting to Exchange Server..";
            try
            {
                // TBD: Change the exchange version as desired, in this case its Exchange 2010 SP3 (Yes it uses Exchange2010_SP2 definition)
                _exchange = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
                {
                    Credentials = new WebCredentials(username, password, domain)
                };
                _exchange.AutodiscoverUrl(emailAddress);

                Text = @"Connected to Exchange Server : " + _exchange.Url.Host;
            }
            catch (Exception ex)
            {
                Text = @"Error Connecting to Exchange Server!!" + ex.Message;
            }

        }

        public void GetMessages()
        {
            var value = ConfigurationManager.AppSettings["name"];
            var ts = new TimeSpan(-5, 0, 0, 0);
            var date = DateTime.Now.Add(ts);
            var filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);

            if (_exchange == null) return;
            var findResults = _exchange.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(1000));
            if (findResults != null)
            {
                InitializeView();
                foreach (var item in findResults)
                {
                    // TBD: Any custom business logic, in this case I am checking if the message has an attachment and the
                    // person who replied last in the message thread is not the name from configuration value
                    if (item != null && (!item.HasAttachments ||
                                         item.LastModifiedName.Contains(value))) continue;
                    if (item == null) continue;
                    var listitem = new ListViewItem(new[]
                    {
                        item.LastModifiedName, item.Subject, item.LastModifiedTime.ToString(CultureInfo.CurrentCulture), item.Size.ToString(),
                        item.Id.ToString()
                    });
                    listView1?.Items.Add(listitem);
                }
                if (findResults.Items.Count <= 0)
                {
                    listView1?.Items.Add("No Messages found!!");
                }
            }
            else
            {
                throw new ArgumentNullException(nameof(findResults));
            }
        }

        private void listView1_Click(object sender, EventArgs e)
        {
            var msgId = listView1.SelectedItems[0].SubItems[4].Text;
            if (msgId == null) throw new ArgumentNullException(nameof(msgId));
            //MessageBox.Show(id);
            if (_exchange == null) return;
            var message = EmailMessage.Bind(_exchange, new ItemId(msgId));
            if (message == null) throw new ArgumentNullException(nameof(message));
            textBox1.ResetText();
            textBox1.Text = $@"To {message.DisplayTo}
Cc {message.DisplayCc}
{message.Body.Text}";
            textBox2.ResetText();
            button1.Text = "Reply";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GetFilteredInbox();
            button1.Text = "Reply";
        }

        private void GetFilteredInbox()
        {
            if (_exchange == null)
            {
                ConnectToExchangeServer();
            }
            GetMessages();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.ResetText();
            textBox2.ResetText();
            button1.Text = "Reply";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var msgId = listView1.SelectedItems[0].SubItems[4].Text;
            if (msgId == null) throw new ArgumentNullException(nameof(msgId));
            //MessageBox.Show(id);
            if (_exchange == null) return;
            var message = EmailMessage.Bind(_exchange, new ItemId(msgId));
            if (message == null) throw new ArgumentNullException(nameof(message));

            if (button1.Text == "Reply")
            {
                // TBD: Any custom business logic, in this case I parse for specific text in the email and set a target
                // path, so while the send button is clicked remainder of the business logic is performed
                var workflow = string.Empty;
                var filename = message.InternetMessageId.Substring(1,
                                   message.InternetMessageId.IndexOf("@", StringComparison.Ordinal) - 1) + ".pdf";
                if (message.Subject.ToUpper().Contains("WORD") || message.Body.Text.ToUpper().Contains("WORD"))
                {
                    workflow = "WORD";
                    textBox2.Text = $@"\\servername\root directory\sub directory\workflow1\{filename}";
                }
                else if (message.Subject.ToUpper().Contains("TEXT") || message.Body.Text.ToUpper().Contains("TEXT"))
                {
                    workflow = "TEXT";
                    textBox2.Text = $@"\\servername\root directory\sub directory\workflow2\{filename}";
                }
                else if (message.Subject.ToUpper().Contains("143") || message.Body.Text.ToUpper().Contains("143"))
                {
                    workflow = "143";
                    textBox2.Text = $@"\\servername\root directory\sub directory\workflow3\{filename}";
                }
                var ti = CultureInfo.CurrentCulture.TextInfo;
                textBox1.ResetText();
                textBox1.Text = $@"Hi {ti.ToTitleCase(message.From.Name.ToLower())}:

File renamed to {filename} and processed to {workflow}.

Kind Regards,
{ConfigurationManager.AppSettings["name"]}";
            }
            else if (button1.Text == "Send")
            {
                if (string.IsNullOrWhiteSpace(textBox2.Text)) return;

                // TBD: Any custom business logic, in this case I download the attachment and copy it to target folder
                if (!message.HasAttachments || !(message.Attachments[0] is FileAttachment)) return;
                const string temp = @"c:\temp";
                if (!Directory.Exists(temp))
                {
                    Directory.CreateDirectory(temp);
                }

                FileAttachment fileAttachment = message.Attachments[0] as FileAttachment;
                // Change the below Path
                fileAttachment.Load($@"{temp}\{fileAttachment.Name}");

                if (!Directory.Exists(
                    textBox2.Text.Substring(0, textBox2.Text.LastIndexOf(@"\", StringComparison.Ordinal)))) return;
                File.Copy($@"{temp}\{fileAttachment.Name}", textBox2.Text);

                // TBD: Any custom business logic, in this case I create a reply and add a cc to an email address
                var replyAllMessage = message.CreateReply(true);
                if (replyAllMessage != null)
                {
                    replyAllMessage.ToRecipients.Add(message.From);

                    replyAllMessage.CcRecipients.AddRange(message.ToRecipients);
                    replyAllMessage.CcRecipients.AddRange(message.CcRecipients);

                    EmailAddress ccEmail = new EmailAddress("Full Name", "email@company.com", "SMTP");
                    replyAllMessage.CcRecipients.Add(ccEmail);
                    replyAllMessage.BodyPrefix = new MessageBody(BodyType.Text, textBox1.Text);
                    // Send and save a copy of the replied email message in the default Sent Items folder. 
                    replyAllMessage.SendAndSaveCopy();
                }

                textBox1.ResetText();
                textBox2.ResetText();
                GetFilteredInbox();
            }
            ChangeButton1Text();

        }

        private void ChangeButton1Text()
        {
            button1.Text = button1.Text == @"Reply" ? "Send" : "Reply";
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            GetFilteredInbox();
        }
    }
}
