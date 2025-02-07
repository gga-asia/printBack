using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    internal class EmailSender
    {
        private static readonly HttpClient client = new HttpClient();

        // Send email function
        public async Task SendEmailAsync(string mailTo, string mailCC, string mailBCC, string subject, string body)
        {
            var requestUrl = "https://wpa.bionetcorp.com:9004/Api/Mail/Send";

            // The body of the POST request
            var mailContent = new
            {
                Subject = subject,
                //From = "PrintKernel@BionetCorp.com",
                From = "",
                To = new string[] { mailTo },
                Cc = new string[] { mailCC },
                Bcc = new string[] { mailBCC },
                Body = "<pre>"+body+"</pre>",
                IsBodyHtml = true,
                SendStatus = 0
            };

            // Convert the object to a JSON string
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(mailContent);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            try
            {
                // Make the POST request
                var response = await client.PostAsync(requestUrl, content);

                // Check if the request was successful
                if (response.IsSuccessStatusCode)
                {
                    //MessageBox.Show("Email sent successfully!");
                }
                else
                {
                    //MessageBox.Show($"Failed to send email. Status code: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error occurred: {ex.Message}");
            }
        }
    }
}
