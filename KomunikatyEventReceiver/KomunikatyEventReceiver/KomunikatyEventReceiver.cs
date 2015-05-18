﻿using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Collections;
using System.Text;
using System.Net;
using System.Collections.Specialized;
using System.Net.Mail;
using Microsoft.SharePoint.Administration;
using System.Linq;
using Microsoft.SharePoint.Linq;

namespace Bizarre.KomunikatyEventReceiver
{
    /// <summary>
    /// Na podstawie informacji w rekordzie tabQuickMessages tworzy i wysyła wiadomośći do listy klientów spełniających kryteria wyboru "serwisów".
    /// </summary>
    public class QuickMessageEventReceiver : SPItemEventReceiver
    {
        
        ArrayList MatchedCustomers = new ArrayList();
        //string SENDER_EMAIL = "noreply@stafix24.pl";
        //string SENDER_NAME = "Biuro Wirtualne - Żelazna 67";

         
         string DEFAULT_FOOTER = "Zespół obsługi biura wirtualnego";

         public static string USERNAME = "5489cdaf-fa02-47d4-9244-886d709f07a3";
         public static string  API_KEY = "5489cdaf-fa02-47d4-9244-886d709f07a3";
         public static string SENDER_EMAIL = "noreply@stafix24.pl";
         public static string SENDER_NAME = "STAFix24";
         public static string REPLY_TO = "stafix24@hotmail.com";

         string DEFAULT_SENDER = String.Format("{0}<{1}>", SENDER_EMAIL, SENDER_NAME);


        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            
            SPSecurity.RunWithElevatedPrivileges(() =>
            {
                ExecuteMain(properties);
            });
        }

        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            
            SPSecurity.RunWithElevatedPrivileges(() =>
            {
                ExecuteMain(properties);
            });
        }

        #region Procedures

        private void ExecuteMain(SPItemEventProperties properties)
        {
            this.EventFiringEnabled = false;

            try
            {
                UpdateDefauls(properties);

                SPListItem item = properties.ListItem;
                bool isSent = false;
                if (item["colWyslana"] != null)
                {
                    Boolean.TryParse(item["colWyslana"].ToString(), out isSent);
                }

                if (!isSent)
                {
                    string operatorMessage = string.Empty;

                    SPFieldLookupValueCollection searchCriteriaCollection = (SPFieldLookupValueCollection)item["colGrupyOdbiorcow"];

                    GetTargetList(searchCriteriaCollection, properties);

                    if (MatchedCustomers.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();

                        //nagłówek
                        sb.AppendLine(String.Format(@"<div>{0}</div>",
                                DateTime.Now.ToString()));

                        sb.AppendLine(String.Format(@"<div>kryteria: ({0})</div>",
                                searchCriteriaCollection.Count.ToString()));
                        sb.AppendLine(String.Format(@"<ul>"));
                        foreach (SPFieldLookupValue cItem in searchCriteriaCollection)
                        {
                            sb.AppendLine(String.Format(@"<li>{0}</li>",
                                cItem.LookupValue.ToString()));
                        }

                        sb.AppendLine(String.Format(@"</ul>"));

                        sb.AppendLine(String.Format(@"<div>odbiorcy ({0})</div>",
                            MatchedCustomers.Count.ToString()));

                        sb.AppendLine(String.Format(@"<ul>"));
                        foreach (Customer cust in MatchedCustomers)
                        {
                            if ((bool)item["colGotoweDoWysylki"])
                            {
                                //SendEmailWithAttachements(cust, properties);
                                //CreateMailRequest(cust, properties);

                                SendEmail(cust, properties, true);

                            }
                            //treść
                            sb.AppendLine(String.Format(@"<li>{0} {1}</li>",
                                cust.Name.ToString(),
                                cust.Email.ToString()));
                        }
                        sb.AppendLine(String.Format(@"</ul>"));

                        operatorMessage = sb.ToString();

                        item["colTarget"] = operatorMessage;
                        if ((bool)item["colGotoweDoWysylki"])
                        {
                            item["colWyslana"] = true;
                        }
                        item.Update();
                    }
                    else
                    {
                        operatorMessage = String.Format(@"Żaden klient nie spełnia zadanych kryteriów ({0})"
                            , searchCriteriaCollection.ToString());

                        item["colTarget"] = operatorMessage;
                        item.Update();
                    }
                }

                SPListItem li = properties.ListItem;
                li["colMemo"] = DateTime.Now.ToString();
                li.Update();


            }
            catch (Exception ex)
            {
                SPListItem li = properties.ListItem;
                li["colMemo"] = ex.ToString();
                li.Update();
                
            }

            this.EventFiringEnabled = true;
        }

        private void UpdateDefauls(SPItemEventProperties properties)
        {

            using (SPSite site = new SPSite(properties.SiteId))
            {

                using (SPWeb web = site.AllWebs[properties.Web.ID])
                {
                    SPList list = web.Lists["Setup"];

                    // get SENDER EMAIL

                    SPListItem item = (from SPListItem li in list.Items
                                       where li["KEY"].ToString() == "SENDER_EMAIL"
                                       select li).First();

                    if (item != null)
                    {
                        DEFAULT_SENDER = item["TEXT"].ToString();
                    }

                    // get SENDER NAME

                    SPListItem item1 = (from SPListItem li in list.Items
                                       where li["KEY"].ToString() == "SENDER_NAME"
                                       select li).First();

                    if (item1 != null)
                    {
                        DEFAULT_SENDER = item1["TEXT"].ToString();
                    }

                    // get REPLY_TO

                    SPListItem item5 = (from SPListItem li in list.Items
                                        where li["KEY"].ToString() == "REPLY_TO"
                                        select li).First();

                    if (item5 != null)
                    {
                        DEFAULT_SENDER = item5["TEXT"].ToString();
                    }

                    // get CONTACT FOOTER

                    SPListItem item2 = (from SPListItem li in list.Items
                                        where li["KEY"].ToString() == "CONTACT_FOOTER1"
                                        select li).First();

                    if (item2 != null)
                    {
                        DEFAULT_FOOTER = item2["MEMO"].ToString();
                    }

                    // get SMTP USER Name

                    SPListItem item3 = (from SPListItem li in list.Items
                                       where li["KEY"].ToString() == "USERNAME"
                                       select li).First();

                    if (item3 != null)
                    {
                        USERNAME = item3["TEXT"].ToString();
                    }

                    // get SMTP API Key

                    SPListItem item4 = (from SPListItem li in list.Items
                                       where li["KEY"].ToString() == "API_KEY"
                                       select li).First();

                    if (item4 != null)
                    {
                        USERNAME = item4["TEXT"].ToString();
                    }
                }
            }
        }

        private void SendEmail(Customer cust, SPItemEventProperties properties, bool formatMessage)
        {
            string strBody = string.Empty;

            if (properties.ListItem["colBody"] != null)
            {
                strBody = properties.ListItem["colBody"].ToString();
            }

            bool mailSent = SendElasticEmail(
                SENDER_EMAIL,
                SENDER_NAME,
                cust.Email,
                properties.ListItem.Title,
                strBody,
                string.Empty, formatMessage);
        }

        //private void CreateMailRequest(Customer cust, SPItemEventProperties properties)
        //{
        //    SPList tList = properties.Web.Lists["Powiadomienia"];
        //    SPListItem item = tList.AddItem();

        //    try
        //    {
        //        item["_Klient"] = cust.Id;
        //        item["_Kontakt"] = cust.Email;
        //        item["_Temat"] = ":: " + properties.ListItem.Title;
        //        item["Operator"] = properties.UserLoginName;
        //        if (properties.ListItem["Body"] != null)
        //        {
        //            item["_Tre_x015b__x0107_"] = properties.ListItem["Body"].ToString();
        //        }
        //        item["_Typ_x0020_powiadomienia"] = @"E-Mail Grupowy";
        //        if (properties.ListItem["colPlanowanyTerminWysylki"] != null)
        //        {
        //            item["Data_x0020_wysy_x0142_ki"] = properties.ListItem["colPlanowanyTerminWysylki"].ToString();
        //        }
        //        item["_OgloszenieId"] = properties.ListItemId;

        //        item.Update();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw;
        //    }
        //}

        private void GetTargetList(SPFieldLookupValueCollection searchCriteriaCollection, SPItemEventProperties properties)
        {
            ///
            StringBuilder sb = new StringBuilder(@"<OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy><Where><IsNotNull><FieldRef Name='_Serwisy' /></IsNotNull></Where>");

            string camlQuery = sb.ToString();

            using (SPSite site = new SPSite(properties.SiteId))
            {
                using (SPWeb web = site.OpenWeb(properties.Web.ID))
                {
                    ///
                    SPList list = web.Lists["Klienci"];
                    SPQuery query = new SPQuery();
                    query.Query = camlQuery;

                    SPListItemCollection items = list.GetItems(query);
                    if (items.Count > 0)
                    {
                        foreach (SPListItem item in items)
                        {
                            //sprawdź czy spełnia kryteria wyboru

                            bool isMatched = false;

                            SPFieldLookupValueCollection itemCriteriaCollection = (SPFieldLookupValueCollection)item["_Serwisy"];
                            foreach (SPFieldLookupValue itemCurrent in itemCriteriaCollection)
                            {
                                foreach (SPFieldLookupValue itemSearched in searchCriteriaCollection)
                                {
                                    if (itemCurrent.LookupId == itemSearched.LookupId)
                                    {
                                        //dodaj klienta do listy wynikowej

                                        Customer cust = new Customer(item.ID, item["_Adres_x0020_e_x002d_mail"].ToString(), item["_Nazwa"].ToString());

                                        MatchedCustomers.Add(cust);

                                        isMatched = true;
                                    }

                                    if (isMatched) break;
                                }

                                if (isMatched) break;
                            }
                        }
                    }
                }
            }
        }

        #endregion

        #region Send ElasticEmail via SMTP Client

        

        private bool SendElasticEmail(string from, string fromName, string to, string subject, string bodyHtml, string bodyText, bool formatMessage)
        {
            if (formatMessage)
            {

                StringBuilder sb = new StringBuilder(@"<body bgcolor=""#E2EBFC""><table style=""width: 100%; font-family: 'Times New Roman', Times, serif;"" cellpadding=""0"" cellspacing=""0"" border=""0""><tr><td valign=""top"" style=""text-align: center; background-color: #E2EBFC"">&nbsp;</td></tr><tr>
	    <td valign=""top"" style=""text-align: center; background-color: #E2EBFC; height: 83px;""><table style=""width: 80%; background-color: #FFFFCC;"" cellpadding=""3"" cellspacing=""0"" border=""1""><tr>
	    <td>***body***</td></tr><tr>
	    <td>***footer***</td></tr><tr>
	    <td style=""text-align: right; background-color: #E2EBFC; border-collapse:collapse"">
	    <span style=""font-size: xx-small"">Jeżeli nie chcesz otrzymywać tego typu 
	    informacji wybierz</span><a href=""{unsubscribe}"" style=""text-decoration: none""><span style=""font-size: xx-small""> 
	    Rezygnuję</span></a></td></tr></table></td></tr></table></body>");

                sb.Replace("***body***", bodyHtml);
                sb.Replace("***footer***", DEFAULT_FOOTER);
                bodyHtml = sb.ToString();

            }

            if (!subject.StartsWith("::"))
            {
                subject = @":: " + subject;
            }


            string result = SendEmail(to, subject, bodyText, bodyHtml, from, fromName);
            return TryStrToGuid(result);
        }

        public static string SendEmail(string to, string subject, string bodyText, string bodyHtml, string from, string fromName)
        {

            WebClient client = new WebClient();
            NameValueCollection values = new NameValueCollection();
            values.Add("username", USERNAME);
            values.Add("api_key", API_KEY);
            values.Add("from", from);
            values.Add("from_name", fromName);
            values.Add("reply_to", REPLY_TO);
            values.Add("reply_to_name", fromName);
            values.Add("to", to);
            values.Add("subject", subject);
            values.Add("body_html", bodyHtml);
            //values.Add("charset", "windows-1250");

            byte[] response = client.UploadValues("https://api.elasticemail.com/mailer/send", values);
            return Encoding.UTF8.GetString(response);
        }

        public static Boolean TryStrToGuid(String s)
        {
            Guid value = Guid.Empty;

            try
            {
                value = new Guid(s);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }


        #endregion
    }

    #region Helper Classes

    public class Customer
    {
        public Customer(int id, string email, string name)
        {
            Id = id;
            Email = email;
            Name = name;
        }
        public int Id { get; set; }
        public string Email { get; set; }
        public string Name { get; set; }
    }

    #endregion
}
