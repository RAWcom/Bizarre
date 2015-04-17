using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Collections;
using System.Text;

namespace Bizarre.OgloszeniaEventReceiver
{  
    /// <summary>
    /// List Item Events
    /// </summary>
    public class OgloszeniaEventReceiver : SPItemEventReceiver
    {

        ArrayList MatchedCustomers = new ArrayList();

       /// <summary>
       /// An item was added.
       /// </summary>
       public override void ItemAdded(SPItemEventProperties properties)
       {
           base.ItemAdded(properties);
           ExecuteMain(properties);
       }

       /// <summary>
       /// An item was updated.
       /// </summary>
       public override void ItemUpdated(SPItemEventProperties properties)
       {
           base.ItemUpdated(properties);
           ExecuteMain(properties);
       }

       #region Main

       private void ExecuteMain(SPItemEventProperties properties)
       {
           this.EventFiringEnabled = false;

           SPListItem item = properties.ListItem;
           bool isSent = false;
           if (item["colWyslany"]!=null)
           {
               Boolean.TryParse(item["colWyslany"].ToString(), out isSent);
           }
           
           
           if (!isSent)
           {

           if (item.ContentType.Name == "Ogłoszenie")
           {
               SPFieldLookupValueCollection searchCriteriaCollection = (SPFieldLookupValueCollection)item["colGrupyOdbiorcow"];

               GetTargetList(searchCriteriaCollection, properties);

               if (MatchedCustomers.Count>0)
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
                           CreateMailRequest(cust, properties);
                       }
                       //treść
                       sb.AppendLine(String.Format(@"<li>{0} {1}</li>",
                           cust.Name.ToString(),
                           cust.Email.ToString()));
                   }
                   sb.AppendLine(String.Format(@"</ul>"));

                   item["colTarget"] = sb.ToString();
                   if ((bool)item["colGotoweDoWysylki"])
                   {
                       item["colWyslany"] = true;
                   }
                   item.Update();
               }
               else
               {
                   item["colTarget"] = String.Format(@"Żaden klient nie spełnia zadanych kryteriów ({0})"
                       ,searchCriteriaCollection.ToString());
                   item.Update();
               }
           }

           }

           this.EventFiringEnabled = true;
       }

       private void CreateMailRequest(Customer cust, SPItemEventProperties properties)
       {
           SPList tList = properties.Web.Lists["Powiadomienia"];
           SPListItem item = tList.AddItem();

           try
           {
               item["_Klient"] = cust.Id;
               item["_Kontakt"] = cust.Email;
               item["_Temat"] = ":: " + properties.ListItem.Title;
               item["Operator"] = properties.UserLoginName;
               if (properties.ListItem["Body"] != null)
               {
                   item["_Tre_x015b__x0107_"] = properties.ListItem["Body"].ToString();
               }
               item["_Typ_x0020_powiadomienia"] = @"E-Mail Grupowy";
               if (properties.ListItem["colPlanowanyTerminWysylki"] != null)
               {
                   item["Data_x0020_wysy_x0142_ki"] = properties.ListItem["colPlanowanyTerminWysylki"].ToString();
               }
               item["_OgloszenieId"] = properties.ListItemId;

               item.Update();
           }
           catch (Exception ex)
           {
               throw;
           }
       }

       //private void SendEmailWithAttachements(Customer cust, SPItemEventProperties properties)
       //{
          
       //    using (SPSite site = new SPSite(properties.SiteId))
       //    {
       //        using (SPWeb web = properties.Web.Site.OpenWeb(properties.Web.ID))
       //        {
       //            System.Web.HttpContext oldContext = HttpContext.Current;

       //            string eMailFrom = @"noreply@stafix24.pl"; //properties.Web.Site.WebApplication.OutboundMailSenderAddress;
       //            string eMailTo = cust.Email;
       //            string eMailSubject = properties.ListItem.Title;
       //            string eMailBody = properties.ListItem["Body"].ToString();
       //            string hostAddress = string.Empty;

       //            SPListItem itemA = properties.ListItem;
       //            MailMessage mailMessage = mailInformation(itemA, web, eMailFrom, eMailTo, eMailSubject, eMailBody);
       //            hostAddress = web.Site.WebApplication.OutboundMailServiceInstance.Server.Address;
       //            sendEmail(mailMessage, hostAddress);

       //            HttpContext.Current = oldContext;
       //        }
       //    }
       //}

        //using System.Net.Mail
        //private static MailMessage mailInformation(SPListItem listItem, SPWeb spWeb, string from, string to, string subject, string body)
        //{
        //    MailMessage mail = new MailMessage();
        //    mail.From = new MailAddress(from);
        //    mail.To.Add(new MailAddress(to));
        //    mail.Subject = subject;
        //    mail.Body = body;
        //    for (int attachment = 0; attachment < listItem.Attachments.Count; attachment++)
        //    {
        //        string fileURL = listItem.Attachments.UrlPrefix + listItem.Attachments[attachment];
        //        SPFile file = spWeb.GetFile(fileURL);
        //        mail.Attachments.Add(new Attachment(file.OpenBinaryStream(), file.Name));
        //    }
        //    return mail;
        //}

        //private static void sendEmail(MailMessage eMail, string host)
        //{
        //    SmtpClient smtp = new SmtpClient(host);
        //    smtp.UseDefaultCredentials = true;
        //    smtp.Send(eMail);
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
                                   if (itemCurrent.LookupId==itemSearched.LookupId)
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
    }

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
}

