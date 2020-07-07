using System;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;


namespace FromExel
{
    /// <summary>
    /// Отправление писем
    /// </summary>
    class MailManager
    {
        /// <summary>
        /// Отправление письма клиенту с заказом
        /// </summary>
        /// <param name="address">Е-mail аддрес получателя</param>
        /// <param name="path">Полное имя файла, который надо прикрепить</param>
        /// <returns></returns>
        public bool SendMail(string address, string path)
        {
            Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
            //NameSpace ns = outlook.GetNamespace("mapi");
            //ns.Logon("NIMatveev ", "851570sbuw", true, true);
            //MAPIFolder mf = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            Accounts ac = outlook.Session.Accounts;
            MailItem mail = (MailItem)outlook.CreateItem(OlItemType.olMailItem);

            //string text = mf.Items[2].Body.ToString();
            
            mail.To = address;
            //mail.SendUsingAccount = ac[1];
            mail.Subject = "Заказ Типография ОТИ НИЯУ МИФИ";
            mail.Attachments.Add(@path);
            mail.Importance = OlImportance.olImportanceNormal;
            mail.Body = "С уважением,\r\nМатвеев Никита Иванович\r\nОператор Копировательно Множительных машин\r\nIP 5196";
            mail.Send();
            
            
            outlook.Quit();

            return true;
        }
    }
}
