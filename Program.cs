using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;

namespace OutlookTest
{
    //This is a spicy meatball.
    class Program
    {
        static void Main(string[] args)
        {
            Outlook.Application app = null;

            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {
                app = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                Console.WriteLine("Got the active outlook object!");
            }
            else
            {
                app = new Outlook.Application();
                Outlook.NameSpace ns = app.GetNamespace("MAPI");
                ns.Logon("", "", Missing.Value, Missing.Value);
                ns = null;
                Console.WriteLine("Created a new session!");
            }

            if (app != null)
            {
                Outlook.NameSpace ns = app.GetNamespace("MAPI");
                foreach (Outlook.MAPIFolder folder in ns.Folders)
                {
                    Console.WriteLine("{0}", folder.Name);
                    foreach (Outlook.MAPIFolder innerFolder in folder.Folders)
                    {
                        Console.WriteLine("++{0}", innerFolder.Name);
                        if (innerFolder.Name == "Inbox")
                        {
                            Outlook.Items items = innerFolder.Items;
                            foreach (Outlook.MailItem item in items)
                            {
                                Console.WriteLine("++--{0}", item.Subject);
                                foreach (Outlook.Attachment att in item.Attachments)
                                {
                                    Console.WriteLine("++--..{0}: {1}", att.FileName, att.PathName);
                                }
                            }
                        }
                    }
                }

                Console.In.Read();

                ns.Logoff();
            }
        }
    }
}
