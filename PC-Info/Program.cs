using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Management;
using System.Runtime;

namespace PC_Info
{
    class Program
    {
        static void Main(string[] args)
        {

           string SN ;
            string Fsec;
            string CN;

            


            Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            CN =   System.Environment.MachineName;

            if (Directory.Exists(@"PATH TO PROGRAM"))
            {
                Fsec = "PROGRAM is installed!";
            }
            else
            {
                Fsec = "PROGRAM is NOT installed!";
            }

            string wmic = "/c wmic bios get serialnumber > serial.txt";
            ProcessStartInfo serN = new ProcessStartInfo("cmd.exe");
            serN.Arguments = wmic;
            Process startCmd = Process.Start(serN);
            startCmd.WaitForExit();
            SN = File.ReadLines("serial.txt").Skip(1).Take(1).First();
            

            mailItem.Subject = "My PC Info";
            mailItem.To = "PUT EMAIL HERE";
            mailItem.Body = "My computer name is:" + Environment.NewLine + CN + Environment.NewLine + "---------------------------" + Environment.NewLine + Fsec + Environment.NewLine + "---------------------------" + Environment.NewLine + "Serial Number is:" + Environment.NewLine + SN;
           
            mailItem.Display(false);
            mailItem.Send();

            

          
         //   var memoryCache = MemoryCache.Default;
        //    string testCrede;
       //    testCrede = CredentialCache.DefaultNetworkCredentials.ToString();

            

        }



    }

        
           
     

}
//}
