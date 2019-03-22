using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SalesOrdersImport.Helpers
{
    class ShowFolder
    {
        public string value { get; set; }
        public delegate void myDeleg(string value);
        public event myDeleg currFunc;



        //კონსტრუქტორი, მხოლოდ ინიციალიზაციისთვის. დეფოლთი კი ისეც აქვს მარა რავიცი რას აშავებს იყოს
        public ShowFolder()
        {

        }


        //ცალკე სრედი მეთოდის გასახსნელად, რადგან ამ დროს აპლიკაციაა გაშვებული 1 ცალ სრედში და მასზე ფეილის გაშვება
        //ახალი სრედის გარეშე არ გამოდის
        public void loadFolder()
        {
            try
            {
                Thread ShowFolderBrowserThread = new System.Threading.Thread(setValue);

                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {

                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);

                    ShowFolderBrowserThread.Start();

                }

                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {

                    ShowFolderBrowserThread.Start();

                    ShowFolderBrowserThread.Join();

                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message + " Thread Problem", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        //folder open method
        private void setValue()
        {
            try
            {
                NativeWindow nws = new NativeWindow();

                OpenFileDialog fdb = new OpenFileDialog();

                //  myProcs = Process.GetProcessesByName("SAP Business One");

                nws.AssignHandle(Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);

                // var mSboForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm; //GET ACTIVE FORM

                if (fdb.ShowDialog(nws) == DialogResult.OK)
                {
                    string test = fdb.FileName;
                    value = test;
                    currFunc(test);
                }
            }
            catch (Exception ex)
            {

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message + " Folder Cant Open", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
    }

}
