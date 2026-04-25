using System;
using System.Windows.Forms;

namespace SimpleEmailClient
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string mailto = null;
            foreach (string arg in args)
            {
                if (arg.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                {
                    mailto = arg;
                    break;
                }
            }

            Form1 mainForm = new Form1();
            if (!string.IsNullOrEmpty(mailto))
            {
                mainForm.SetMailto(mailto);
            }
            Application.Run(mainForm);
        }
    }
}