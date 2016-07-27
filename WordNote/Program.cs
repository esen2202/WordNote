/// <summary>
/// Program Author  :   Erkan ESEN (Esen Software And Design)
/// Created Date    :   2016.07.11 22:36
/// Revision Date   :   2016.07.11 22:36
/// Description     :   Bu Program *.* Dosyaların Shell Context Menüsüne Eklenir.
///                     Tıkladığında Seçili Olan Tüm Dosyaların Adreslerini Argüman Olarak Alır.
///                     Sistemde Bulunan Microsoft Word Programının Uzantısına Göre Her Seçili Dosya İçin
///                     Dosya ile Aynı Dizine Aynı Dosya İsmi ile Word Dökümanı Oluşturur ve
///                     Word Dökümanına Veri Girilmek Üzere Açar.
/// Communication   :   erkanesen2202@gmail.com
///                 :   2016.07.27  Added SourceControl
/// </summary>

using System;
using System.IO;
using System.Windows.Forms;

namespace WordNote
{
    static class Program
    {
 
        static string WordExtention = ".docx";
        static string filePath; 

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(String[] arguments)
        {
            if (!OfficeControl())
            {
                return;
            }

            if (arguments.Length > 0)
            {
                filePath = arguments[0];
                if (!File.Exists(filePath.Replace(Path.GetExtension(filePath), "") + WordExtention))
                {
                    File.Create(filePath.Replace(Path.GetExtension(filePath), "") + WordExtention);
                }
                System.Diagnostics.Process.Start(filePath.Replace(Path.GetExtension(filePath), "") + WordExtention);
            }
            else
            {
                MessageBox.Show("Dosya Seçilmedi || Not Selected File");
            }

        }


        /// <summary>
        /// Word Kurulumu, Kurulu ise Uzantıyı Ayarla
        /// </summary>
        /// <returns></returns>
        static bool OfficeControl()
        {
            // EsenClassLib.VersionNumberConvertName OfficeVersionNameStock = new EsenClassLib.VersionNumberConvertName();        
            // MessageBox.Show(EsenClassLib.GetOfficeVersion.GetVersionName(EsenClassLib.OfficeComponent.Word, OfficeVersionNameStock));

            bool IsThereTheWord = false;
            int WordVersionNumber = EsenClassLib.GetOfficeVersion.GetVersionNumber(EsenClassLib.OfficeComponent.Word);

            if (WordVersionNumber == 0)
            {
                MessageBox.Show("No Microsoft Word on Your System");
                IsThereTheWord = false;
            }
            else
            {
                if (WordVersionNumber > 11) WordExtention = ".docx";
                else WordExtention = ".doc";

                IsThereTheWord = true;
            }

            return IsThereTheWord;
        }
    }
}
