using Aspose.OMR;
using Aspose.OMR.Api;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;

namespace AsposeOMRWorkflow
{
    class Program
    {
        static void Main(string[] args)
        {
            License omrLicense = new License();
            omrLicense.SetLicense("Aspose.Total.lic");

            var menuSelection = "";

            while (menuSelection != "X" && menuSelection != "x")
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Generate Template - T");
                Console.WriteLine("Print Questionnaire - P");
                Console.WriteLine("Scan Flatbed Docs - S");               
                Console.WriteLine("Evaluate Scanned Docs - E");
                Console.WriteLine("Exit - X");
                Console.ForegroundColor = ConsoleColor.White;

                menuSelection = Console.ReadLine();
                
                switch (menuSelection)
                {
                    case "T":
                    case "t":
                        GenerateFormTemplateAndImage();
                        break;
                    case "P":
                    case "p":
                        PrintFormImage();
                        break;
                    case "S":
                    case "s":
                        ScanForms(4);
                        break;
                    case "E":
                    case "e":
                        ProcessScansAndExportData();
                        break;
                    case "X":
                    case "x":
                    default:
                        break;
                }
                Console.BackgroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("Finished");
                Console.BackgroundColor = ConsoleColor.Black;
            }
        }

        static void GenerateFormTemplateAndImage()
        {
            //fully qualified path to the form markup file
            var formMarkupFilePath = AppDomain.CurrentDomain.BaseDirectory + "QuestionnaireMarkup.txt";

            //initialize an instance of the Aspose OMR engine
            var omrEngine = new OmrEngine();

            //use the Aspose OMR engine to generate the template and image from the markup file
            var result = omrEngine.GenerateTemplate(formMarkupFilePath);
            if (result.ErrorCode != 0)
            {
                Console.WriteLine($"ERROR: {result.ErrorCode} - {result.ErrorMessage}");
            }
            else
            {
                //save the files as OmrOutput.omr for the template, and OmrOutput.png for the form image
                result.Save("", "OmrOutput");
            }
        }

        static void PrintFormImage()
        {
            //fully qualified path to the form image
            var questionnaireImagePath = AppDomain.CurrentDomain.BaseDirectory + "OmrOutput.png";
            
            //we will be using the default printer to print the form
            PrintDocument pd = new PrintDocument();

            //event handler fired when a print job is requested
            pd.PrintPage += (sender, args2) =>
            {
                //obtain image to print
                Image i = Image.FromFile(questionnaireImagePath);

                //obtain paper margins
                Rectangle m = args2.MarginBounds;

                //ensure we scale the image proportionally on the page
                if ((double)i.Width / (double)i.Height > (double)m.Width / (double)m.Height) // image is wider
                {   
                    m.Height = (int)((double)i.Height / (double)i.Width * (double)m.Width);
                }
                else
                {
                    m.Width = (int)((double)i.Width / (double)i.Height * (double)m.Height);
                }

                //render the image 
                args2.Graphics.DrawImage(i, m);
            };
            
            //prints one copy of the form image
            pd.Print();
        }

        static void ScanForms(int numberOfForms)
        {
            //fully qualified output path where the scanned images of the forms will be located.
            var scannedImagePath = AppDomain.CurrentDomain.BaseDirectory + @"Scans\";
                        
            //obtain the first defined scanner found on the computer 
            var scanner = WIAScanner.GetDevices().FirstOrDefault();

            //will store scanned images of the forms
            List<Image> images = new List<Image>();
            int idx = 0;
            
            //scans multiple forms and stores the images in the images list
            images = WIAScanner.Scan(scanner.DeviceID, numberOfForms, WIAScanQuality.Final, WIAPageSize.Letter, DocumentSource.Feeder);

            if (images.Count > 0)
            {
                //create the Scans directory if it doesn't already exist
                Directory.CreateDirectory(scannedImagePath);
            }

            //save each image obtained from the scanner into the Scans folder.
            foreach (var img in images)
            {
                var fileName = "img" + ++idx;
                img.Save(scannedImagePath + fileName + ".jpg", ImageFormat.Jpeg);
            }
        }

        static void ProcessScansAndExportData()
        {
            //fully qualified output path where the scanned images of the forms will be located.
            var scannedImagePath = AppDomain.CurrentDomain.BaseDirectory + @"Scans\";

            ///fully qualified path to the form template
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + "OmrOutput.omr";

            //initialize an instance of the Aspose OMR engine
            var omrEngine = new OmrEngine();

            //retrieve all images from the Scans folder
            var dirInfo = new DirectoryInfo(scannedImagePath);
            var files = dirInfo.GetFiles("*.jpg");

            //use the omrEngine to create an instance of the template processor based on the generated template
            var templateProcessor2 = omrEngine.GetTemplateProcessor(templatePath);

            foreach (var file in files)
            {
                //use the template processor to extract form data from the form image
                string jsonResults = templateProcessor2.RecognizeImage(file.FullName, 28).GetJson();

                //save the extracted data in a json file
                File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + Path.GetFileNameWithoutExtension(file.FullName) + "_scan_results.json", jsonResults);
            }
        }
    }
}
