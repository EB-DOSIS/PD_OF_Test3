//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//+ Auswertung der Offenfeld EPIDS  für TB2258  VB4434  CL4160
//+ nutzt TestPatient 772016_OF_An_2258 etc
//+ Plan OF_2258_2020_II etc
//+ Feld OF_2258_X6 etc
//+ nutzt Excel Dateien PD_OF_Test3_2258.xlsx etc
//+ created by Eyck Blank
//+ assisted by Maximilian Grohmann
//+ 13.10.2020 FindFirst() Problem gelöst
//+ 15.10.2020 CL4160 eingebunden
//+ 16.11.2020 mit Marian Grafik zum Laufen gebracht
//+ 17.11.2020 mit Marian Grafiken und Layout optimiert
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



//using System;
//using System.Collections.Generic;
using System.IO;
//using System.Linq;
//using VMS.TPS.Common.Model.API;

using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using VMS.CA.Scripting;
using System.Drawing;
using System.Reflection; //Versions Attributes
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
//using System.Threading;
using VMS.DV.PD.Scripting;
using System.Windows.Media;



// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
// [assembly: AssemblyVersion("1.0.0.1")]
// [assembly: AssemblyFileVersion("1.0.0.1")]
// [assembly: AssemblyInformationalVersion("1.01")]


namespace VMS.DV.PD.Scripting 
{

    public class Script
    {

        public Script()
        {
        }
        const string SCRIPT_NAME = "PD_OF_Test3 Script";

        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {
            // TODO : Add here your code that is called when the script is launched from Portal Dosimetry

            String Maschine2258 = "Q:/TrueBeam2258/QA_xls/PD_OF_Test3_2258.xlsx";
            String Maschine4434 = "Q:/VitalBeam4434/QA_xls/PD_OF_Test3_4434.xlsx";
            String Maschine4160 = "Q:/Clinac4160/QA_xls/PD_OF_Test3_4160.xlsx";

            int pdBeamCount = 0;
            PDPlanSetup planID = context.PDPlanSetup;
            PDBeam fld = null;                       // Folgende Deklarations sequenz nur zur 
            foreach(PDBeam fld1 in planID.Beams)     // Bestimmen der Anzahl der Felder
            {
                pdBeamCount = fld1.PortalDoseImages.Count();
                fld = fld1;
                
            }
            int jPDbeam = 0;
            for (int i = 0; i < pdBeamCount; i++)
            {
                PortalDoseImage imgm = fld.PortalDoseImages[i];
                // ....
                Frame Fm = imgm.Image.FramesRT.LastOrDefault() ;               // .FirstOrDefault();
                string status = Fm.Image.ImageStatus.ToString();
                if (status == "Reviewed")
                {
                    jPDbeam = i;
             
                    int xsizem = Fm.XSize; int ysizem = Fm.YSize;
                    ushort[,] pixelsm = new ushort[xsizem, ysizem];
                    double resXm = Fm.XRes;  double resYm = Fm.YRes;
                    Fm.GetVoxels(0, pixelsm);

                    String Datum = Fm.CreationDateTime.Value.ToString("dd.MM.yyyy");
                    String Zeit = Fm.CreationDateTime.Value.ToString("HH:mm:ss");
                    String FName = fld.Id;
                
                    DoseImage imgp = fld.PredictedDoseImage;
                    Frame Fp = imgp.Image.FramesRT.LastOrDefault();
                    int xsizep = Fp.XSize; int ysizep = Fp.YSize;
                    ushort[,] pixelsp = new ushort[xsizep, ysizep];
                    double resXp = Fp.XRes;  double resYp = Fp.YRes;
                    Fp.GetVoxels(0, pixelsp);

                    int seitl   = 380;
                    int laengst = 280;
                    String  Name1;
                                        
                    // wegen untersch Kassettengrößen von Clinac und TrueBeam
                    // Verkomlizierung da bei 4160 die Felder Feld1 heissen
                    if (FName == "OF_4160_X6")
                    {
                        seitl   = 300;
                        laengst = 280; 
                       
                    }

                    else if (FName == "OF_4160_X15")
                    {
                        seitl   = 300;
                        laengst = 280; 
                       
                    }
                    else 
                    {
                        seitl   = 410;
                        laengst = 410; 
                       
                    }
                   
                    int CenterXm = Convert.ToInt32(Math.Ceiling(xsizem / 2.0));
                    int CenterYm = Convert.ToInt32(Math.Ceiling(ysizem / 2.0));

                    double CenterWertm = Fm.VoxelToDisplayValue(pixelsm[CenterXm, CenterYm]);
                    double ObenWertm = Fm.VoxelToDisplayValue(pixelsm[CenterXm, CenterYm - laengst]);
                    double LinksWertm = Fm.VoxelToDisplayValue(pixelsm[CenterXm - seitl, CenterYm]);
                    double RechtsWertm = Fm.VoxelToDisplayValue(pixelsm[CenterXm + seitl, CenterYm]);
                    double UntenWertm = Fm.VoxelToDisplayValue(pixelsm[CenterXm, CenterYm + laengst]);
                    // Chart vorbereiten
                    double[] ProfilX_m = new double[xsizem];
                    double[] ProfilXm = new double[xsizem];
                    for (int ii = 0; ii<xsizem;ii++)
                    {
                        ProfilX_m[ii] =  Fm.VoxelToDisplayValue(pixelsm[ii, CenterYm]);
                        ProfilXm[ii] =  ii*resXm - xsizem*resXm/2;
                    }
                    double[] ProfilY_m = new double[ysizem];
                    double[] ProfilYm = new double[ysizem];
                    for (int jj = 0; jj<ysizem;jj++)
                    {
                        ProfilY_m[jj] =  Fm.VoxelToDisplayValue(pixelsm[CenterXm, jj]);
                        ProfilYm[jj] =  jj*resYm - ysizem*resYm/2;
                    }


                    int CenterXp = Convert.ToInt32(Math.Ceiling(xsizep / 2.0));
                    int CenterYp = Convert.ToInt32(Math.Ceiling(ysizep / 2.0));

                    double CenterWertp = Fp.VoxelToDisplayValue(pixelsp[CenterXp, CenterYp]);
                    double ObenWertp = Fp.VoxelToDisplayValue(pixelsp[CenterXp, CenterYp - laengst]);
                    double LinksWertp = Fp.VoxelToDisplayValue(pixelsp[CenterXp - seitl, CenterYp]);
                    double RechtsWertp = Fp.VoxelToDisplayValue(pixelsp[CenterXp + seitl, CenterYp]);
                    double UntenWertp = Fp.VoxelToDisplayValue(pixelsp[CenterXp, CenterYp + laengst]);
                    // Chart vorbereiten
                    double[] ProfilX_p = new double[xsizep];
                    double[] ProfilXp = new double[xsizep];
                    for (int iii = 0; iii<xsizep;iii++)
                    {
                        ProfilX_p[iii] =  Fp.VoxelToDisplayValue(pixelsp[iii, CenterYp]);
                        ProfilXp[iii] =  iii*resXp - xsizep*resXp/2;       //* xsizem/xsizep;;
                    }
                    double[] ProfilY_p = new double[ysizep];
                    double[] ProfilYp = new double[ysizep];
                    for (int jjj = 0; jjj<ysizep;jjj++)
                    {
                        ProfilY_p[jjj] =  Fp.VoxelToDisplayValue(pixelsp[CenterXp, jjj]);
                        ProfilYp[jjj] =  jjj*resYp - ysizep*resYp/2;      //* ysizem/ysizep;
                    }

                    double dosisFehler = (CenterWertm - CenterWertp) / CenterWertp * 100;
                    double SymmetryTrans = (LinksWertm - RechtsWertm) / (LinksWertm + RechtsWertm) * 2 * 100;
                    double SymmetryLong = (ObenWertm - UntenWertm) / (ObenWertm + UntenWertm) * 2 * 100;
                    double HomFehler = (LinksWertm + RechtsWertm + ObenWertm + UntenWertm - 4 * CenterWertm) / (LinksWertm + RechtsWertm + ObenWertm + UntenWertm) * 100;
                   
                    var createHTML = HTMLBuilder.StaticText(Datum, Zeit, FName, ProfilXm, ProfilX_m, ProfilXp, ProfilX_p, ProfilYm, ProfilY_m, ProfilYp, ProfilY_p, 
                        dosisFehler.ToString("F2"), SymmetryTrans.ToString("F2"), SymmetryLong.ToString("F2"), HomFehler.ToString("F2"));
                    var runner = new HTMLRunner(createHTML);
                    runner.Launch("Test");
                    // chart ende


                    string messageObenm = string.Format("              " + "\t" + ObenWertm.ToString("F3") + "\n\r");
                    string messageMittem = string.Format(LinksWertm.ToString("F3") + "\t" + CenterWertm.ToString("F3") + "\t" + RechtsWertm.ToString("F3") + "\n\r");
                    string messageUntenm = string.Format("              " + "\t" + UntenWertm.ToString("F3"));

                    string messageObenp = string.Format("              " + "\t" + ObenWertp.ToString("F3") + "\n\r");
                    string messageMittep = string.Format(LinksWertp.ToString("F3") + "\t" + CenterWertp.ToString("F3") + "\t" + RechtsWertp.ToString("F3") + "\n\r");
                    string messageUntenp = string.Format("              " + "\t" + UntenWertp.ToString("F3"));

                    MessageBoxResult result = MessageBox.Show(
                        "Status " + " " + status + " " + "\r\n" + jPDbeam + " " + Datum + " " + Zeit + "\r\n" + 
                        "Portal Image" + "\r\n" + "--------------" + "\r\n"
                        + messageObenm + "\r\n" + messageMittem + "\r\n" + messageUntenm
                        + "\r\n" + "\r\n" + "\r\n"
                        + "Predicted Image" + "\r\n" + "-----------------" + "\r\n"
                        + messageObenp + "\r\n" + messageMittep + "\r\n" + messageUntenp + "\r\n" + "\r\n"
                        + "Evaluation" + "\r\n" + "------------" + "\r\n" + "\r\n"
                        + "Dose error " + "\t" + dosisFehler.ToString("F3") + " %" + "\n\r"
                        + "Sym-trans  " + "\t" + SymmetryTrans.ToString("F3") + " %" + "\n\r"
                        + "Sym-long   " + "\t" + SymmetryLong.ToString("F3") + " %" + "\n\r"
                        + "Hom-error  " + "\t" + HomFehler.ToString("F3") + " %"
                        , SCRIPT_NAME, MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation);

                   
                    switch (FName)
                    {
                        case "OF_2258_X6":
                            // in Excel X6 schreiben
                            UpdateExcel(Maschine2258, Datum, Zeit, "2258_X6", CenterWertp.ToString("F3"), ObenWertp.ToString("F3"), UntenWertp.ToString("F3"), LinksWertp.ToString("F3"), RechtsWertp.ToString("F3"),
                                CenterWertm.ToString("F3"), ObenWertm.ToString("F3"), UntenWertm.ToString("F3"), LinksWertm.ToString("F3"), RechtsWertm.ToString("F3"),
                                dosisFehler.ToString("F3"), SymmetryTrans.ToString("F3"), SymmetryLong.ToString("F3"), HomFehler.ToString("F3"));
                            break;
                        case "OF_2258_X15":
                            // in Excel X15 schreiben
                            UpdateExcel(Maschine2258, Datum, Zeit, "2258_X15", CenterWertp.ToString("F3"), ObenWertp.ToString("F3"), UntenWertp.ToString("F3"), LinksWertp.ToString("F3"), RechtsWertp.ToString("F3"),
                                CenterWertm.ToString("F3"), ObenWertm.ToString("F3"), UntenWertm.ToString("F3"), LinksWertm.ToString("F3"), RechtsWertm.ToString("F3"),
                                dosisFehler.ToString("F3"), SymmetryTrans.ToString("F3"), SymmetryLong.ToString("F3"), HomFehler.ToString("F3"));
                            break;
                                
                        case "OF_4434_X6":
                            // in Excel X6 schreiben
                            UpdateExcel(Maschine4434, Datum, Zeit, "4434_X6", CenterWertp.ToString("F3"), ObenWertp.ToString("F3"), UntenWertp.ToString("F3"), LinksWertp.ToString("F3"), RechtsWertp.ToString("F3"),
                                CenterWertm.ToString("F3"), ObenWertm.ToString("F3"), UntenWertm.ToString("F3"), LinksWertm.ToString("F3"), RechtsWertm.ToString("F3"),
                                dosisFehler.ToString("F3"), SymmetryTrans.ToString("F3"), SymmetryLong.ToString("F3"), HomFehler.ToString("F3"));
                            break;
                        case "OF_4434_X15":
                            // in Excel X15 schreiben
                            UpdateExcel(Maschine4434, Datum, Zeit, "4434_X15", CenterWertp.ToString("F3"), ObenWertp.ToString("F3"), UntenWertp.ToString("F3"), LinksWertp.ToString("F3"), RechtsWertp.ToString("F3"),
                                CenterWertm.ToString("F3"), ObenWertm.ToString("F3"), UntenWertm.ToString("F3"), LinksWertm.ToString("F3"), RechtsWertm.ToString("F3"),
                                dosisFehler.ToString("F3"), SymmetryTrans.ToString("F3"), SymmetryLong.ToString("F3"), HomFehler.ToString("F3"));
                            break;

                        case "OF_4160_X6":
                            // in Excel X6 schreiben
                            UpdateExcel(Maschine4160, Datum, Zeit, "4160_X6", CenterWertp.ToString("F3"), ObenWertp.ToString("F3"), UntenWertp.ToString("F3"), LinksWertp.ToString("F3"), RechtsWertp.ToString("F3"),
                                CenterWertm.ToString("F3"), ObenWertm.ToString("F3"), UntenWertm.ToString("F3"), LinksWertm.ToString("F3"), RechtsWertm.ToString("F3"),
                                dosisFehler.ToString("F3"), SymmetryTrans.ToString("F3"), SymmetryLong.ToString("F3"), HomFehler.ToString("F3"));
                            break;
                        case "OF_4160_X15":
                            // in Excel X15 schreiben
                            UpdateExcel(Maschine4160, Datum, Zeit, "4160_X15", CenterWertp.ToString("F3"), ObenWertp.ToString("F3"), UntenWertp.ToString("F3"), LinksWertp.ToString("F3"), RechtsWertp.ToString("F3"),
                                CenterWertm.ToString("F3"), ObenWertm.ToString("F3"), UntenWertm.ToString("F3"), LinksWertm.ToString("F3"), RechtsWertm.ToString("F3"),
                                dosisFehler.ToString("F3"), SymmetryTrans.ToString("F3"), SymmetryLong.ToString("F3"), HomFehler.ToString("F3"));
                            break;



                        default:
                            // Nichts
                             MessageBox.Show("Maschine unbekannt");

                            break;

                    }
                    
                    
                }
            }

        }

        public class HTMLRunner // Create HTML File 
        {
            public string Text { get; set; }
            public string TempFolder { get; set; }
            public HTMLRunner(string text)
            {
                TempFolder = Path.GetTempPath();
                Text = text.ToString();
            }
            public void Launch(string title)
            {
                var fileName = Path.Combine(TempFolder, title + ".html");
                File.WriteAllText(fileName, Text);
                System.Diagnostics.Process.Start(fileName);
            }
        }
        public class HTMLBuilder //Generate the text for the plot to be passed to HTMLRunner class 
        {
            public static string StaticText(string Datum, string Zeit, string maschine, IEnumerable<double> xsXm, IEnumerable<double> ysXm, IEnumerable<double> xsXp, IEnumerable<double> ysXp,
                IEnumerable<double> xsYm, IEnumerable<double> ysYm, IEnumerable<double> xsYp, IEnumerable<double> ysYp, string doseFehler, string symFehlerTrans, string symFehlerLong, string homFehler)
            // public static string StaticText(IEnumerable<double> xs, IEnumerable<double> ys, IEnumerable<double> zs)
            {
                var preX = @" 
                <!DOCTYPE html> 
                <html lang='en'> 
  
                    <head> 
                        <meta charset=utf-8> 
                        <title>Plot.ly 2D</title> 
                        <script src='Q:/ESAPI/Projects/PD_OF_Test3/test.js'></script> 
                    </head> 
  
                    <body> 
                        <form>
                            <button type='button' id='print' onclick='window.print()' style='float: right;'>Drucken</button>
                        </form>
                        <h1>PD_OF_Test3</h1> 
                        <h2>" + maschine + @", &nbsp;&nbsp; " + Datum + @", &nbsp; " + Zeit + @"</h2> 
                        <h3>Crossplane Profile</h3> 
                        <div id='chartX'> </div> 

                        <h3>Inplane Profile</h3> 
                        <div id='chartY'> </div> 

                        <h2> </h2>
                        <h2> </h2>
                        <h2>Fehlerwerte:</h2> 
                        <h2>--------------------------------------</h2> 
                        <h3>Dosisfehler     :" + doseFehler + @"</h3> 
                        <h3> </h3> 
                        <h3>Sym Fehler Trans:" + symFehlerTrans + @"</h3> 
                        <h3> </h3> 
                        <h3>Sym Fehler Long  :" + symFehlerLong + @"</h3> 
                        <h3> </h3> 
                        <h3>Hom Fehler       :" + homFehler + @"</h3> 

                    </body> 
            
                    <script> 
                        var chartX = document.getElementById('chartX'); 
                        var layoutX = {
                            xaxis: {title: 'Profil mm'},  
                            yaxis: {title: 'Intensität'},
                            autosize: false,
                            width: 1100,
                            height: 500,
                            title: 'Crossplane Querprofil',
                            margin: 
                            { 
                                t: 0 
                            }
                        };
                        var dataXm = 
                        { 
                            ";
                            var xXm = "x:[" + string.Join(",", xsXm) + "],";
                            var yXm = "y:[" + string.Join(",", ysXm) + "],";    
                            var posYXm = @" 
                            marker: {color:'red',size: 12,  line:{color:'red',width: 0.5}},
                            mode: 'lines',
                            name: 'Measured',
                            type: 'scatter2d' 
                        }; 
               
                        var dataXp = 
                        { 
                            ";
                            var xXp = "x:[" + string.Join(",", xsXp) + "],";
                            var yXp = "y:[" + string.Join(",", ysXp) + "],";
                            var posYXp = @" 
                            marker: {color:'blue',size: 12,  line:{color:'blue',width: 0.5}},
                            mode: 'lines',
                            name: 'Predicted',
                            type: 'scatter2d' 
                        }; 
                                               
                        var data = [dataXm, dataXp];
             
                        Plotly.newPlot(chartX, data, layoutX); 


                        var chartY = document.getElementById('chartY'); 
                        var layoutY =  {
                            xaxis: {title: 'Profil mm'},  
                            yaxis: {title: 'Intensität'},
                            autosize: false,
                            width: 1100,
                            height: 500,
                            title: 'Inplane Querprofil',
                            margin: 
                            { 
                                t: 0 
                            }
                        };
                        var dataYm = 
                        { 
                            ";

                            //var x = "x:[1,2,3,4],";
                            var xYm = "x:[" + string.Join(",", xsYm) + "],";
                            var yYm = "y:[" + string.Join(",", ysYm) + "],";  
                            var posYYm = @" 
                            marker: {color:'red',size: 12,  line:{color:'red',width: 0.5}},
                            mode: 'lines',
                            name: 'Measured',
                            type: 'scatter2d' 
                        }; 
               
                        var dataYp = 
                        { 
                            ";
                            var xYp = "x:[" + string.Join(",", xsYp) + "],";
                            var yYp = "y:[" + string.Join(",", ysYp) + "],";
                            var posYYp = @" 
                            marker: {color:'blue',size: 12,  line:{color:'blue',width: 0.5}},
                            mode: 'lines',
                            name: 'Predicted',
                            type: 'scatter2d' 
                        }; 
                                               
                        var data = [dataYm, dataYp];
                              
                        Plotly.newPlot(chartY, data, layoutY); 

                    </script> 

                </html> 
                ";
                return preX + xXm + yXm + posYXm + xXp + yXp + posYXp   + xYm + yYm + posYYm + xYp + yYp + posYYp ;
                //return preX + x + y + z + posY;
            }
        }


        private void UpdateExcel(string Maschine, string Datum, string Zeit, string sheetName, string data1, string data2, string data3, string data4, string data5,
            string data6, string data7, string data8, string data9, string data10,
            string data11, string data12, string data13, string data14)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;

            //DateTime localDate = DateTime.Now;
            //String Datum = localDate.ToString("dd.MM.yyyy");
            //String Zeit = localDate.ToString("HH:mm:ss");

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(Maschine);
                oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                // auslesen der letzten Reihenzahl
                string sReihe = oSheet.Cells[3, 2].Value == null ? "-" : oSheet.Cells[3, 2].Value.ToString();
                int iReihe = Convert.ToInt32(sReihe);
                iReihe = iReihe + 1;

                // Datum Zeit
                oSheet.Cells[iReihe, 1] = Datum;
                oSheet.Cells[iReihe, 2] = Zeit;
                // Predicted
                oSheet.Cells[iReihe, 3] = data1;
                oSheet.Cells[iReihe, 4] = data2;
                oSheet.Cells[iReihe, 5] = data3;
                oSheet.Cells[iReihe, 6] = data4;
                oSheet.Cells[iReihe, 7] = data5;
                // Measures
                oSheet.Cells[iReihe, 9] = data6;
                oSheet.Cells[iReihe, 10] = data7;
                oSheet.Cells[iReihe, 11] = data8;
                oSheet.Cells[iReihe, 12] = data9;
                oSheet.Cells[iReihe, 13] = data10;
                // Evaluated
                oSheet.Cells[iReihe, 15] = data11;
                oSheet.Cells[iReihe, 16] = data12;
                oSheet.Cells[iReihe, 17] = data13;
                oSheet.Cells[iReihe, 18] = data14;

                // Schreiben neuer Zeilenzahl
                sReihe = Convert.ToString(iReihe);
                oSheet.Cells[3, 2] = sReihe;

                oWB.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, null, null);
                    oXL.Quit();
                }

                //oWB.Close(false);

            }

            MessageBox.Show("Done");
        }
    }
}
