using System;
//using Microsoft.Win32;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Linq;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Model.Model3D;
using TFlex.Drawing;
using TFlex.Command;

using TFlex.Model.Data;
using TFlex.Model.Data.ProductStructure;

namespace FRAGMENTSTREE_PLG
{
    public class CommandManager
    {
        public int Level, exportSet;
        public string BasePath;
        public string pathSTPRU, pathDXFRU, pathPDFRU;
        public bool exportDXF, exportSTEP, exportPDF;
        DateTime t1, t2; //поля класса формы
        private readonly System.Diagnostics.Stopwatch uptime = new System.Diagnostics.Stopwatch();

        private Page GetPageDXF(Document doc, string namePage)
        {
            foreach (Page page in doc.GetPages())
            {
                if (page.Name == namePage)
                {
                    return page;
                }
            }
            return null;
        }

        private Page GetPageDXF2(Document doc, PageType pageType)
        {
            foreach (Page page in doc.GetPages())
            {
                if (page.PageType == pageType)
                {
                    return page;
                }
            }
            return null;
        }

        private List<Page> GetPagesPDF(Document doc, PageType pageType)
        {
            List<Page> pagesList = new List<Page>();
            foreach (Page page in doc.GetPages())
            {
                if (page.PageType == pageType)
                {
                    pagesList.Add(page);
                }
            }
            if (pagesList.Count != 0) return pagesList;
            else return null;
        }

        public void OK(Document doc, ATTR_COM par)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            //System.Threading.Thread.Sleep(10000);

            if (par.pDXF == 1) exportDXF = true;
            if (par.pSTEP == 1) exportSTEP = true;
            if (par.pPDF == 1) exportPDF = true;
            if (!exportDXF && !exportSTEP && !exportPDF) exportSet = -1;
            else exportSet = 1;

            /*if (par.pDXF == 1 && par.pSTEP == 1) exportSet = 2;
            else if (par.pDXF == 1 && par.pSTEP == 0) exportSet = 1;
            else if (par.pDXF == 0 && par.pSTEP == 1) exportSet = 0;
            else exportSet = -1;*/
            if (exportSet != -1)
            {
                string logname = doc.FileName;
                logname = logname.Replace(".grb", ".log");
                using (StreamWriter sw = new StreamWriter(logname))
                {
                    if (doc == null)
                        return;

                    try
                    {
                        string oboz = "-";
                        string naim = "-";
                        Variable voboz = doc.FindVariable("$Обозначение");
                        Variable vnaim = doc.FindVariable("$Наименование");

                        if (voboz != null)
                        {
                            oboz = voboz.TextValue;
                        }
                        else oboz = "";
                        if (naim != null)
                        {
                            naim = vnaim.TextValue;
                        }
                        else naim = "";

                        FileInfo parFile = new FileInfo(doc.FileName);
                        DirectoryInfo parDir = new DirectoryInfo(parFile.DirectoryName);
                        BasePath = parDir.FullName;
                        string subpathSTPRU = @"STP";
                        string subpathDXFRU = @"DXF";
                        string subpathPDFRU = @"PDF";
                        pathSTPRU = pathDXFRU = pathPDFRU = parFile.DirectoryName;
                        DirectoryInfo dirInfo = new DirectoryInfo(pathSTPRU);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        if (exportSTEP)
                            dirInfo.CreateSubdirectory(subpathSTPRU);
//                        pathSTPRU = $"{pathSTPRU}\\{subpathSTPRU}";
                        pathSTPRU = String.Format("{0}\\{1}",pathSTPRU, subpathSTPRU);
                        if (exportDXF)
                            dirInfo.CreateSubdirectory(subpathDXFRU);
//                        pathDXFRU = $"{pathDXFRU}\\{subpathDXFRU}";
                        pathDXFRU = String.Format("{0}\\{1}",pathDXFRU, subpathDXFRU);
                        if (exportPDF)
                            dirInfo.CreateSubdirectory(subpathPDFRU);
//                        pathPDFRU = $"{pathPDFRU}\\{subpathPDFRU}";
                        pathPDFRU = String.Format("{0}\\{1}",pathPDFRU, subpathPDFRU);
                        string file_name = doc.FileName;
                        Level = 0;
                        GetFragmentData(doc, sw, file_name, doc.FilePath, oboz, naim, Level);
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message, "Ошибка", System.Windows.Forms.MessageBoxButtons.OK);
                        TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";
                    }
                }
            }

            TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";

            doc.Selection.DeselectAll();
            /*switch (exportSet)
            {
                case 2:
                    System.Windows.Forms.MessageBox.Show("Экспорт STEP и DXF завершён", "Разбор по фрагментам(сборка)");
                    break;
                case 1:
                    System.Windows.Forms.MessageBox.Show("Экспорт STEP завершён", "Разбор по фрагментам(сборка)");
                    break;
                case 0:
                    System.Windows.Forms.MessageBox.Show("Экспорт DXF завершён", "Разбор по фрагментам(сборка)");
                    break;
                case -1:
                    System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено(не выбраны варианты)", "Разбор по фрагментам(сборка)");
                    break;
            }*/
            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено\nRunTime: " + elapsedTime, "Разбор по фрагментам(сборка)");
            else System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено(не выбраны параметры)", "Разбор по фрагментам(сборка)");
            /*if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено");
            else if (exportSet == -1)
                System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено (не выбраны варианты)");*/

            return;
        }

        public void OK2(Document doc, ATTR_COM par)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            //System.Threading.Thread.Sleep(10000);
            if (par.pDXF == 1) exportDXF = true;
            if (par.pSTEP == 1) exportSTEP = true;
            if (par.pPDF == 1) exportPDF = true;
            if (!exportDXF && !exportSTEP && !exportPDF) exportSet = -1;

            /*if (par.pDXF == 1 && par.pSTEP == 1) exportSet = 2;
            else if (par.pDXF == 1 && par.pSTEP == 0) exportSet = 1;
            else if (par.pDXF == 0 && par.pSTEP == 1) exportSet = 0;
            else exportSet = -1;*/
            if (exportSet != -1)
            {
                string logname = doc.FileName;
                logname = logname.Replace(".grb", ".log");
                using (StreamWriter sw = new StreamWriter(logname))
                {
                    if (doc == null)                        return;

                    try
                    {
                        string oboz = "-";
                        string naim = "-";
                        Variable voboz = doc.FindVariable("$Обозначение");
                        Variable vnaim = doc.FindVariable("$Наименование");

                        if (voboz != null)
                        {
                            oboz = voboz.TextValue;
                        }
                        else oboz = "";
                        if (naim != null)
                        {
                            naim = vnaim.TextValue;
                        }
                        else naim = "";

                        FileInfo parFile = new FileInfo(doc.FileName);
                        DirectoryInfo parDir = new DirectoryInfo(parFile.DirectoryName);
                        BasePath = parDir.FullName;
                        string subpathSTPRU = @"STP";
                        string subpathDXFRU = @"DXF";
                        string subpathPDFRU = @"PDF";
                        pathSTPRU = pathDXFRU = pathPDFRU = parFile.DirectoryName;
                        DirectoryInfo dirInfo = new DirectoryInfo(pathSTPRU);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        if (exportSTEP)
                            dirInfo.CreateSubdirectory(subpathSTPRU);
//                        pathSTPRU = $"{pathSTPRU}\\{subpathSTPRU}";
                 pathSTPRU = "{1}\\{2}, pathSTPRU, subpathSTPRU";
                        if (exportDXF)
                            dirInfo.CreateSubdirectory(subpathDXFRU);
//                        pathDXFRU = $"{pathDXFRU}\\{subpathDXFRU}";
                 pathDXFRU = "{1}\\{2}, pathDXFRU, subpathDXFRU";
                        if (exportPDF)
                            dirInfo.CreateSubdirectory(subpathPDFRU);
//                        pathPDFRU = $"{pathPDFRU}\\{subpathPDFRU}";
                 pathPDFRU = "{1}\\{2}, pathPDFRU, subpathPDFRU";
                        string file_name = doc.FileName;
                        Level = 0;
                        Set(doc, sw, file_name, doc.FilePath, oboz, naim, Level);
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message, "Ошибка", System.Windows.Forms.MessageBoxButtons.OK);
                        TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";
                    }
                }
            }


            TFlex.Application.ActiveMainWindow.StatusBar.Prompt = "";

            doc.Selection.DeselectAll();
            /* switch (exportSet)
             {
                 case 2:
                     System.Windows.Forms.MessageBox.Show("Генерирование STEP и DXF завершено", "Генерирование по конфигурациям(деталь)");
                     break;
                 case 1:
                     System.Windows.Forms.MessageBox.Show("Генерирование STEP завершено", "Генерирование по конфигурациям(деталь)");
                     break;
                 case 0:
                     System.Windows.Forms.MessageBox.Show("Генерирование DXF завершено", "Генерирование по конфигурациям(деталь)");
                     break;
                 case -1:
                     System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено(не выбраны варианты)", "Генерирование по конфигурациям(деталь)");
                     break;
             }*/
            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            if (exportSet != -1)
                System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено\nRunTime: " + elapsedTime, "Генерирование по конфигурациям(деталь)");
            else System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено(не выбраны параметры)", "Генерирование по конфигурациям(деталь)");
            //System.Windows.Forms.MessageBox.Show("Выполнение плагина завершено");

            return;
        }

  private ProductStructure FindProductStructure(Document doc, string psName)
  {
string text = "";
bool empty = false;

   foreach (ProductStructure ps in doc.GetProductStructures())
    {
     if (ps.Name == psName)
      {
   ReadFragInfo(doc, ps, "1", ref text, ref empty);
       return ps;
      }
    }

   return null;
  }

  bool ReadFragInfo(Document doc, ProductStructure ps, string frName, ref string Labeltext, ref bool empty)
  {
   ParameterDescriptor param = null; // nalezeni spravneho sloupecku ze ktereho brat popis
   foreach (ParameterDescriptor paramcur in ps.GetScheme().Parameters)
    {
     if (paramcur.Name == "Раздел") param = paramcur;
    }

   empty = true;
   if (param == null) return false;

   foreach (RowElement elem in ps.GetAllRowElements())
    {
//elem[param.UID].ValueAsString 

    Fragment fr1 = elem.SourceFragmentFirstLevel;
     string Name = null;

     if (fr1 != null)
      {
       Fragment3D frag = fr1.FindAssociatedFragment();
       if (frag != null) Name = frag.Name;
      }

     if (Name == null) continue;

     if (Name == frName)
      {
       if (((string)elem[param].Value).Length > 0)
         empty = false;

       Labeltext = elem.Position.ValueAsString + " - " + elem[param].Value; // pozice a popis z produktove struktury

       return true;
      }
    }

   return false;
  }

        private void GetFragmentData(Document doc, StreamWriter sw, string name, string path, string oboz, string naim, int lev)
        {
            RegenerateOptions rg = new RegenerateOptions();
            rg.Full = true;
            doc.Regenerate(rg);
            string offset = "";
            for (int nn = 0; nn < lev; nn++)
                offset += "   ";


 FindProductStructure(doc, "Спецификация");
 


            string confName = doc.ModelConfigurations.GetConfigurationWithCurrentValues();
            if (((doc.FindVariable("$Наименование").TextValue != "") || (doc.FindVariable("$Обозначение").TextValue != ""))
                && (doc.ModelConfigurations.ConfigurationCount != 0))
            {
                if (exportSTEP)
                {
                    ExportToStep exportSTPRU = new ExportToStep(doc);
//                    string fileNameSTPRU = ($"{pathSTPRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.stp");
             string fileNameSTPRU = ("{0}\\{1}_{2}.stp, pathSTPRU, confName, doc.FindVariable(\"$Наименование\").TextValue");
                    if (!File.Exists(fileNameSTPRU))
                    {
                        sw.WriteLine(offset + "Имя: " + name);
                        sw.WriteLine(offset + "Путь: " + path);
                        sw.WriteLine(offset + "Наименование: " + naim);
                        sw.WriteLine(offset + "Обозначение: " + oboz);
                        sw.WriteLine(offset + "STEP Export: OK");
                        exportSTPRU.Export(fileNameSTPRU);
                    }
                }
                if (exportDXF)
                {
                    ExportToDXF exportDXFRU = new ExportToDXF(doc);
                    Page pgRU = GetPageDXF(doc, "Развертка");
                    Page pgRU2 = GetPageDXF(doc, "Unfolding");
                    //Page pgRU = GetPageDXF2(doc, PageType.Auxiliary);
                    if (pgRU != null || pgRU2 != null)
                    {
                        List<Page> pgDXFRU = new List<Page>();
                        pgDXFRU.Add(pgRU);
                        pgDXFRU.Add(pgRU2);
                        exportDXFRU.ExportPages = pgDXFRU;
//                        string fileNameDXFRU = ($"{pathDXFRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.dxf");
                        string fileNameDXFRU = ("{0}\\{1}_{2}.dxf, pathDXFRU, confName, doc.FindVariable(\"$Наименование\").TextValue");
                        if (!File.Exists(fileNameDXFRU))
                        {
                            sw.WriteLine(offset + "DXF Export: OK");
                            exportDXFRU.Export(fileNameDXFRU);
                        }
                    }
                }
                if (exportPDF)
                {
                    ExportToPDF exportPDFnormalconf = new ExportToPDF(doc);
                    List<Page> pgnormalconf = GetPagesPDF(doc, PageType.Normal);
                    if (pgnormalconf != null)
                    {
                        exportPDFnormalconf.ExportPages = pgnormalconf;
                        exportPDFnormalconf.OpenExportFile = false;
//                        string fileNamePDFRU = ($"{pathPDFRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.pdf");
                        string fileNamePDFRU = String.Format("{0}\\{1}_{2}.pdf", pathPDFRU, confName, doc.FindVariable("$Наименование").TextValue);
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDF Export: OK");
                            exportPDFnormalconf.Export(fileNamePDFRU);
                        }
                    }
                    ExportToPDF exportPDFBOMconf = new ExportToPDF(doc);
                    List<Page> pgPDFBOMconf = GetPagesPDF(doc, PageType.BillOfMaterials);
                    if (pgPDFBOMconf != null)
                    {
                        exportPDFBOMconf.ExportPages = pgPDFBOMconf;
                        exportPDFBOMconf.OpenExportFile = false;
//                        string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_СП.pdf");
                        string fileNamePDFRU = String.Format("{0}\\{1}_{2}_СП.pdf", pathPDFRU, doc.FindVariable("$Обозначение").TextValue, doc.FindVariable("$Наименование").TextValue);
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFBOM Export: OK");
                            exportPDFBOMconf.Export(fileNamePDFRU);
                        }
                    }
                }
            }
            else if ((doc.FindVariable("$Наименование").TextValue != "") || (doc.FindVariable("$Обозначение").TextValue != ""))
            {
                if (exportSTEP)
                {
                    ExportToStep exportSTPRU = new ExportToStep(doc);
//                    string fileNameSTPRU = ($"{pathSTPRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}.stp");
                    string fileNameSTPRU = ("{0}\\{1}_{2}.stp, pathSTPRU, doc.FindVariable(\"$Обозначение\").TextValue, doc.FindVariable(\"$Наименование\").TextValue");
                    if (!File.Exists(fileNameSTPRU))
                    {
                        sw.WriteLine(offset + "Имя: " + name);
                        sw.WriteLine(offset + "Путь: " + path);
                        sw.WriteLine(offset + "Наименование: " + naim);
                        sw.WriteLine(offset + "Обозначение: " + oboz);
                        sw.WriteLine(offset + "STEP Export: OK");
                        exportSTPRU.Export(fileNameSTPRU);
                    }
                }
                if (exportDXF)
                {
                    ExportToDXF exportDXFRU = new ExportToDXF(doc);
                    Page pgRUDXF = GetPageDXF(doc, "Развертка");
                    Page pgRUDXF2 = GetPageDXF(doc, "Unfolding");
                    //Page pgRUDXF = GetPageDXF2(doc, PageType.Auxiliary);
                    if (pgRUDXF != null || pgRUDXF2 != null)
                    {
                        List<Page> pgDXFRU = new List<Page>();
                        pgDXFRU.Add(pgRUDXF);
                        pgDXFRU.Add(pgRUDXF2);
                        exportDXFRU.ExportPages = pgDXFRU;
//                        string fileNameDXFRU = ($"{pathDXFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}.dxf");
                        string fileNameDXFRU = ("{0}\\{1}_{2}.dxf, pathDXFRU, doc.FindVariable(\"$Обозначение\").TextValue, doc.FindVariable(\"$Наименование\").TextValue");
                        if (!File.Exists(fileNameDXFRU))
                        {
                            sw.WriteLine(offset + "DXF Export: OK");
                            exportDXFRU.Export(fileNameDXFRU);
                        }
                    }
                }
                if (exportPDF)
                {
                    ExportToPDF exportPDFnormalRU = new ExportToPDF(doc);
                    List<Page> pgPDFnormalRU = GetPagesPDF(doc, PageType.Normal);
                    if (pgPDFnormalRU != null)
                    {
                        exportPDFnormalRU.ExportPages = pgPDFnormalRU;
                        exportPDFnormalRU.OpenExportFile = false;
//                        string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}.pdf");
                        string fileNamePDFRU = String.Format("{0}\\{1}_{2}.pdf", pathPDFRU, doc.FindVariable("$Обозначение").TextValue, doc.FindVariable("$Наименование").TextValue);
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFnormal Export: OK");
                            exportPDFnormalRU.Export(fileNamePDFRU);
                        }
                    }
                    ExportToPDF exportPDFBOMRU = new ExportToPDF(doc);
                    List<Page> pgPDFBOMRU = GetPagesPDF(doc, PageType.BillOfMaterials);
                    if (pgPDFBOMRU != null)
                    {
                        exportPDFBOMRU.ExportPages = pgPDFBOMRU;
                        exportPDFBOMRU.OpenExportFile = false;
//                        string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_СП.pdf");
                        string fileNamePDFRU = String.Format("{0}\\{1}_{2}_СП.pdf", pathPDFRU, doc.FindVariable("$Обозначение").TextValue, doc.FindVariable("$Наименование").TextValue);
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFBOM Export: OK");
                            exportPDFBOMRU.Export(fileNamePDFRU);
                        }
                    }
                }
            }
            int n_fr = doc.GetFragments3D().Count;
            foreach (Fragment3D frag in doc.GetFragments3D())
            {
                if (frag.Suppression.Suppress) continue;
                else
                {
                    Document docFR = null;
                    string obozF = "-";
                    string naimF = "-";
                    bool err = false;
                    string str_err = "";
                    string FRname = frag.FullFilePath;
                    if (lev > 0)
                    {
                        FRname = frag.FilePath;
                        FRname = TFlex.Application.FindPathName(FRname);
                    }

                    if (File.Exists(FRname))
                    {
                        Fragment.OpenPartOptions options = new Fragment.OpenPartOptions();
                        options.DontShowDocument = true;
                        options.QuietMode = true;
                        options.SubstituteGeometry = true;
                        options.SubstituteVariables = true;
                        options.SubstituteStatus = true;
                        docFR = frag.OpenPart(options);

                        if (docFR != null)
                        {
                            Variable voboz = docFR.FindVariable("$Обозначение");
                            Variable vnaim = docFR.FindVariable("$Наименование");
                            if (vnaim != null)
                            {
                                naimF = vnaim.TextValue;
                            }
                            else
                            {
                                naimF = "Переменная $Наименование не найдена";
                            };
                            if (voboz != null)
                            {
                                obozF = voboz.TextValue;
                            }
                            else
                            {
                                obozF = "Переменная $Обозначение не найдена";
                            };
                        }
                        else
                        {
                            err = true;
                            str_err = "Ошибка открытия";
                        }
                    }
                    else
                    {
                        err = true;
                        str_err = "Файл не найден";
                    }

                    if (err == false)
                    {
                        GetFragmentData(docFR, sw, FRname, frag.FullFilePath, obozF, naimF, lev + 1);
                        docFR.Close();
                    }
                }
            }
            return;
        }

        private void Set(Document doc, StreamWriter sw, string name, string path, string oboz, string naim, int lev)
        {
            string offset = "";
            for (int nn = 0; nn < lev; nn++)
                offset += "   ";

            FileInfo file = new FileInfo(name);

            int ConfCount = doc.ModelConfigurations.ConfigurationCount;
            ModelConfiguration[] ConfArray = new ModelConfiguration[ConfCount];
            RegenerateOptions rg = new RegenerateOptions();
            rg.Full = true;
            for (int i = 0; i < ConfArray.Length; i++)
            {
                doc.BeginChanges("Пересчёт");
                string confName = doc.ModelConfigurations.GetConfigurationName(i);
                doc.ModelConfigurations.LoadConfigurationVariables(confName);
                doc.Regenerate(rg);
                if (exportSTEP)
                {
                    ExportToStep exportSTPconf = new ExportToStep(doc);
//                    string fileNameSTPconf = ($"{pathSTPRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.stp");
                    string fileNameSTPconf = ("{0}\\{1}_{2}.stp, pathSTPRU, comfName, doc.FindVariable(\"$Наименование\").TextValue");
                    if (!File.Exists(fileNameSTPconf))
                    {
                        sw.WriteLine(offset + "Имя: " + name);
                        sw.WriteLine(offset + "Путь: " + path);
                        sw.WriteLine(offset + "Наименование: " + naim);
                        sw.WriteLine(offset + "Обозначение: " + confName);
                        sw.WriteLine(offset + "STEP Export: OK");
                        exportSTPconf.Export(fileNameSTPconf);
                    }
                }
                if (exportDXF)
                {
                    ExportToDXF exportDXFconf = new ExportToDXF(doc);
                    Page pgconfDXF = GetPageDXF(doc, "Развертка");
                    Page pgconfDXF2 = GetPageDXF(doc, "Unfolding");
                    //Page pgconfDXF = GetPageDXF2(doc, PageType.Auxiliary);
                    if (pgconfDXF != null || pgconfDXF2 != null)
                    {
                        List<Page> pgDXFconf = new List<Page>();
                        pgDXFconf.Add(pgconfDXF);
                        pgDXFconf.Add(pgconfDXF2);
                        exportDXFconf.ExportPages = pgDXFconf;
//                        string fileNameDXFRU = ($"{pathDXFRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.dxf");
                        string fileNameDXFRU = ("{0}\\{1}_{2}.dxf, pathDXFRU, confName, doc.FindVariable(\"$Наименование\").TextValue");
                        if (!File.Exists(fileNameDXFRU))
                        {
                            sw.WriteLine(offset + "DXF Export: OK");
                            exportDXFconf.Export(fileNameDXFRU);
                        }
                    }
                }
                if (exportPDF)
                {
                    ExportToPDF exportPDFnormalconf = new ExportToPDF(doc);
                    List<Page> pgnormalconf = GetPagesPDF(doc, PageType.Normal);
                    if (pgnormalconf != null)
                    {
                        exportPDFnormalconf.ExportPages = pgnormalconf;
                        exportPDFnormalconf.OpenExportFile = false;
//                        string fileNamePDFRU = ($"{pathPDFRU}\\{confName}_{doc.FindVariable("$Наименование").TextValue}.pdf");
                        string fileNamePDFRU = String.Format("{0}\\{1}_{2}.pdf", pathPDFRU, confName, doc.FindVariable("$Наименование").TextValue);
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFnormal Export: OK");
                            exportPDFnormalconf.Export(fileNamePDFRU);
                        }
                    }
                    ExportToPDF exportPDFBOMconf = new ExportToPDF(doc);
                    List<Page> pgPDFBOMconf = GetPagesPDF(doc, PageType.BillOfMaterials);
                    if (pgPDFBOMconf != null)
                    {
                        exportPDFBOMconf.ExportPages = pgPDFBOMconf;
                        exportPDFBOMconf.OpenExportFile = false;
//                        string fileNamePDFRU = ($"{pathPDFRU}\\{doc.FindVariable("$Обозначение").TextValue}_{doc.FindVariable("$Наименование").TextValue}_СП.pdf");
                        string fileNamePDFRU = String.Format("{0}\\{1}_{2}_СП.pdf", pathPDFRU, doc.FindVariable("$Обозначение").TextValue, doc.FindVariable("$Наименование").TextValue);
                        if (!File.Exists(fileNamePDFRU))
                        {
                            sw.WriteLine(offset + "PDFBOM Export: OK");
                            exportPDFBOMconf.Export(fileNamePDFRU);
                        }
                    }
                }
                doc.EndChanges();
            }

            int n_fr = doc.GetFragments3D().Count;

            foreach (Fragment3D frag in doc.GetFragments3D())
            {
                if (frag.Suppression.Suppress) continue;
                else
                {
                    {
                        Document docFR = null;
                        string obozF = "-";
                        string naimF = "-";
                        bool err = false;
                        string str_err = "";
                        string FRname = frag.FullFilePath;
                        if (lev > 0)
                        {
                            FRname = frag.FilePath;
                            FRname = TFlex.Application.FindPathName(FRname);
                        }

                        if (File.Exists(FRname))
                        {
                            Fragment.OpenPartOptions options = new Fragment.OpenPartOptions();
                            options.DontShowDocument = true;
                            options.QuietMode = true;
                            options.SubstituteGeometry = true;
                            options.SubstituteVariables = true;
                            docFR = frag.OpenPart(options);

                            if (docFR != null)
                            {
                                Variable voboz = docFR.FindVariable("$Обозначение");
                                Variable vnaim = docFR.FindVariable("$Наименование");
                                if (vnaim != null)
                                {
                                    naimF = vnaim.TextValue;
                                }
                                else
                                {
                                    naimF = "Переменная $Наименование не найдена";
                                };
                                if (voboz != null)
                                {
                                    obozF = voboz.TextValue;
                                }
                                else
                                {
                                    obozF = "Переменная $Обозначение не найдена";
                                };
                            }
                            else
                            {
                                err = true;
                                str_err = "Ошибка открытия";
                            }
                        }
                        else
                        {
                            err = true;
                            str_err = "Файл не найден";
                        }

                        if (err == false)
                        {
                            Set(docFR, sw, FRname, frag.FullFilePath, obozF, naimF, lev + 1);
                            docFR.Close();
                        }
                    }
                }
            }
            return;
        } 
    }

    public class ATTR_COM
    {
        public Int16 attr;
        public int pDXF // Экспорт в STP
        {
            get { return (attr & 0x0001); }
            set
            {
                attr = (Int16)(attr & 0xFFFE);
                attr = (Int16)(attr | (Int16)value);
            }
        }

        public int pSTEP // Экспорт в DXF
        {
            get { return ((attr & 0x0002) >> 1); }
            set
            {
                attr = (Int16)(attr & 0xFFFD);
                attr = (Int16)(attr | (Int16)(value << 1));
            }
        }
        public int pPDF // Экспорт в PDF
        {
            get { return ((attr & 0x0004) >> 2); }
            set
            {
                attr = (Int16)(attr & 0xFFFB);
                attr = (Int16)(attr | (Int16)(value << 2));
            }
        }

        public void Set(int dxfs, int steps, int pdfs)
        {
            pDXF = dxfs;
            pSTEP = steps;
            pPDF = pdfs;
        }
    }
}