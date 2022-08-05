using System;
using System.Globalization;
using Microsoft.Win32;
using System.IO;
using System.Windows.Forms;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
//using TFlex.Model.Model3D;
using TFlex.Command;
using System.Drawing;

namespace CREATEBYEXCEL_PLG
{
    public class Factory : PluginFactory
    {
        public override Plugin CreateInstance()
        {
            return new FRAGMENTSTREE_PLG_Plugin(this);
        }

        public override Guid ID
        {
            get { return new Guid("{0f8fad5b-d9cb-469f-a165-70867728950e}"); }
        }

        public override string Name
        {
            get { return "Пакетный экспорт по Excel"; }
        }
    };

    enum Commands
    {
        Create = 1, //Команда создания
        Status =2,
        Debug =3,
    };

    class FRAGMENTSTREE_PLG_Plugin : Plugin
    {
        public FRAGMENTSTREE_PLG_Plugin(Factory factory) : base(factory)
        {
        }

        public static string regedit_str = @"Software\TF Plugins\CREATEBYEXCEL_PLG";

        public static ATTRIBUTES_COM EXT_PAR;

        System.Drawing.Bitmap LoadBitmapResource(string name)
        {
            System.IO.Stream stream = GetType().Assembly.GetManifestResourceStream("CREATEBYEXCEL_PLG.Resource_Files." + name + ".bmp");
            return new System.Drawing.Bitmap(stream);
        }

        public System.Drawing.Icon LoadIconResource(string name)
        {
            System.IO.Stream stream = GetType().Assembly.GetManifestResourceStream("CREATEBYEXCEL_PLG.Resource_Files." + name + ".ico");
            return new System.Drawing.Icon(stream);
        }

        protected override void OnInitialize()
        {
            base.OnInitialize();
            EXT_PAR = new ATTRIBUTES_COM();
        }

        protected override void OnCreateTools()
        {
            base.OnCreateTools();

            RegisterCommand((int)Commands.Create,
                "Пакетный экспорт по Excel", LoadIconResource("Коннектор_small"), LoadIconResource("Коннектор"));

            int[] CmdIDs = new int[]
            {
                (int)Commands.Create,
            };

            TFlex.Menu submenu = new TFlex.Menu();
            submenu.CreatePopup();

            submenu.Append((int)Commands.Create, "&Пакетный экспорт по Excel", this);
            TFlex.RibbonGroup ribbonGroup = TFlex.RibbonBar.ApplicationsTab.AddGroup("Пакетный экспорт по Excel");
            ribbonGroup.AddButton((int)Commands.Create, this);
            TFlex.Application.ActiveMainWindow.InsertPluginSubMenu(this.Name, submenu, TFlex.MainWindow.InsertMenuPosition.PluginSamples, this);

            CreateToolbar(this.Name, CmdIDs);
        }

        protected override void OnCommand(Document document, int id)
        {
            switch ((Commands)id)
            {
                default:
                    base.OnCommand(document, id);
                    break;

                case Commands.Create:
                    {
                        ComParams par = new ComParams(EXT_PAR);
                        if (par.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            par.SetParams(EXT_PAR);
                            CommandManager Command = new CommandManager();
                            Command.OK(document, EXT_PAR);
                        }
                        else System.Windows.Forms.MessageBox.Show("Выполнение плагина отменено", "Пакетный экспорт по Excel");
                        break;
                    }
            }
        }

        protected override void OnUpdateCommand(CommandUI cmdUI)
        {
            if (cmdUI == null)
                return;

            if (cmdUI.Document == null)
            {
                cmdUI.Enable(false);
                return;
            }

            cmdUI.Enable();
        }

        protected override void NewDocumentCreatedEventHandler(DocumentEventArgs args)
        {
            args.Document.AttachPlugin(this);
        }

        protected override void DocumentOpenEventHandler(DocumentEventArgs args)
        {
            args.Document.AttachPlugin(this);
        }
    }
}
