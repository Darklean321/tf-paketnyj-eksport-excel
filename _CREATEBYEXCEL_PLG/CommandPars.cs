using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CREATEBYEXCEL_PLG
{
    public partial class ComParams : Form
    {
        public ComParams(ATTRIBUTES_COM par)
        {
            InitializeComponent();

            pEXSTEP.CheckState = par.parameterSTEP == 1 ? CheckState.Checked : CheckState.Unchecked;
            pEXDXF.CheckState = par.parameterDXF == 1 ? CheckState.Checked : CheckState.Unchecked;
            pEXPDF.CheckState = par.parameterPDF == 1 ? CheckState.Checked : CheckState.Unchecked;
            pEXDOCs.CheckState = par.parameterDOCs == 1? CheckState.Checked : CheckState.Unchecked;
            pEXEXCEL.CheckState = par.parameterDOCs == 1 ? CheckState.Checked : CheckState.Unchecked;
        }

        public void SetParams(ATTRIBUTES_COM par)
        {
            par.parameterDXF = pEXDXF.CheckState == CheckState.Checked ? 1 : 0;
            par.parameterSTEP = pEXSTEP.CheckState == CheckState.Checked ? 1 : 0;
            par.parameterPDF = pEXPDF.CheckState == CheckState.Checked ? 1 : 0;
            par.parameterDOCs = pEXDOCs.CheckState == CheckState.Checked ? 1 : 0;
            par.parameterEXCEL = pEXEXCEL.CheckState == CheckState.Checked ? 1 : 0;

            RegistryKey test = Registry.CurrentUser.OpenSubKey(FRAGMENTSTREE_PLG_Plugin.regedit_str, RegistryKeyPermissionCheck.ReadWriteSubTree);

            if (test == null)
            {
                test = Registry.CurrentUser.CreateSubKey(FRAGMENTSTREE_PLG_Plugin.regedit_str);
            }

            test.SetValue("ATTR_COM", par.attribute.ToString());
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            /*par.pSTEP = 0;
            par.pDXF = 0;
            par.pPDF = 0;*/
            return;
        }

    }
}
