using Microsoft.Win32;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CCGExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Directory)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".zip")]
    public class CCGExtension : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            if (SelectedItemPaths.Count() > 1)
                return false;
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem tsmi = new ToolStripMenuItem
            {
                Text = "用CCGViewer打开"
            };

            tsmi.Click += (sender, e) => Open();
            menu.Items.Add(tsmi);
            return menu;
        }

        private void Open()
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\CCGViewer");
            RegistryKey dd = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64);
            if (null == key)
            {
                MessageBox.Show("没有找到CCGViewer的注册表信息！");
                return;
            }
            string k = (string)key.GetValue("path", string.Empty);
            if (k == string.Empty)
            {
                MessageBox.Show("CCGViewer的注册表信息有误！");
                return;
            }
            key.Close();
            System.Diagnostics.Process.Start(k, SelectedItemPaths.ElementAt(0));
        }
    }
}
