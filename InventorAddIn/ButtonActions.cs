#region Namespaces

using System;
using System.Text;
using System.Linq;
using System.Xml;
using System.Reflection;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

using Inventor;

#endregion

namespace InventorAddIn
{
    public class ButtonActions
    {
        public static void Button1_Execute()
        {
            Form1 f = new Form1();
            f.Show();
        }
	}
}
