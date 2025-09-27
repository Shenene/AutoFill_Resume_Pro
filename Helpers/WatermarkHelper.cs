using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AutoFill_Resume_Pro.Helpers
{
    public static class WatermarkHelper
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, string lp);

        private const int EM_SETCUEBANNER = 0x1501;

        public static void SetCue(TextBox textBox, string cueText)
        {
            if (textBox.IsHandleCreated)
            {
                SendMessage(textBox.Handle, EM_SETCUEBANNER, (IntPtr)1, cueText);
            }
        }
    }
}
