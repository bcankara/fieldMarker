using System;
using System.Drawing;
using System.Windows.Forms;

namespace fieldMarker
{
    public static class NotificationForm
    {
        public static void Show(string message, string title, bool isError = false)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, 
                isError ? MessageBoxIcon.Error : MessageBoxIcon.Information);
        }
    }
} 