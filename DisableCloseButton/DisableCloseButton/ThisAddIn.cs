using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace DisableCloseButton
{
    public partial class ThisAddIn
    {
        // Win32 APIのインポート
        [DllImport("USER32.DLL")]
        private static extern IntPtr
            GetSystemMenu(IntPtr hWnd, UInt32 bRevert);

        [DllImport("USER32.DLL")]
        private static extern IntPtr
            DrawMenuBar(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "FindWindowA")]
        private static extern IntPtr
            FindWindow(string lpClassName, string lpWindowName);

        // #1
        [DllImport("USER32.DLL")]
        private static extern IntPtr
            DeleteMenu(IntPtr hMenu, UInt32 nPosition, UInt32 wFlags);

        /*
        // #2
        [DllImport("USER32.DLL")]
        private static extern IntPtr
            ModifyMenu(IntPtr hMenu, UInt32 nPosition, UInt32 wFlags, UInt32 nFunction, string lpWindowName);
        */

        /*
        // #3
        [DllImport("USER32.DLL")]
        private static extern IntPtr
            EnableMenuItem (IntPtr hMenu, UInt32 nPosition, UInt32 wFlags);
        */

        // ［閉じる］ボタンを無効化するための値
        private const UInt32 MFS_ENABLED = 0x00000000;
        private const UInt32 MFS_GLAYED = 0x00000001;
        private const UInt32 MFS_DISENABLED = 0x00000002;
        private const UInt32 SC_CLOSE = 0x0000F060;
        private const UInt32 MF_BYCOMMAND = 0x00000000;

        public static void DisableCloseButton()
        {
            IntPtr rc;
            IntPtr hWnd;
            hWnd = FindWindow("rctrl_renwnd32", null);

            // ［閉じる］ボタンの無効化
            IntPtr hMenu = GetSystemMenu(hWnd, 0);

            // #1
            rc = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND);
            // #2
            //  rc = ModifyMenu(hMenu, SC_CLOSE, MF_BYCOMMAND, MFS_GLAYED, "閉じる");
            // #3
            //  rc = EnableMenuItem (hMenu, SC_CLOSE, MF_BYCOMMAND | MFS_GLAYED);

            rc = DrawMenuBar(hWnd);

        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            DisableCloseButton();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    を Outlook のシャットダウン時に実行する必要があります。https://go.microsoft.com/fwlink/?LinkId=506785 をご覧ください
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
