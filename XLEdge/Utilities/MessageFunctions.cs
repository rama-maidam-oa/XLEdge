using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using XLEdge.Helpers;
using XLEdge.Views;
using Excel=Microsoft.Office.Interop.Excel;

namespace XLEdge.Utilities
{
#nullable enable
    public static class MessageFunctions
    {
        /// <summary>
        /// Shows the GLSense WPF message window safely on a UI (STA) thread.
        /// - Prefers Application.Current.Dispatcher (existing WPF UI thread).
        /// - Sets owner to Excel HWND for correct modality/z-order.
        /// - Falls back to a temporary STA thread if no dispatcher exists.
        /// </summary>

        public static MessageBoxResult XLEdgeMessage(
                string msg,
                MessageBoxIcon msgIcon,
                MessageBoxButtons buttons = MessageBoxButtons.OK)
        {
            try
            {
                ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp);
                IntPtr excelHwnd = IntPtr.Zero;

                try { excelHwnd = (IntPtr)(excelApp?.Hwnd ?? 0); }
                catch (Exception ex)
                {
                    // Safe to ignore: best-effort Excel HWND lookup; falls back to CenterScreen below.
                    LogUtility.LogDebug($"{nameof(XLEdgeMessage)}: failed to read Excel HWND - {ex.Message}");
                }

                return WpfAppManager.InvokeOnWpfThread(() =>
                {
                    var win = new XLEdgeMessageWindow(msg, msgIcon, buttons);

                    if (excelHwnd != IntPtr.Zero)
                    {
                        win.SetExcelOwner(excelHwnd);
                        win.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                    }
                    else
                    {
                        win.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    }

                    win.ShowDialog();

                    if (excelHwnd != IntPtr.Zero)
                    {
                        try { NativeMethods.SetForegroundWindow(excelHwnd); }
                        catch (Exception ex)
                        {
                            // Safe to ignore: best-effort refocus of Excel after the message window
                            // closes; cosmetic only.
                            LogUtility.LogDebug($"{nameof(XLEdgeMessage)}: failed to restore Excel foreground focus - {ex.Message}");
                        }
                    }

                    return win.Result;
                });
            }
            catch (Exception ex)
            {
                // Last-resort fallback to a native MessageBox if the custom XLEdgeMessageWindow
                // itself fails to construct or show.
                LogUtility.LogException(ex);
                return System.Windows.MessageBox.Show(msg, "Orbit GLSense",
                    ConvertButtons(buttons), ConvertIcon(msgIcon));
            }
        }

        private static MessageBoxButton ConvertButtons(MessageBoxButtons buttons)
        {
            return buttons switch
            {
                MessageBoxButtons.OK => MessageBoxButton.OK,
                MessageBoxButtons.OKCancel => MessageBoxButton.OKCancel,
                MessageBoxButtons.YesNo => MessageBoxButton.YesNo,
                MessageBoxButtons.YesNoCancel => MessageBoxButton.YesNoCancel,
                _ => MessageBoxButton.OK
            };
        }

        private static MessageBoxImage ConvertIcon(MessageBoxIcon icon)
        {
            return icon switch
            {
                MessageBoxIcon.Error => MessageBoxImage.Error,
                MessageBoxIcon.Warning => MessageBoxImage.Warning,
                MessageBoxIcon.Information => MessageBoxImage.Information,
                MessageBoxIcon.Question => MessageBoxImage.Question,
                _ => MessageBoxImage.None
            };
        }
        internal static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("user32.dll")]
            internal static extern bool SetForegroundWindow(IntPtr hWnd);
        }
    }
#nullable restore
}
