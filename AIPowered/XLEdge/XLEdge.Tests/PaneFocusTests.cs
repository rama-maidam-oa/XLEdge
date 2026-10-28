using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Hosting;

namespace XLEdge.Tests
{
    /// <summary>
    /// A click on the worksheet grid takes the keyboard back from the pane. Found in the Excel check: with the focus in the
    /// pane's WebView2 (a window of the browser process), a single click on a cell left every keystroke going to the pane.
    /// </summary>
    [TestClass]
    public class PaneFocusTests
    {
        /// <summary>A pane with a focused text box (the WebView2 stand-in) and a "grid" with a focusable control.</summary>
        private static void WithPaneAndGrid(Action<Form, TextBox, Control> test) => Win32Test.OnSta(() =>
        {
            using (var pane = new Form { Text = "pane" })
            using (var excel = new Form { Text = "excel" })
            {
                var paneText = new TextBox();
                pane.Controls.Add(paneText);
                var grid = new TextBox { Text = "grid" };
                excel.Controls.Add(grid);
                excel.Show();
                pane.Show();
                pane.Activate();
                paneText.Focus();
                Win32Test.Pump(TimeSpan.FromMilliseconds(100));
                Assert.AreEqual(paneText.Handle, Win32Test.GetFocus(), "precondition: the pane has the keyboard");
                test(pane, paneText, grid);
            }
        });

        [TestMethod]
        public void ClickOnTheGrid_WithTheFocusInThePane_MovesTheFocusToTheGrid()
        {
            WithPaneAndGrid((pane, paneText, grid) =>
            {
                Assert.IsTrue(PaneFocus.ReturnToGrid(new[] { pane.Handle }, grid.Handle, h => h == grid.Handle));
                Assert.AreEqual(grid.Handle, Win32Test.GetFocus());
            });
        }

        [TestMethod]
        public void SelectionChange_NotFromAClickOnTheGrid_LeavesTheFocusInThePane()
        {
            // Code that selects cells while the user types in the pane must not take the keyboard away.
            WithPaneAndGrid((pane, paneText, grid) =>
            {
                Assert.IsFalse(PaneFocus.ReturnToGrid(new[] { pane.Handle }, pane.Handle, h => h == grid.Handle));
                Assert.IsFalse(PaneFocus.ReturnToGrid(new[] { pane.Handle }, IntPtr.Zero, h => h == grid.Handle));
                Assert.AreEqual(paneText.Handle, Win32Test.GetFocus());
            });
        }

        [TestMethod]
        public void FocusOutsideThePanes_IsLeftAlone()
        {
            WithPaneAndGrid((pane, paneText, grid) =>
            {
                using (var other = new Form())
                {
                    Assert.IsFalse(PaneFocus.ReturnToGrid(new[] { other.Handle, IntPtr.Zero }, grid.Handle, h => h == grid.Handle));
                    Assert.AreEqual(paneText.Handle, Win32Test.GetFocus());
                }
            });
        }

        [TestMethod]
        public void IsWorksheetGrid_IsExcelsGridClass()
        {
            Win32Test.OnSta(() =>
            {
                using (var form = new Form())
                {
                    Assert.IsFalse(PaneFocus.IsWorksheetGrid(form.Handle));
                    Assert.IsFalse(PaneFocus.IsWorksheetGrid(IntPtr.Zero));
                }
                IntPtr grid = Win32Test.CreateWindowOfClass("EXCEL7");
                try { Assert.IsTrue(PaneFocus.IsWorksheetGrid(grid), "a window of Excel's grid class"); }
                finally { Win32Test.DestroyWindow(grid); }
            });
        }
    }
}
