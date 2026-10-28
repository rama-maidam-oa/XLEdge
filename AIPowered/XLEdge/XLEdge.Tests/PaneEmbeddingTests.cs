using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Hosting;

namespace XLEdge.Tests
{
    /// <summary>The shell turns the core's top-level content window into a child of the task pane, sized to it.</summary>
    [TestClass]
    public class PaneEmbeddingTests
    {
        /// <summary>A captioned, resizable popup window - the styles a borderless WPF/WinForms core window may carry.</summary>
        private sealed class PopupForm : Form
        {
            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams cp = base.CreateParams;
                    cp.Style |= unchecked((int)Win32Test.WS_POPUP) | (int)Win32Test.WS_CAPTION | (int)Win32Test.WS_THICKFRAME;
                    return cp;
                }
            }
        }

        private static void WithPaneAndContent(Action<Form, Form> test) => Win32Test.OnSta(() =>
        {
            using (var pane = new Form { Width = 400, Height = 600 })
            using (var content = new PopupForm { Text = "content", Width = 200, Height = 100 })
            {
                pane.Show();
                content.Show();
                test(pane, content);
            }
        });

        [TestMethod]
        public void Embed_MakesTheContentAChildSizedToThePane()
        {
            WithPaneAndContent((pane, content) =>
            {
                long before = Win32Test.Style(content.Handle);
                Assert.AreEqual(Win32Test.WS_POPUP | Win32Test.WS_CAPTION | Win32Test.WS_THICKFRAME,
                    before & (Win32Test.WS_POPUP | Win32Test.WS_CAPTION | Win32Test.WS_THICKFRAME), "precondition: a popup with caption and frame");
                Assert.IsTrue(PaneEmbedding.Embed(content.Handle, pane.Handle, 380, 550));
                long style = Win32Test.Style(content.Handle);
                Assert.AreNotEqual(0L, style & Win32Test.WS_CHILD, "WS_CHILD");
                Assert.AreEqual(0L, style & (Win32Test.WS_POPUP | Win32Test.WS_CAPTION | Win32Test.WS_THICKFRAME), "no popup/caption/frame");
                Assert.AreEqual(pane.Handle, Win32Test.GetParent(content.Handle));
                Assert.AreEqual((380, 550), Win32Test.ClientSize(content.Handle));
            });
        }

        [TestMethod]
        public void Resize_FollowsThePane()
        {
            WithPaneAndContent((pane, content) =>
            {
                PaneEmbedding.Embed(content.Handle, pane.Handle, 380, 550);
                PaneEmbedding.Resize(content.Handle, 300, 200);
                Assert.AreEqual((300, 200), Win32Test.ClientSize(content.Handle));
            });
        }

        [TestMethod]
        public void Release_UnparentsTheContent()
        {
            WithPaneAndContent((pane, content) =>
            {
                PaneEmbedding.Embed(content.Handle, pane.Handle, 380, 550);
                PaneEmbedding.Release(content.Handle);
                // A child taken out of its parent belongs to the desktop; hidden, so it cannot flash at the screen origin
                // before the core destroys it.
                Assert.AreEqual(Win32Test.GetDesktopWindow(), Win32Test.GetParent(content.Handle));
                Assert.IsFalse(Win32Test.IsWindowVisible(content.Handle));
                Assert.IsTrue(PaneEmbedding.IsAlive(content.Handle));
            });
        }

        [TestMethod]
        public void Embed_DeadHandle_ReturnsFalse()
        {
            WithPaneAndContent((pane, content) =>
            {
                IntPtr dead;
                using (var gone = new Form()) { dead = gone.Handle; }
                Assert.IsFalse(PaneEmbedding.IsAlive(dead));
                Assert.IsFalse(PaneEmbedding.Embed(dead, pane.Handle, 10, 10));
                Assert.IsFalse(PaneEmbedding.Embed(IntPtr.Zero, pane.Handle, 10, 10));
                PaneEmbedding.Resize(dead, 10, 10);
                PaneEmbedding.Release(dead);
            });
        }
    }
}
