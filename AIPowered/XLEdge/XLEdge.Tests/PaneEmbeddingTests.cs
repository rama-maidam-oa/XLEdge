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
        public void Embed_ParentGone_ReturnsFalse_AndLeavesTheContentAsItWas()
        {
            WithPaneAndContent((pane, content) =>
            {
                IntPtr deadParent;
                using (var gone = new Form()) { deadParent = gone.Handle; }
                long before = Win32Test.Style(content.Handle);
                Assert.IsFalse(PaneEmbedding.Embed(content.Handle, deadParent, 100, 100, out _));
                Assert.AreEqual(before, Win32Test.Style(content.Handle), "a failed embed must not leave a parentless WS_CHILD");
            });
        }

        [TestMethod]
        public void CreateUnderDpiOf_GivesTheContentThePanesDpiAwareness_SoItCanBeEmbedded()
        {
            // Office switches the thread's DPI awareness context; Windows refuses to parent a window into one of another
            // awareness (in-process, without mixed hosting). The core's window is therefore created under the pane's context.
            Win32Test.OnSta(() =>
            {
                IntPtr previous = Win32Test.SetThreadDpiAwarenessContext(Win32Test.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
                Assert.AreNotEqual(IntPtr.Zero, previous, "per-thread DPI awareness needs Windows 10 1607 or later");
                try
                {
                    using (var pane = new Form { Width = 400, Height = 600 })
                    {
                        pane.Show();
                        Win32Test.SetThreadDpiAwarenessContext(Win32Test.DPI_AWARENESS_CONTEXT_UNAWARE);

                        using (var mismatched = new Form())
                        {
                            mismatched.Show();
                            long before = Win32Test.Style(mismatched.Handle);
                            Assert.IsFalse(PaneEmbedding.Embed(mismatched.Handle, pane.Handle, 100, 100, out int refused), "precondition: a mismatched window cannot be embedded");
                            Assert.AreNotEqual(0, refused, "the error of the refused SetParent is reported (for the log)");
                            Assert.AreEqual(before, Win32Test.Style(mismatched.Handle), "a refused embed must not leave a parentless WS_CHILD");
                        }

                        Form matched = null;
                        IntPtr handle = PaneEmbedding.CreateUnderDpiOf(pane.Handle, () => { matched = new Form(); matched.Show(); return matched.Handle; });
                        using (matched)
                        {
                            Assert.IsTrue(PaneEmbedding.Embed(handle, pane.Handle, 100, 100, out int error), "error " + error);
                            Assert.IsTrue(Win32Test.AreDpiAwarenessContextsEqual(Win32Test.GetThreadDpiAwarenessContext(), Win32Test.DPI_AWARENESS_CONTEXT_UNAWARE),
                                "the thread's own context is restored");
                        }
                    }
                }
                finally
                {
                    Win32Test.SetThreadDpiAwarenessContext(previous);
                }
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
