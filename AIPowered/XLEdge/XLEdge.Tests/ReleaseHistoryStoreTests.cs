using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class ReleaseHistoryStoreTests
    {
        private static ReleaseEntry Entry(string folder, string version = "11.1.2", string date = "2026-10-28T10:00:00") =>
            new ReleaseEntry { Version = version, ReleaseDate = date, FolderName = folder, Checksum = "c", Notes = "n", ContractVersion = 1 };

        private static void MakeRelease(string versionsPath, string folder, bool withDll = true)
        {
            Directory.CreateDirectory(Path.Combine(versionsPath, folder));
            if (withDll) File.WriteAllText(Path.Combine(versionsPath, folder, "XLEdge.Addin.Core.dll"), "x");
        }

        [TestMethod]
        public void Read_MissingEmptyNull_ReturnsEmpty()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json");
                Assert.AreEqual(0, ReleaseHistoryStore.ReadAll(file).Count, "missing");
                foreach (var content in new[] { "", "   ", "null" })
                {
                    File.WriteAllText(file, content);
                    Assert.AreEqual(0, ReleaseHistoryStore.ReadAll(file).Count, "content: '" + content + "'");
                }
            }
        }

        [TestMethod]
        public void Read_Corrupt_ThrowsJsonException()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json");
                File.WriteAllText(file, "[{\"Version\":");
                Assert.ThrowsExactly<JsonException>(() => ReleaseHistoryStore.ReadAll(file));
            }
        }

        [TestMethod]
        public void Append_WritesPascalCaseAndDedupes()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("AddinCore", "ReleaseHistory.json");
                ReleaseHistoryStore.Append(file, Entry("v11.1.2_a"));
                ReleaseHistoryStore.Append(file, Entry("v11.1.2_a"));
                ReleaseHistoryStore.Append(file, Entry("v11.1.3_b", "11.1.3"));

                var all = ReleaseHistoryStore.ReadAll(file);
                CollectionAssert.AreEqual(new[] { "v11.1.2_a", "v11.1.3_b" }, all.Select(e => e.FolderName).ToArray());
                Assert.AreEqual("Install", all[0].Source);
                Assert.AreEqual(1, all[0].ContractVersion);
                string json = File.ReadAllText(file);
                StringAssert.Contains(json, "\"FolderName\"");
                StringAssert.Contains(json, "\"ContractVersion\"");
            }
        }

        [TestMethod]
        public void Read_IsCaseInsensitive()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json");
                File.WriteAllText(file, "[{\"version\":\"11.1.2\",\"folderName\":\"v1\",\"contractVersion\":1}]");
                var e = ReleaseHistoryStore.ReadAll(file).Single();
                Assert.AreEqual("11.1.2", e.Version);
                Assert.AreEqual("v1", e.FolderName);
                Assert.AreEqual(1, e.ContractVersion);
            }
        }

        [TestMethod]
        public void Write_FirstTimeUsesMove_LeavesNoTmp()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json");
                ReleaseHistoryStore.Append(file, Entry("v1"));
                ReleaseHistoryStore.Append(file, Entry("v2"));
                Assert.IsTrue(File.Exists(file));
                Assert.IsFalse(File.Exists(file + ".tmp"));
                Assert.AreEqual(2, ReleaseHistoryStore.ReadAll(file).Count);
            }
        }

        [TestMethod]
        public void Write_TargetBrieflyOpenElsewhere_StillSucceeds()
        {
            // Antivirus / indexers open a freshly written file for a moment; File.Replace then fails with
            // "Unable to remove the file to be replaced" unless the write is retried.
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json");
                ReleaseHistoryStore.Append(file, Entry("v1"));
                var reader = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.Read);
                var closer = new Thread(() => { Thread.Sleep(150); reader.Dispose(); });
                closer.Start();
                try
                {
                    ReleaseHistoryStore.Append(file, Entry("v2"));
                }
                finally
                {
                    closer.Join();
                }
                Assert.AreEqual(2, ReleaseHistoryStore.ReadAll(file).Count);
                Assert.IsFalse(File.Exists(file + ".tmp"));
            }
        }

        [TestMethod]
        public void Reconcile_DropsMissingFolderAndFolderWithoutDll_RewritesOnlyOnChange()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json"), versions = t.Combine("Versions");
                MakeRelease(versions, "keep");
                MakeRelease(versions, "nodll", withDll: false);
                foreach (var f in new[] { "keep", "nodll", "gone" }) ReleaseHistoryStore.Append(file, Entry(f));
                ReleaseHistoryStore.Append(file, Entry(""));

                var kept = ReleaseHistoryStore.Reconcile(file, versions);
                CollectionAssert.AreEqual(new[] { "keep" }, kept.Select(e => e.FolderName).ToArray());
                CollectionAssert.AreEqual(new[] { "keep" }, ReleaseHistoryStore.ReadAll(file).Select(e => e.FolderName).ToArray());

                var stamp = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                File.SetLastWriteTimeUtc(file, stamp);
                ReleaseHistoryStore.Reconcile(file, versions);
                Assert.AreEqual(stamp, File.GetLastWriteTimeUtc(file), "unchanged catalog must not be rewritten");
            }
        }

        [TestMethod]
        public void Delete_Results()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json"), versions = t.Combine("Versions");
                MakeRelease(versions, "v-old");
                MakeRelease(versions, "v-active");
                foreach (var f in new[] { "v-old", "v-active", "..", "a\\b", "C:\\x" }) ReleaseHistoryStore.Append(file, Entry(f));

                Assert.AreEqual(ReleaseDeleteResult.NotFound, ReleaseHistoryStore.Delete(file, versions, "nope", "v-active", out _));
                Assert.AreEqual(ReleaseDeleteResult.ActiveRelease, ReleaseHistoryStore.Delete(file, versions, "V-ACTIVE", "v-active", out _));
                foreach (var bad in new[] { "..", "a\\b", "C:\\x" })
                    Assert.AreEqual(ReleaseDeleteResult.InvalidFolderName, ReleaseHistoryStore.Delete(file, versions, bad, "v-active", out _), bad);

                Assert.AreEqual(ReleaseDeleteResult.Deleted, ReleaseHistoryStore.Delete(file, versions, "v-old", "v-active", out var error));
                Assert.IsNull(error);
                Assert.IsFalse(Directory.Exists(Path.Combine(versions, "v-old")));
                Assert.IsTrue(Directory.Exists(Path.Combine(versions, "v-active")));
                Assert.IsFalse(ReleaseHistoryStore.ReadAll(file).Any(e => e.FolderName == "v-old"));
            }
        }

        [TestMethod]
        public void Delete_LockedFolder_ReportsFolderDeleteFailed()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json"), versions = t.Combine("Versions");
                MakeRelease(versions, "v-locked");
                ReleaseHistoryStore.Append(file, Entry("v-locked"));
                using (File.Open(Path.Combine(versions, "v-locked", "XLEdge.Addin.Core.dll"), FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    Assert.AreEqual(ReleaseDeleteResult.FolderDeleteFailed, ReleaseHistoryStore.Delete(file, versions, "v-locked", null, out var error));
                    Assert.IsNotNull(error);
                }
                Assert.IsTrue(ReleaseHistoryStore.ReadAll(file).Any(e => e.FolderName == "v-locked"), "catalog keeps the entry");
            }
        }

        [TestMethod]
        public void BuildFolderName_ReplacesInvalidChars()
        {
            Assert.AreEqual("v11.1.2_2026-10-28T14-05-09", ReleaseHistoryStore.BuildFolderName("11.1.2", "2026-10-28T14:05:09"));
            Assert.AreEqual("v11.1.2_", ReleaseHistoryStore.BuildFolderName("11.1.2", null));
        }

        [TestMethod]
        public void IsSafeFolderName()
        {
            Assert.IsTrue(ReleaseHistoryStore.IsSafeFolderName("v11.1.2_2026-10-28T14-05-09"));
            foreach (var bad in new[] { null, "", " ", ".", "..", "a\\b", "a/b", "C:\\x", "\\x" })
                Assert.IsFalse(ReleaseHistoryStore.IsSafeFolderName(bad), bad ?? "<null>");
        }

        [TestMethod]
        public void Lock_HeldElsewhere_TimesOut()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json");
                using (var held = new ManualResetEventSlim())
                using (var release = new ManualResetEventSlim())
                {
                    var holder = new Thread(() =>
                    {
                        using (var m = new Mutex(false, ReleaseHistoryStore.LockName(file)))
                        {
                            m.WaitOne();
                            held.Set();
                            release.Wait();
                            m.ReleaseMutex();
                        }
                    });
                    holder.Start();
                    held.Wait();
                    try
                    {
                        var ex = Assert.ThrowsExactly<TimeoutException>(() => ReleaseHistoryStore.ReadAll(file, TimeSpan.FromMilliseconds(200)));
                        StringAssert.Contains(ex.Message, file);
                    }
                    finally
                    {
                        release.Set();
                        holder.Join();
                    }
                }
            }
        }

        [TestMethod]
        public void Lock_Abandoned_CountsAsAcquired()
        {
            using (var t = new TempFolder())
            {
                string file = t.Combine("ReleaseHistory.json");
                var abandoner = new Thread(() => new Mutex(false, ReleaseHistoryStore.LockName(file)).WaitOne());
                abandoner.Start();
                abandoner.Join();
                ReleaseHistoryStore.Append(file, Entry("v1"));
                Assert.AreEqual(1, ReleaseHistoryStore.ReadAll(file).Count);
            }
        }

        [TestMethod]
        public void LockName_IsPerFile()
        {
            string a = ReleaseHistoryStore.LockName(@"C:\Users\a\AppData\Local\X\AddinCore\ReleaseHistory.json");
            string b = ReleaseHistoryStore.LockName(@"C:\Users\b\AppData\Local\X\AddinCore\ReleaseHistory.json");
            StringAssert.StartsWith(a, @"Global\XLEdge_ReleaseHistory_");
            Assert.AreNotEqual(a, b);
            Assert.AreEqual(a, ReleaseHistoryStore.LockName(@"c:\users\A\appdata\local\x\addincore\releasehistory.json"));
        }
    }
}
