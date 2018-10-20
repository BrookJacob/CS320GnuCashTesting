using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestStack.White;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.MenuItems;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.InputDevices;
using System.Windows.Automation;

namespace UnitTestProject2
{
    [TestClass]
    public class UnitTest1
    {
        private string appPathUnderTest = @"C:\Users\jbrook19\Desktop\GnuTests";
        private string appUnderTest = @"Jacob.gnucash";
        private string windowPrefix = "Jacob.gnucash ";
        private string[] menuFileExit = { "File", "Exit" };

        class AppUnderTest
        {
            public Application app;
            public Window w;
        }

        [TestMethod]
        public void Closed()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                SetDimensions(aut, 1000, 1000);
                TerminateApp(aut);
                Assert.IsFalse(aut.w.IsClosed);
            }
        }

        [TestMethod]
        public void Maximize()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                SetDimensions(aut, 1000, 1000);
                var m = aut.w.Get<Button>("Maximize");
                m.Click();
                
                Assert.AreEqual(aut.w.DisplayState, DisplayState.Maximized);
                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void RestoreDown()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var m = aut.w.Get<Button>("Restore");
                m.Click();
                Assert.AreEqual(aut.w.DisplayState, DisplayState.Restored);
                SetDimensions(aut, 1000, 1000);
                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void Minimize()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var m = aut.w.Get<Button>("Minimize");
                m.Click();
                Assert.AreEqual(aut.w.DisplayState, DisplayState.Minimized);
                TerminateApp(aut);
            }
        }
        //Does not work when run with rest of test suite, but it works and we don't know why
        [TestMethod]
        public void RightClickMax()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                SetDimensions(aut, 1000, 1000);
                var m = aut.w.TitleBar;
                m.RightClick();
                var a = new System.Windows.Point(560, 122);
                aut.w.Mouse.Click(a);
                Assert.AreEqual(aut.w.DisplayState, DisplayState.Maximized);
                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void RightClickRestore()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var m = aut.w.TitleBar;
                m.RightClick();
                var a = new System.Windows.Point(1005, 25);
                aut.w.Mouse.Click(a);
                Assert.AreEqual(aut.w.DisplayState, DisplayState.Restored);
                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void RightClickMin()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                SetDimensions(aut, 1000, 1000);
                var m = aut.w.TitleBar;
                m.RightClick();
                var a = new System.Windows.Point(560, 100);
                aut.w.Mouse.Click(a);
                Assert.AreEqual(aut.w.DisplayState, DisplayState.Minimized);
                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void RightClickClose()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                SetDimensions(aut, 1000, 1000);
                var m = aut.w.TitleBar;
                m.RightClick();
                var a = new System.Windows.Point(560, 150);
                aut.w.Mouse.Click(a);
                TerminateApp(aut);
                Assert.IsFalse(aut.w.IsClosed);
            }
        }

        [TestMethod]
        public void DoubleClickTitleBar()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                SetDimensions(aut, 1000, 1000);
                var m = aut.w.TitleBar;
                m.DoubleClick();
                System.Threading.Thread.Sleep(3000);
                Assert.AreEqual(aut.w.DisplayState, DisplayState.Maximized);
                m.DoubleClick();
                SetDimensions(aut, 1000, 1000);
                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void QuitWithoutSaving()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(26, 44);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(117, 365);
                aut.w.Mouse.Click(b);
                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void CreateAccounts()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(151, 45);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(253, 170);
                aut.w.Mouse.Click(b);
                var c = new System.Windows.Point(755, 677);
                aut.w.Mouse.Click(c);
                var d = new System.Windows.Point(677, 677);
                aut.w.Mouse.Click(d);
                var e = new System.Windows.Point(677, 677);
                aut.w.Mouse.Click(e);
                var f = new System.Windows.Point(390, 420);
                aut.w.Mouse.Click(f);
                var g = new System.Windows.Point(755, 677);
                aut.w.Mouse.Click(g);
                var h = new System.Windows.Point(677, 677);
                aut.w.Mouse.Click(h);

                TerminateApp(aut);

                var z = new System.Windows.Point(657, 585);
                aut.w.Mouse.Click(z);
            }

            AppUnderTest aut2 = StartApp();
            if (aut2.w != null)
            {
                var a = new System.Windows.Point(85, 185);
                aut2.w.Mouse.DoubleClick(a);
                var b = new System.Windows.Point(100, 205);
                aut2.w.Mouse.Click(b);
                var c = new System.Windows.Point(275, 70);
                aut2.w.Mouse.Click(c);
                aut2.w = GetWindow(aut2, "Edit Account - ");
                Assert.IsTrue(aut2.w.Title.Contains("Edit Account - "));
                aut2.w.Close();
                aut2.w = GetWindow(aut2, "Jacob.gnucash ");

                TerminateApp(aut2);

                var z = new System.Windows.Point(657, 585);
                aut2.w.Mouse.Click(z);
            }
        }

        //Does not like to click on subaccounts under expenses
        [TestMethod]
        public void ChangeAccountName()
        {
            string AccountTitleToTest = "Taco Bell Fund";
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(130, 185);
                aut.w.Mouse.DoubleClick(a);
                var b = new System.Windows.Point(85, 207);
                aut.w.Mouse.Click(b);
                var c = new System.Windows.Point(273, 73);
                aut.w.Mouse.Click(c);
                var d = new System.Windows.Point(500, 300);
                aut.w.Mouse.DoubleClick(d);
                aut.w = GetWindow(aut, "Edit Account - ");
                aut.w.Keyboard.Enter(AccountTitleToTest);
                var e = new System.Windows.Point(658, 789);
                aut.w.Mouse.Click(e);
                aut.w = GetWindow(aut, "Jacob.gnucash ");
                var f = new System.Windows.Point(130, 185);
                aut.w.Mouse.DoubleClick(f);
                TerminateApp(aut);

                var g = new System.Windows.Point(657, 585);
                aut.w.Mouse.Click(g);
            }

            AppUnderTest aut2 = StartApp();
            if (aut2.w != null)
            {
                var a = new System.Windows.Point(130, 185);
                aut2.w.Mouse.DoubleClick(a);
                var b = new System.Windows.Point(85, 207);
                aut2.w.Mouse.Click(b);
                var c = new System.Windows.Point(273, 73);
                aut2.w.Mouse.Click(c);
                var d = new System.Windows.Point(500, 300);
                aut2.w.Mouse.DoubleClick(d);
                aut2.w = GetWindow(aut2, "Edit Account - ");
                Assert.IsTrue(aut2.w.Title.Contains(AccountTitleToTest));

                aut2.w.Close();
                aut2.w = GetWindow(aut2, "Jacob.gnucash ");

                TerminateApp(aut2);

                var f = new System.Windows.Point(657, 585);
                aut2.w.Mouse.Click(f);
            }

        }


        [TestMethod]
        public void DeleteAccounts()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(73, 183);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(430, 80);
                aut.w.Mouse.Click(b);
                var c = new System.Windows.Point(243, 412);
                aut.w.Mouse.Click(c);
                var d = new System.Windows.Point(760, 727);
                aut.w.Mouse.Click(d);
                var e = new System.Windows.Point(623, 561);
                aut.w.Mouse.Click(e);

                TerminateApp(aut);

                var f = new System.Windows.Point(657, 585);
                aut.w.Mouse.Click(f);
            }

            AppUnderTest aut2 = StartApp();
            if (aut2.w != null)
            {
                var a = new System.Windows.Point(73, 183);
                aut2.w.Mouse.Click(a);
                var b = new System.Windows.Point(430, 80);
                aut2.w.Mouse.Click(b);
                var ws = aut2.app.GetWindows();
                Assert.AreEqual(1, ws.Count);

                TerminateApp(aut2);

                var f = new System.Windows.Point(657, 585);
                aut2.w.Mouse.Click(f);
            }
        }

        [TestMethod]
        public void ClickFile()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(26, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickEdit()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(65, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickView()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(102, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickActionsTheySpeakLouderThanWords()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(156, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickBusiness()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(213, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickReports()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(272, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickTools()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(331, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickWindowsItsBetterThanOSX()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(386, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ClickHelpIsWhatWeNeed()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(443, 44);
                aut.w.Mouse.Click(a);
                aut.w.Mouse.Click(a);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ChangeAccountColor()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(19, 183);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(85, 207);
                aut.w.Mouse.Click(b);
                var c = new System.Windows.Point(273, 73);
                aut.w.Mouse.Click(c);
                var d = new System.Windows.Point(480, 530);
                aut.w.Mouse.Click(d);
                var e = new System.Windows.Point(275, 455);
                aut.w.Mouse.Click(e);
                var f = new System.Windows.Point(590, 681);
                aut.w.Mouse.Click(f);
                var g = new System.Windows.Point(630, 830);
                aut.w.Mouse.Click(g);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ChangeAccountFraction()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(19, 183);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(85, 207);
                aut.w.Mouse.Click(b);
                var c = new System.Windows.Point(273, 73);
                aut.w.Mouse.Click(c);
                var d = new System.Windows.Point(530, 490);
                aut.w.Mouse.Click(d);
                var e = new System.Windows.Point(509, 771);
                aut.w.Mouse.Click(e);
                var g = new System.Windows.Point(630, 830);
                aut.w.Mouse.Click(g);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void HideToolBar()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(102, 44);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(102, 70);
                aut.w.Mouse.Click(b);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ShowToolBar()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(102, 44);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(102, 70);
                aut.w.Mouse.Click(b);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void HideSummary()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(102, 44);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(102, 93);
                aut.w.Mouse.Click(b);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ShowSummary()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(102, 44);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(102, 93);
                aut.w.Mouse.Click(b);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void HideStatus()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(102, 44);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(102, 118);
                aut.w.Mouse.Click(b);

                TerminateApp(aut);
            }
        }

        [TestMethod]
        public void ShowStatus()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var a = new System.Windows.Point(102, 44);
                aut.w.Mouse.Click(a);
                var b = new System.Windows.Point(102, 118);
                aut.w.Mouse.Click(b);

                TerminateApp(aut);
            }
        }


        [TestMethod]
        public void ChangeTransactionDate()
        {
            AppUnderTest aut = StartApp();
            if (aut.w != null)
            {
                var x = new System.Windows.Point(67, 185);
                aut.w.Mouse.DoubleClick(x);
                var a = new System.Windows.Point(93, 204);
                aut.w.Mouse.DoubleClick(a);
                var b = new System.Windows.Point(102, 182);
                aut.w.Mouse.Click(b);
                var c = new System.Windows.Point(52, 249);
                aut.w.Mouse.DoubleClick(c);

                TerminateApp(aut);

                var e = new System.Windows.Point(666, 600);
                aut.w.Mouse.DoubleClick(e);
                var f = new System.Windows.Point(666, 600);
                aut.w.Mouse.DoubleClick(f);

            }
        }

        private AppUnderTest StartApp()
        {
            AppUnderTest aut = new AppUnderTest();
            var appPath = Path.Combine(appPathUnderTest, appUnderTest);
            aut.app = Application.Launch(appPath);
            var ws = aut.app.GetWindows();
            var start = DateTime.Now;
            var timeout = new TimeSpan(0, 0, 30);
            while ((ws == null || ws.Count == 0) && DateTime.Now - start < timeout)
            {
                ws = aut.app.GetWindows();
            }
            while (aut.w == null && DateTime.Now - start < timeout)
            {
                System.Diagnostics.Debug.Write(".");
                try
                {
                    foreach (var win in ws)
                    {
                        System.Diagnostics.Debug.Write(win.Title);
                        if (win.Title.StartsWith(windowPrefix))
                            aut.w = win;
                    }
                }
                catch
                {
                    //Might end up here if the app has a splash screen, and that window goes away. Refresh the windows list
                    ws = aut.app.GetWindows();
                }
            }

            return aut;
        }

        private Window GetWindow(AppUnderTest aut, string windowNameStartsWith)
        {
            var ws = aut.app.GetWindows();
            var w = aut.w;
            foreach (var win in ws)
            {
                System.Diagnostics.Debug.Write(win.Title);
                if (win.Title.Contains(windowNameStartsWith))
                    w = win;
            }
            return w;
        }

        private void TerminateApp(AppUnderTest aut)
        {
            try
            {
                if (aut.w.MenuBar != null)
                {
                    var m = aut.w.MenuBar.MenuItem(menuFileExit);
                    if (m != null)
                    {
                        m.Click();
                    }
                }
                else
                {
                    aut.w.Close();
                }
            } catch (TestStack.White.AutomationException)
            {
                aut.w.Close();
            }
        }

        private void SetDimensions(AppUnderTest aut, double width, double height)
        {
            var trans = (TransformPattern)aut.w.AutomationElement.GetCurrentPattern(TransformPattern.Pattern);
            trans.Resize(width, height);
            trans.Move(0, 0);
        }

        private void InterrogateApp()
        {
            AppUnderTest aut = StartApp();

            if (aut.w != null)
            {
                InterrogateItem(aut.w);
                TerminateApp(aut);
            }
        }

        private void InterrogateItem(Window w)
        {
            InterrogateItem(w, w);
        }

        private void InterrogateItem(Window w, IUIItem i)
        {
            System.Diagnostics.Debug.Print(i.ToString());
            w.Mouse.Location = new System.Windows.Point(i.Bounds.Left + i.Bounds.Width / 2, i.Bounds.Top + i.Bounds.Height / 2);

            if (i is UIItemContainer)
            {
                var ic = i as UIItemContainer;
                foreach (var mb in ic.MenuBars)
                {
                    System.Diagnostics.Debug.Print(mb.ToString());

                    w.Mouse.Location = new System.Windows.Point(mb.Bounds.Left + mb.Bounds.Width / 2, mb.Bounds.Top + mb.Bounds.Height / 2);

                }

                foreach (var x in ic.Items)
                {
                    InterrogateItem(w, x);

                    w.Mouse.Location = new System.Windows.Point(x.Bounds.Left + x.Bounds.Width / 2, x.Bounds.Top + x.Bounds.Height / 2);
                }
            }
        }
    }
}
