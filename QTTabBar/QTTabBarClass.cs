//    This file is part of QTTabBar, a shell extension for Microsoft
//    Windows Explorer.
//    Copyright (C) 2007-2022  Quizo, Paul Accisano, indiff
//
//    QTTabBar is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    QTTabBar is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with QTTabBar.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Media;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using BandObjectLib;
using Microsoft.Win32;
using QTPlugin;
using QTTabBarLib.Interop;
using SHDocVw;
using Timer = System.Windows.Forms.Timer;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using QTTabBarLib.Common;
using Control = System.Windows.Forms.Control;
using IShellFolder = QTTabBarLib.Interop.IShellFolder;
using IShellView = QTTabBarLib.Interop.IShellView;
using ToolTip = System.Windows.Forms.ToolTip;
using System.Management;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Controls.Primitives;

namespace QTTabBarLib {
    /**
     sealed Ӧ����ĳ����ʱ��sealed ���η�����ֹ������̳��Ը��ࡣ�������ʾ���У��� B �̳����� A����û������Լ̳����� B��
     class A {}
     sealed class B : A {}
     */
    [ComVisible(true), Guid("d2bf470e-ed1c-487f-a333-2bd8835eb6ce")]
    public partial class QTTabBarClass : TabBarBase
    {
       
        private VisualStyleRenderer bgRenderer;
        private BreadcrumbBar breadcrumbBar;
        private ToolStripDropDownButton buttonNavHistoryMenu;

        private IContainer components;
        private ContextMenuStripEx contextMenuDropped;
        private QTabItem ContextMenuedTab;

        
        
        private Cursor curTabCloning;
        private Cursor curTabDrag;
        private Rectangle DraggingDestRect;
        private QTabItem DraggingTab;
        private DropTargetWrapper dropTargetWrapper;
        private NativeWindowController explorerController;
        
        private bool fHideExplorer;
        private bool fDrivesContainedDD;
        private static bool fInitialized;
        private readonly bool fIsFirstLoad;
        private volatile bool FirstNavigationCompleted;
        private bool fAutoNavigating;
        
        private bool fNeedsNewWindowPulse;
        private bool fNowQuitting;
        private bool fNowRestoring;
        private bool fNowTravelByTree;
        private bool fToggleTabMenu;
        private IntPtr hHook_Key;
        private IntPtr hHook_Mouse;
        private IntPtr hHook_Msg;
        private HookProc hookProc_GetMsg;
        private HookProc hookProc_Key;
        private HookProc hookProc_Mouse;
        private ShellContextMenu shellContextMenu = new ShellContextMenu();
        private int iModKeyStateDD;
        private const int INTERVAL_SELCTTAB = 5000;
        private const int INTERVAL_SHOWMENU = 0x4b0;
        private int iSequential_WM_CLOSE;
        private bool IsShown;
        private byte[] lastAttemptedBrowseObjectIDL;
        private byte[] lastCompletedBrowseObjectIDL;
        
        private static FileHashComputerForm md5Form;
        private ToolStripTextBox menuTextBoxTabAlias;


        
        private string strDraggingDrive;
        private string strDraggingStartPath;
        private SubDirTipForm subDirTip_Tab;
        public QTabControl tabControl1;
        private QTabItem tabForDD;
        private TabSwitchForm tabSwitcher;
        private Timer timerOnTab;
        private Timer timerSelectionChanged;
        
        private ToolTip toolTipForDD;
        private NativeWindowController travelBtnController;

        private IntPtr TravelToolBarHandle;
        private TreeViewWrapper treeViewWrapper;

        private readonly int WM_NEWTREECONTROL = PInvoke.RegisterWindowMessage("QTTabBar_NewTreeControl");
        private readonly int WM_BROWSEOBJECT = PInvoke.RegisterWindowMessage("QTTabBar_BrowseObject");
        private readonly int WM_HEADERINALLVIEWS = PInvoke.RegisterWindowMessage("QTTabBar_HeaderInAllViews");
        private readonly int WM_LISTREFRESHED = PInvoke.RegisterWindowMessage("QTTabBar_ListRefreshed");
        private readonly int WM_SHOWHIDEBARS = PInvoke.RegisterWindowMessage("QTTabBar_ShowHideBars");
        private readonly int WM_CHECKPULSE = PInvoke.RegisterWindowMessage("QTTabBar_CheckPulse");
        private readonly int WM_SELECTFILE = PInvoke.RegisterWindowMessage("QTTabBar_SelectFile");


        internal bool CanNavigateBackward { get { return ((navBtnsFlag & 1) != 0); } }
        internal bool CanNavigateForward { get { return ((navBtnsFlag & 2) != 0); } }
        internal int TabCount { get { return tabControl1.TabCount; } }
        internal int SelectedTabIndex {
            get {
                return tabControl1.TabPages.IndexOf(CurrentTab);
            }
            set {
                if(0 <= value && value < tabControl1.TabPages.Count) {
                    tabControl1.SelectTab(value);
                }
            }
        }


        #region qwop �Զ�����
        private static QTTabBarClass lstTabBar;

        public static QTTabBarClass GetThreadTabBar()
        {
            return InstanceManager.GetThreadTabBar(); 
        }

        /// <summary>
        /// ����һ����ǰ�̵߳� ·����ǩ
        /// </summary>
        /// <param name="address">·��</param>
        /// <param name="index">λ��</param>
        /// <param name="fLocked">�Ƿ�����</param>
        /// <param name="fSelect">�Ƿ�ѡ��</param>
        /// <returns></returns>
        public static bool CreateTab(QTTabBarClass tabBar, Address address, int index, bool fLocked, bool fSelect)
        {
            if (null == tabBar) {
                tabBar = GetThreadTabBar();    
            }
            
            if (null == tabBar)
            {
                if (lstTabBar == null)
                {
                    return false;
                }
                tabBar = lstTabBar;
            } 
            
            using (IDLWrapper wrapper = new IDLWrapper(address))
            {
                address.ITEMIDLIST = wrapper.IDL;
                address.Path = wrapper.Path;
            }
            if ((address.ITEMIDLIST == null) || (address.ITEMIDLIST.Length <= 0))
            {
                return false;
            }

            QTabItem tab = new QTabItem(QTUtility2.MakePathDisplayText(address.Path, false), address.Path, tabBar.tabControl1);
            tab.NavigatedTo(address.Path, address.ITEMIDLIST, -1, false);
            tab.ToolTipText = QTUtility2.MakePathDisplayText(address.Path, true);
            tab.TabLocked = fLocked;
            if (index < 0)
            {
                tabBar.AddInsertTab(tab);
            }
            else
            {
                if (index > tabBar.tabControl1.TabCount)
                {
                    index = tabBar.tabControl1.TabCount;
                }
                tabBar.tabControl1.TabPages.Insert(index, tab);
            }
            if (fSelect)
            {
                tabBar.tabControl1.SelectTab(tab);
            }
            return true;
        }
        #endregion

        public QTTabBarClass() {
            QTUtility.Initialize();

            try {
                
                string installDateString;
                DateTime installDate;
                string minDate = DateTime.MinValue.ToString();
                using(RegistryKey key = Registry.LocalMachine.OpenSubKey(RegConst.Root)) {
                    installDateString = key == null ? minDate : (string)key.GetValue("InstallDate", minDate);
                    // ʱ���ʽ���� ���ܻᵼ�³�ʼ��ʧ��
                    try
                    {
                        QTUtility2.log("installDateString " + installDateString);
                        installDate = DateTime.Parse(installDateString);
                    }
                    catch (Exception e)
                    {
                        installDate = DateTime.ParseExact(installDateString, "yyyy/MM/dd HH:mm:ss", CultureInfo.CurrentCulture);
                        // ignore exception 
                    }
                }
                using(RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root)) {
                    DateTime lastActivation;
                    // DateTime lastActivation = DateTime.Parse((string)key.GetValue("ActivationDate", minDate));
                    var value = (string)key.GetValue("ActivationDate", minDate);
                    try
                    {
                        QTUtility2.log("ActivationDate " + value);
                        lastActivation = DateTime.Parse(value);
                    }
                    catch (Exception e)
                    {
                        lastActivation = DateTime.ParseExact(value, "yyyy/MM/dd HH:mm:ss", CultureInfo.CurrentCulture);
                        // ignore exception 
                    }

                    fIsFirstLoad = installDate.CompareTo(lastActivation) > 0;
                    // ʱ���ʽ���� ���ܻᵼ�³�ʼ��ʧ��
                    if (fIsFirstLoad)
                        key.SetValue("ActivationDate", installDateString);
                }
            }
            catch (Exception e ){
                QTUtility2.MakeErrorLog(e, "QTTabBarClass ���캯����ʼ����װʱ��");
            }
            if(!fInitialized) {
                InitializeStaticFields();
            }
            // ��ʼ���߶�
            BandHeight = Config.Skin.TabHeight + BandHeightSpace;
            // BandHeight = Config.Skin.TabHeight + 10;
            InitializeComponent();
            lstActivatedTabs.Add(CurrentTab);

            // reocrd the last qttabbarclass instance, add by indiff .
            lstTabBar = this;
            
            // Ĭ�ϻ�ȡ�Ƿ�������־
            QTUtility2.ENABLE_LOGGER = Config.Misc.EnableLog;
        }

        private void AddInsertTab(QTabItem tab) {
            QTUtility2.log(  "QTTabBarClass AddInsertTab  " );
            switch(Config.Tabs.NewTabPosition) {
                case TabPos.Leftmost:
                    tabControl1.TabPages.Insert(0, tab);
                    break;

                case TabPos.Right:
                case TabPos.Left: {
                    int index = tabControl1.TabPages.IndexOf(CurrentTab);
                    if(index == -1) {
                        tabControl1.TabPages.Add(tab);
                    }
                    else {
                        tabControl1.TabPages.Insert(Config.Tabs.NewTabPosition == TabPos.Right ? (index + 1) : index, tab);    
                    }
                    break;
                }

                default: // TabPos.Rightmost
                    tabControl1.TabPages.Add(tab);
                    break;
            }
        }

        private void AddStartUpTabs(string openingGRP, string openingPath) {
            QTUtility2.log(  "QTTabBarClass AddStartUpTabs openingGRP "  + openingGRP + " openingPath " + openingPath);
            if(ModifierKeys == Keys.Shift || InstanceManager.GetTotalInstanceCount() != 0) return;
            if(Config.Window.RestoreOnlyLocked) {
                RestoreTabsOnInitialize(1, openingPath);
            }
            else if(Config.Window.RestoreSession || fIsFirstLoad) {
                RestoreTabsOnInitialize(0, openingPath);
            }
        }

        // TODO: Kill this.
        private void AsyncComplete_FolderTree(IAsyncResult ar) {
            AsyncResult result = (AsyncResult)ar;
            ((WaitTimeoutCallback)result.AsyncDelegate).EndInvoke(ar);
            if(IsHandleCreated) {
                Invoke(new FormMethodInvoker(CallbackFolderTree), new object[] { result.AsyncState });
            }
        }

        // This function is used as a more available version of BeforeNavigate2.
        // Return true to suppress the navigation.  Target IDL should not be relied
        // upon; it's not guaranteed to be accurate.
        private bool BeforeNavigate(IDLWrapper target, bool autonav) {
            if(!IsShown) return false;
            HideSubDirTip_Tab_Menu();
            NowTabDragging = false;
            fAutoNavigating = autonav;
            if(!NavigatedByCode) {
                SaveSelectedItems(CurrentTab);
            }
            if(NowInTravelLog) {
                if(CurrentTravelLogIndex > 0) {
                    CurrentTravelLogIndex--;
                    if(!IsSpecialFolderNeedsToTravel(target.Path)) {
                        NavigateBackToTheFuture();
                    }
                }
                else {
                    NowInTravelLog = false;
                }
            }
            lastAttemptedBrowseObjectIDL = target.IDL;
            return false;
        }

        private void CallbackFolderTree(object obj) {
            bool fShow = (bool)obj;
            ShowFolderTree(fShow);
            if(fShow) {
                PInvoke.SetRedraw(ExplorerHandle, true);
                PInvoke.RedrawWindow(ExplorerHandle, IntPtr.Zero, IntPtr.Zero, 0x289);
            }
        }

        private IntPtr CallbackGetMsgProc(int nCode, IntPtr wParam, IntPtr lParam) {
            if(nCode >= 0) {
                MSG msg = (MSG)Marshal.PtrToStructure(lParam, typeof(MSG));
                // QTUtility2.debugMessage(msg);
                try {
                    if(QTUtility.IsXP) {
                        if(msg.message == WM.CLOSE) {
                            if(iSequential_WM_CLOSE > 0) {
                                Marshal.StructureToPtr(new MSG(), lParam, false);
                                return PInvoke.CallNextHookEx(hHook_Msg, nCode, wParam, lParam);
                            }
                            iSequential_WM_CLOSE++;
                        }
                        else {
                            iSequential_WM_CLOSE = 0;
                        }
                    }

                    if(msg.message == WM_NEWTREECONTROL)
                    {
                        QTUtility2.log("CallbackGetMsgProc WM_NEWTREECONTROL");
                        object obj = Marshal.GetObjectForIUnknown(msg.wParam);
                        try {
                            if(obj != null) {
                                IOleWindow window = obj as IOleWindow;
                                if(window != null) {
                                    IntPtr hwnd;
                                    window.GetWindow(out hwnd);
                                    if(hwnd != IntPtr.Zero && PInvoke.IsChild(ExplorerHandle, hwnd)) {
                                        hwnd = WindowUtils.FindChildWindow(hwnd,
                                                child => PInvoke.GetClassName(child) == "SysTreeView32");
                                        if(hwnd != IntPtr.Zero) {
                                            INameSpaceTreeControl control = obj as INameSpaceTreeControl;
                                            if(control != null) {
                                                if(treeViewWrapper != null) {
                                                    treeViewWrapper.Dispose();
                                                }
                                                treeViewWrapper = new TreeViewWrapper(hwnd, control);
                                                treeViewWrapper.TreeViewClicked += FolderLinkClicked;
                                                QTUtility2.log("CallbackGetMsgProc regedit TreeViewClicked");
                                                obj = null; // Release the object only if we didn't get this far.
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        finally {
                            if(obj != null) {
                                QTUtility2.log("ReleaseComObject obj");
                                Marshal.ReleaseComObject(obj);
                            }
                        }
                        return PInvoke.CallNextHookEx(hHook_Msg, nCode, wParam, lParam);   
                    }
                    else if(msg.message == WM_LISTREFRESHED) {
                        HandleF5();
                        return PInvoke.CallNextHookEx(hHook_Msg, nCode, wParam, lParam);
                    }
                    else if (msg.message == WM_SELECTFILE)
                    {
                        QTUtility2.log(" select file 1 " + " wparam " + wParam + " lparam " + lParam );
                        // lParam: point to an array of ITEMIDLISTs;
                        /*IShellView shellView = null;
                        if (0 == ShellBrowser.GetIShellBrowser().QueryActiveShellView(out shellView))
                        {
                            // new Guid("{0000010e-0000-0000-C000-000000000046}")
                            var iid = new Guid("{0000010e-0000-0000-C000-000000000046}");
                            object ppv;
                            shellView.GetItemObject((uint)SVSIF.SELECT, ref iid, out ppv);
                            if (ppv != null)
                            {
                                System.Runtime.InteropServices.ComTypes.IDataObject pDataObject = (System.Runtime.InteropServices.ComTypes.IDataObject)ppv;
                                // var iid = new Guid(ShellIIDGuid.IShellItemArray);
                                var shellObjectCollection = ShellObjectCollection.FromDataObject(pDataObject);
                                if (shellObjectCollection.Count>0)
                                {
                                    foreach (ShellObject so in shellObjectCollection)
                                    {
                                        QTUtility2.log("so.Name " + so.Name);
                                    }
                                }
                            }
                        }*/
                        // IntPtr pIDL = Marshal.ReadIntPtr(lParam);
                        // SFVCB_SELECTINFO sel = (SFVCB_SELECTINFO)Marshal.PtrToStructure(lParam, typeof(SFVCB_SELECTINFO));
                        // QTUtility2.log("SFVCB_SELECTINFO UOldState " + sel.UOldState);
                        // QTUtility2.log("SFVCB_SELECTINFO UNewState " + sel.UNewState);
                        // QTUtility2.log("SFVCB_SELECTINFO Pid1 " + sel.Pid1);
                        return PInvoke.CallNextHookEx(hHook_Msg, nCode, wParam, lParam);
                    }

                    switch(msg.message) {
                        /* TODO: Handle FolderView clicks on XP.
                        case WM.LBUTTONDOWN:
                        case WM.LBUTTONUP:
                            if((QTUtility.IsXP && !Config.NoMidClickTree) && ((((int)((long)msg.wParam)) & 4) != 0)) {
                                HandleLBUTTON_Tree(msg, msg.message == 0x201);
                            }
                            break;

                        
                        */
                        case WM.MBUTTONUP:
                            if (!Explorer.Busy) // && !Config.NoMidClickTree
                            {
                                QTUtility2.log("CallbackGetMsgProc MBUTTONUP NoMidClickTree");
                                /*
                                object obj = Marshal.GetObjectForIUnknown(msg.wParam);
                                try {
                                    if(obj != null) {
                                        IOleWindow window = obj as IOleWindow;
                                        if(window != null) {
                                            IntPtr hwnd;
                                            window.GetWindow(out hwnd);
                                            if(hwnd != IntPtr.Zero && PInvoke.IsChild(ExplorerHandle, hwnd)) {
                                                hwnd = WindowUtils.FindChildWindow(hwnd,
                                                        child => PInvoke.GetClassName(child) == "SysTreeView32");
                                                if(hwnd != IntPtr.Zero) {
                                                    INameSpaceTreeControl control = obj as INameSpaceTreeControl;
                                                    if(control != null) {
                                                        if(treeViewWrapper != null) {
                                                            treeViewWrapper.Dispose();
                                                        }
                                                        treeViewWrapper = new TreeViewWrapper(hwnd, control);
                                                        treeViewWrapper.TreeViewClicked += FolderLinkClicked;
                                                        QTUtility2.log("CallbackGetMsgProc regedit TreeViewClicked");
                                                        obj = null; // Release the object only if we didn't get this far.
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                finally {
                                    if(obj != null) {
                                        Marshal.ReleaseComObject(obj);
                                    }
                                }
                                */
                                Handle_MButtonUp_Tree(msg);
                            }
                            break;
                        case WM.SYSCOLORCHANGE:
                            QTUtility.InNightMode = QTUtility.getNightMode();
                            QTUtility2.log("SYSCOLORCHANGE SwitchNighMode");
                            Config.Skin.SwitchNightMode(QTUtility.InNightMode ); // ����ر��Զ���ɫ�򲻽��б�ɫ
                            ConfigManager.UpdateConfig(true);
                            this.tabControl1.InitializeColors();
                            PInvoke.SetRedraw(ExplorerHandle, true);
                            PInvoke.RedrawWindow(ExplorerHandle, IntPtr.Zero, IntPtr.Zero, 0x289);
                            break;

                        case WM.CLOSE:  // �رմ���
                            if(QTUtility.IsXP) {
                                if((msg.hwnd == ExplorerHandle) && HandleCLOSE(msg.lParam)) {
                                    Marshal.StructureToPtr(new MSG(), lParam, false);
                                }
                                break;
                            }

                            
                            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root))
                            {
                                string[] list = (from QTabItem item2 in tabControl1.TabPages
                                                 where item2.TabLocked
                                                 select item2.CurrentPath).ToArray();
                               // MessageBox.Show(String.Join(",", list));
                                QTUtility2.WriteRegBinary(list, "TabsLocked", key);
                            }
                            if(msg.hwnd == WindowUtils.GetShellTabWindowClass(ExplorerHandle)) { // �����ǩ�� handle ����Դ��������ƥ��
                                try {
                                    bool flag = tabControl1.TabCount == 1;
                                    string currentPath = tabControl1.SelectedTab.CurrentPath;
                                    if(!Directory.Exists(currentPath) && // �����ǰ·��Ŀ¼������
                                       currentPath.Length > 3 
                                       /* && currentPath.Substring(1, 2) == @":\" */ ) {
                                        if(flag) {
                                            WindowUtils.CloseExplorer(ExplorerHandle, 2);
                                        }
                                        else {
                                            CloseTab(tabControl1.SelectedTab, true);
                                        }
                                    }
                                }
                                catch(Exception e ) {
                                    QTUtility2.MakeErrorLog(e, "CallbackGetMsgProc WM.Close");
                                }
                                Marshal.StructureToPtr(new MSG(), lParam, false);
                            }
                            break;

                        case WM.COMMAND:
                            if(QTUtility.IsXP) {
                                int num = ((int)((long)msg.wParam)) & 0xffff;
                                if(num == 0xa021) {
                                    WindowUtils.CloseExplorer(ExplorerHandle, 3);
                                    Marshal.StructureToPtr(new MSG(), lParam, false);
                                }
                            }
                            break;

                        /*case 48648:  // not work
                            QTUtility2.log("48648");
                            break;*/
                    }
                }
                catch(Exception ex) {
                    QTUtility2.MakeErrorLog(ex, String.Format("Message: {0:x4}", msg.message));
                }
            }
            return PInvoke.CallNextHookEx(hHook_Msg, nCode, wParam, lParam);
        }


       

        /**
         * ע���ݼ��Ļص�
         */
        private IntPtr CallbackKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam) {
            const uint KB_TRANSITION_FLAG = 0x80000000;
            const uint KB_PREVIOUS_STATE_FLAG = 0x40000000;
            if(nCode < 0 || NowModalDialogShown) {
                return PInvoke.CallNextHookEx(hHook_Key, nCode, wParam, lParam);
            }

            try {
                uint flags = (uint)((long)lParam);
                bool isKeyPress = (flags & KB_TRANSITION_FLAG) == 0;
                bool isRepeat = (flags & KB_PREVIOUS_STATE_FLAG) != 0;
                Keys key = (Keys)((int)wParam);

                if(key == Keys.ShiftKey) {
                    if(isKeyPress || !isRepeat) {
                        listView.HandleShiftKey();
                    }
                }

                if(isKeyPress) {
                    if(HandleKEYDOWN(key, isRepeat)) {
                        return new IntPtr(1);
                    }
                }
                else {
                    listView.HideThumbnailTooltip(3);
                    if(NowTabDragging && DraggingTab != null) {
                        Cursor = Cursors.Default;
                    }

                    switch(key) {
                        case Keys.ControlKey:
                            if(Config.Keys.UseTabSwitcher) {
                                HideTabSwitcher(true);
                            }
                            break;

                        case Keys.Menu: // Alt key
                            if(Config.Tabs.ShowCloseButtons && Config.Tabs.CloseBtnsWithAlt) {
                                tabControl1.ShowCloseButton(false);
                            }
                            break;

                        case Keys.Tab:
                            if(Config.Keys.UseTabSwitcher && tabSwitcher != null && tabSwitcher.IsShown) {
                                tabControl1.SetPseudoHotIndex(tabSwitcher.SelectedIndex);
                            }
                            break;
                    }
                }
            }
            catch(Exception ex) {
                QTUtility2.MakeErrorLog(ex,
                        String.Format("LParam: {0:x4}, WParam: {1:x4}", (long)lParam, (long)wParam));
            }
            return PInvoke.CallNextHookEx(hHook_Key, nCode, wParam, lParam);
        }

        private IntPtr CallbackMouseProc(int nCode, IntPtr wParam, IntPtr lParam) {
            try {
                if(nCode >= 0 && !NowModalDialogShown) {
                    IntPtr ptr = (IntPtr)1;
                    switch(((int)wParam)) {
                        case WM.MOUSEWHEEL:
                            if(!HandleMOUSEWHEEL(lParam)) {
                                break;
                            }
                            return ptr;

                        /*case WM.MBUTTONDOWN: // add by indiff
                            QTUtility2.log("CallbackMouseProc Handle_MButtonUp_Tree");
                            if (Handle_MButtonUp_Tree(ExplorerHandle,lParam))
                            {
                                break;
                            }
                            return ptr;*/

                        case WM.XBUTTONDOWN:
                             
                        case WM.XBUTTONUP:
                            MouseButtons mouseButtons = MouseButtons;
                            Keys modifierKeys = ModifierKeys;
                            MouseChord chord = mouseButtons == MouseButtons.XButton1
                                    ? MouseChord.X1
                                    : mouseButtons == MouseButtons.XButton2 ? MouseChord.X2 : MouseChord.None;
                            if(chord == MouseChord.None) break;
                            chord = QTUtility.MakeMouseChord(chord, modifierKeys);
                            BindAction action;
                            if(!Config.Mouse.GlobalMouseActions.TryGetValue(chord, out action)) {
                                break;
                            }
                            if(((int)wParam) == WM.XBUTTONUP && !Explorer.Busy) {
                                QTUtility2.log("QTTabBarClass WM.XBUTTONUP " + action);
                                DoBindAction(action);
                            }
                            return ptr;
                    }
                }
            }
            catch(Exception ex) {
                QTUtility2.MakeErrorLog(ex, String.Format("LParam: {0:x4}, WParam: {1:x4}", (long)lParam, (long)wParam));
            }
            return PInvoke.CallNextHookEx(hHook_Mouse, nCode, wParam, lParam);
        }

        private void CallbackMultiPath(object obj) {
            object[] objArray = (object[])obj;
            string[] collection = (string[])objArray[0];
            int num = (int)objArray[1];
            switch(num) {
                case 0:
                    foreach(string str in collection) {
                        OpenNewTab(str, true);
                    }
                    break;

                case 1: {
                        bool flag = true;
                        foreach(string str2 in collection) {
                            OpenNewTab(str2, !flag);
                            flag = false;
                        }
                        break;
                    }
                default:
                    StaticReg.CreateWindowPaths.Assign(collection);
                    using(IDLWrapper wrapper = new IDLWrapper(collection[0])) {
                        OpenNewWindow(wrapper);
                    }
                    break;
            }
            if(num == 1) {
                InstanceManager.RemoveFromTrayIcon(Handle);
                WindowUtils.BringExplorerToFront(ExplorerHandle);
            }
        }

        private void CancelFailedNavigation(string failedPath, bool fRollBackForward, int countRollback) {
            ShowMessageNavCanceled(failedPath, false);
            if(fRollBackForward) {
                for(int i = 0; i < countRollback; i++) {
                    CurrentTab.GoForward();
                }
            }
            else {
                for(int j = 0; j < countRollback; j++) {
                    CurrentTab.GoBackward();
                }
            }
            NavigatedByCode = false;
        }

        

        private void ChangeViewMode(bool fUp) {
            FVM orig = ShellBrowser.ViewMode;
            FVM mode = orig;
            switch(mode) {
                case FVM.ICON:
                    mode = fUp ? FVM.TILE : FVM.LIST;
                    break;

                case FVM.LIST:
                    mode = fUp ? FVM.ICON : FVM.DETAILS;
                    break;

                case FVM.DETAILS:
                    if(fUp) {
                        mode = FVM.LIST;
                    }
                    break;

                case FVM.THUMBNAIL:
                    mode = fUp ? FVM.THUMBSTRIP : FVM.TILE;
                    break;

                case FVM.TILE:
                    mode = fUp ? FVM.THUMBNAIL : FVM.ICON;
                    break;

                case FVM.THUMBSTRIP:
                    if(!fUp) {
                        mode = FVM.THUMBNAIL;
                    }
                    break;
            }
            if(mode != orig) {
                ShellBrowser.ViewMode = mode;
            }
        }

        private static bool CheckProcessID(IntPtr hwnd1, IntPtr hwnd2) {
            uint num;
            uint num2;
            PInvoke.GetWindowThreadProcessId(hwnd1, out num);
            PInvoke.GetWindowThreadProcessId(hwnd2, out num2);
            return ((num == num2) && (num != 0));
        }

        private void ChooseNewDirectory() {
            NowModalDialogShown = true;
            bool nowTopMost = NowTopMost;
            if(nowTopMost) {
                ToggleTopMost();
            }
			// ��ʾ����Ŀ¼��ť
            using(FolderBrowserDialog dialog = new FolderBrowserDialog()) {
                dialog.ShowNewFolderButton = true; // add by indiff
                dialog.SelectedPath = CurrentAddress;

                //IntPtr hWndInsertAfter = false ? ((IntPtr)(-1)) : ((IntPtr)(-2));
                //PInvoke.SetWindowPos(Handle, hWndInsertAfter, 100, 100, 0, 0, 0x53);


                if(DialogResult.OK == dialog.ShowDialog()) {
                    OpenNewTab(dialog.SelectedPath);
                }
            }
            NowModalDialogShown = false;
            if(nowTopMost) {
                ToggleTopMost();
            }
        }

        private void ClearTravelLogs() {
            IEnumTravelLogEntry ppenum = null;
            try {
                if((TravelLog.EnumEntries(0x30, out ppenum) != 0) || (ppenum == null)) {
                    return;
                }
                int num = 0;
            Label_0018:
                ITravelLogEntry entry2 = null;
                try {
                    if(ppenum.Next(1, out entry2, 0) == 0) {
                        IntPtr ptr;
                        if((num++ != 0) && (entry2.GetURL(out ptr) == 0)) {
                            string path = Marshal.PtrToStringUni(ptr);
                            PInvoke.CoTaskMemFree(ptr);
                            if(!IsSpecialFolderNeedsToTravel(path)) {
                                TravelLog.RemoveEntry(entry2);
                            }
                        }
                        goto Label_0018;
                    }
                }
                finally {
                    if(entry2 != null) {
                        QTUtility2.log("ReleaseComObject entry2");
                        Marshal.ReleaseComObject(entry2);
                    }
                }
            }
            catch(Exception exception) {
                QTUtility2.MakeErrorLog(exception);
            }
            finally {
                if(ppenum != null) {
                    QTUtility2.log("ReleaseComObject ppenum");
                    Marshal.ReleaseComObject(ppenum);
                }
            }
        }

        internal void CloneCurrentTab(bool fSelect = true) {
            CloneTabButton(CurrentTab, null, fSelect, -1);
        }

        private void CloneTabButton(QTabItem tab, LogData log) {
            NowTabCloned = true;
            QTabItem item = tab.Clone();
            AddInsertTab(item);
            using(IDLWrapper wrapper = new IDLWrapper(log.IDL)) {
                if(wrapper.Available) {
                    item.NavigatedTo(wrapper.Path, wrapper.IDL, log.Hash, false);
                }
            }
            tabControl1.SelectTab(item);
        }

        private QTabItem CloneTabButton(QTabItem tab, string optionURL, bool fSelect, int index)
        {
            QTUtility2.log("QTTabBarLib.QTTabBarClass.CloneTabButton optionURL " + optionURL +
                            " fSelect " + fSelect + 
                            " index " + index 
            );
            NowTabCloned = fSelect;
            QTabItem item = tab.Clone();
            if(index < 0) {
                AddInsertTab(item);
            }
            else if((-1 < index) && (index < (tabControl1.TabCount + 1))) {
                tabControl1.TabPages.Insert(index, item);
            }
            else {
                AddInsertTab(item);
            }
            if(optionURL != null) {
                using(IDLWrapper wrapper = new IDLWrapper(optionURL)) {
                    item.NavigatedTo(optionURL, wrapper.IDL, -1, false);
                }
            }
            if(fSelect) {
                tabControl1.SelectTab(item);
            }
            else {
                item.RefreshRectangle();
                tabControl1.Refresh();
            }
            return item;
        }

        private List<string> CloseAllTabsExcept(QTabItem leaveThisOne, bool leaveLocked = true) {
            List<QTabItem> tabs = tabControl1.TabPages.Where(item => 
                !(leaveLocked && item.TabLocked) && item != leaveThisOne).ToList();
            List<string> paths = tabs.Select(tab => tab.CurrentPath).ToList();
            CloseTabs(tabs, !leaveLocked);
            return paths;
        }

        /**
         *����رմ����¼� by indiff
         */
        public override void CloseDW(uint dwReserved) {
            try {
                string[] list = (from QTabItem item2 in tabControl1.TabPages
                                 where item2.TabLocked
                                 select item2.CurrentPath).ToArray();
                if(treeViewWrapper != null) {
                    treeViewWrapper.Dispose();
                    treeViewWrapper = null;
                }
                if(listViewManager != null) {
                    listViewManager.Dispose();
                    listViewManager = null;
                }
                if(subDirTip_Tab != null) {
                    subDirTip_Tab.Dispose();
                    subDirTip_Tab = null;
                }
                if(IsShown) {
                    if(hHook_Key != IntPtr.Zero) {
                        PInvoke.UnhookWindowsHookEx(hHook_Key);
                        hHook_Key = IntPtr.Zero;
                    }
                    if(hHook_Mouse != IntPtr.Zero) {
                        PInvoke.UnhookWindowsHookEx(hHook_Mouse);
                        hHook_Mouse = IntPtr.Zero;
                    }
                    if(hHook_Msg != IntPtr.Zero) {
                        PInvoke.UnhookWindowsHookEx(hHook_Msg);
                        hHook_Msg = IntPtr.Zero;
                    }
                    if(explorerController != null) {
                        explorerController.ReleaseHandle();
                        explorerController = null;
                    }
                    if(rebarController != null) {
                        rebarController.Dispose();
                        rebarController = null;
                    }
                    if(!QTUtility.IsXP && (travelBtnController != null)) {
                        travelBtnController.ReleaseHandle();
                        travelBtnController = null;
                    }

                    // �Ƴ����½�ͼ��
                    if (null != Handle)
                    {
                        InstanceManager.RemoveFromTrayIcon(Handle);
                    }

                    // TODO: check this
                    using(RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root)) {
                        if(Config.Misc.KeepHistory) {
                            foreach(QTabItem item in tabControl1.TabPages) {
                                AddToHistory(item);
                            }
                            QTUtility.SaveRecentlyClosed(key);
                        }
                        if(Config.Misc.KeepRecentFiles) {
                            QTUtility.SaveRecentFiles(key);
                        }

                        /*foreach (QTabItem item in tabControl1.TabPages)
                        {
                            if ( item.TabLocked ) {
                                MessageBox.Show(item.CurrentPath);
                            }
                        }*/

                        // �ر���Ϣȥ��д��������ǩ�ĵ���
                        if (list != null && list.Length > 0)
                        {
                            QTUtility2.WriteRegBinary(list, "TabsLocked", key);
                        }

                        InstanceManager.UnregisterTabBar();
                        if(0x80000 != ((int)PInvoke.Ptr_OP_AND(PInvoke.GetWindowLongPtr(ExplorerHandle, -20), 0x80000))) {
                            QTUtility.WindowAlpha = 0xff;
                        }
                        else {
                            byte num;
                            int num2;
                            int num3;
                            if(PInvoke.GetLayeredWindowAttributes(ExplorerHandle, out num2, out num, out num3)) {
                                QTUtility.WindowAlpha = num;
                            }
                            else {
                                QTUtility.WindowAlpha = 0xff;
                            }
                        }
                        key.SetValue("WindowAlpha", QTUtility.WindowAlpha);
                        IDLWrapper.SaveCache(key);
                    }
                    if((md5Form != null) && !md5Form.InvokeRequired) {
                        md5Form.SaveMD5FormStat();
                        md5Form.Dispose();
                        md5Form = null;
                    }
                    Cursor = Cursors.Default;
                    if((curTabDrag != null) && (curTabDrag != Cursors.Default)) {
                        PInvoke.DestroyIcon(curTabDrag.Handle);
                        GC.SuppressFinalize(curTabDrag);
                        curTabDrag = null;
                    }
                    if((curTabCloning != null) && (curTabCloning != Cursors.Default)) {
                        PInvoke.DestroyIcon(curTabCloning.Handle);
                        GC.SuppressFinalize(curTabCloning);
                        curTabCloning = null;
                    }
                    if(dropTargetWrapper != null) {
                        dropTargetWrapper.Dispose();
                        dropTargetWrapper = null;
                    }
                    OptionsDialog.ForceClose();
                    if(tabSwitcher != null) {
                        tabSwitcher.Dispose();
                        tabSwitcher = null;
                    }
                }
                if(TravelLog != null) {
                    QTUtility2.log("ReleaseComObject TravelLog");
                    Marshal.FinalReleaseComObject(TravelLog);
                    TravelLog = null;
                }
                if(shellContextMenu != null) {
                    shellContextMenu.Dispose();
                    shellContextMenu = null;
                }
                if(ShellBrowser != null) {
                    ShellBrowser.Dispose();
                    ShellBrowser = null;
                }
                foreach(ITravelLogEntry entry in LogEntryDic.Values) {
                    if(entry != null) {
                        QTUtility2.log("ReleaseComObject entry");
                        Marshal.FinalReleaseComObject(entry);
                    }
                }
                LogEntryDic.Clear();
                fFinalRelease = true;
                base.CloseDW(dwReserved);
            }
            catch(Exception exception2) {
                QTUtility2.MakeErrorLog(exception2, "tabbar closing");
            }
        }

        private void CloseLeftRight(bool fLeft, int index) {
            if(index == -1) {
                index = tabControl1.SelectedIndex;
            }
            if(fLeft ? (index <= 0) : (index >= (tabControl1.TabCount - 1))) return;
            CloseTabs(fLeft
                    ? tabControl1.TabPages.Take(index).ToList()
                    : tabControl1.TabPages.Skip(index + 1).ToList());
        }
        


      

        private void contextMenuSys_ItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            if(e.ClickedItem == tsmiOption) {
                OptionsDialog.Open();
            }
            else if(e.ClickedItem == tsmiCloseAllButCurrent) {
                if(tabControl1.TabCount != 1) {
                    CloseAllTabsExcept(CurrentTab);
                }
            }
            else if(e.ClickedItem == tsmiBrowseFolder) {
                ChooseNewDirectory();
            }
            else if(e.ClickedItem == tsmiCloseWindow) {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root))
                {
                    string[] list = (from QTabItem item2 in tabControl1.TabPages
                                     where item2.TabLocked
                                     select item2.CurrentPath).ToArray();

                    // MessageBox.Show(String.Join(",", list));
                    QTUtility2.WriteRegBinary(list, "TabsLocked", key);
                }
                WindowUtils.CloseExplorer(ExplorerHandle, 1);
            }
            else {
                if(e.ClickedItem == tsmiLastActiv) {
                    try {
                        tabControl1.SelectTab(lstActivatedTabs[lstActivatedTabs.Count - 2]);
                        return;
                    }
                    catch (Exception ex)
                    {
                        QTUtility2.MakeErrorLog(ex, "tabControl1.SelectTab");
                        return;
                    }
                }
                if(e.ClickedItem == tsmiLockToolbar) {
                    rebarController.Locked = !tsmiLockToolbar.Checked;
                }
                else if(e.ClickedItem == tsmiMergeWindows) {
                    MergeAllWindows();
                }
            }
        }
        
        // �ؼ����ڴ�ʱ����
        private void contextMenuSys_Opening(object sender, CancelEventArgs e) {
            InitializeSysMenu(false);
            // ��ʱ����ؼ�����
            contextMenuSys.SuspendLayout();
            tsmiUndoClose.DropDown.SuspendLayout();

            MenuUtility.CreateUndoClosedItems(tsmiUndoClose);
            if((lstActivatedTabs.Count > 1) && tabControl1.TabPages.Contains(lstActivatedTabs[lstActivatedTabs.Count - 2])) {
                tsmiLastActiv.ToolTipText = lstActivatedTabs[lstActivatedTabs.Count - 2].CurrentPath;
                tsmiLastActiv.Enabled = true;
            }
            else {
                tsmiLastActiv.ToolTipText = string.Empty;
                tsmiLastActiv.Enabled = false;
            }
            while(tsmiExecuted.DropDownItems.Count > 0) {
                tsmiExecuted.DropDownItems[0].Dispose();
            }
            List<ToolStripItem> list = MenuUtility.CreateRecentFilesItems();
            if(list.Count > 0) {
                tsmiExecuted.DropDown.SuspendLayout();
                tsmiExecuted.DropDownItems.AddRange(list.ToArray());
                tsmiExecuted.DropDown.ResumeLayout();
            }
            tsmiExecuted.Enabled = tsmiExecuted.DropDownItems.Count > 0;
            tsmiMergeWindows.Enabled = InstanceManager.GetTotalInstanceCount() > 1;
            tsmiLockToolbar.Checked = rebarController.Locked;
            if((lstPluginMenuItems_Sys != null) && (lstPluginMenuItems_Sys.Count > 0)) {
                foreach(ToolStripItem item in lstPluginMenuItems_Sys) {
                    item.Dispose();
                }
                lstPluginMenuItems_Sys = null;
            }
            tsmiUndoClose.DropDown.ResumeLayout();
            contextMenuSys.ResumeLayout();
        }

        private void contextMenuTab_Closed(object sender, ToolStripDropDownClosedEventArgs e) {
            tabControl1.SetContextMenuState(false);
            if(ContextMenuedTab != CurrentTab) {
                tabControl1.Refresh();
            }
        }

        private void contextMenuTab_ItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            if(ContextMenuedTab != null) {
                if(e.ClickedItem == tsmiClose) {
                    if(tabControl1.TabCount == 1) {
                        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root))
                        {
                            string[] list = (from QTabItem item2 in tabControl1.TabPages
                                             where item2.TabLocked
                                             select item2.CurrentPath).ToArray();

                            // MessageBox.Show(String.Join(",", list));
                            QTUtility2.WriteRegBinary(list, "TabsLocked", key);
                        }
                        WindowUtils.CloseExplorer(ExplorerHandle, 1);
                    }
                    else {
                        CloseTab(ContextMenuedTab);
                    }
                }
                else if(e.ClickedItem == tsmiCloseAllButThis) {
                    CloseAllTabsExcept(ContextMenuedTab);
                }
                else if(e.ClickedItem == tsmiCloseLeft) {
                    int index = tabControl1.TabPages.IndexOf(ContextMenuedTab);
                    if(index > 0) {
                        CloseLeftRight(true, index);
                    }
                }
                else if(e.ClickedItem == tsmiCloseRight) {
                    int num2 = tabControl1.TabPages.IndexOf(ContextMenuedTab);
                    if(num2 >= 0) {
                        CloseLeftRight(false, num2);
                    }
                }
                else if(e.ClickedItem == tsmiLockThis) {
                    ContextMenuedTab.TabLocked = !ContextMenuedTab.TabLocked;
                }
                else if(e.ClickedItem == tsmiCloneThis) {
                    CloneTabButton(ContextMenuedTab, null, true, -1);
                }
                else if(e.ClickedItem == tsmiCreateWindow) {
                    using(IDLWrapper wrapper = new IDLWrapper(ContextMenuedTab.CurrentIDL)) {
                        OpenNewWindow(wrapper);
                    }
                    if(/*!Config.KeepOnSeparate != */ ((ModifierKeys & Keys.Shift) != Keys.None)) {
                        CloseTab(ContextMenuedTab);
                    }
                }
                else if(e.ClickedItem == tsmiCopy) {
                    string currentPath = ContextMenuedTab.CurrentPath;
                    if(currentPath.IndexOf("???") != -1) {
                        currentPath = currentPath.Substring(0, currentPath.IndexOf("???"));
                    }
                    else if(currentPath.IndexOf("*?*?*") != -1) {
                        currentPath = currentPath.Substring(0, currentPath.IndexOf("*?*?*"));
                    }
                    QTUtility2.SetStringClipboard(currentPath);
                }
                else if(e.ClickedItem == tsmiProp) {
                    ShellMethods.ShowProperties(ContextMenuedTab.CurrentIDL);
                }
                else if (e.ClickedItem == tsmiOpenCmd) { // add by qwop.
                    OpenCmd( null );
                } else if (e.ClickedItem == enableApiHook)
                {
                    EnableApiHook();
                }
            }
        }

        /**
         * ������������ API Hook
         */
        private void EnableApiHook()
        {
            // Create and enable the API hooks
            HookLibManager.Initialize();
            QTUtility2.log("QTUtility ������������ API hooks");
        }

        private void contextMenuTab_Opening(object sender, CancelEventArgs e) {
            try {
                InitializeTabMenu(false);
                int index = tabControl1.TabPages.IndexOf(ContextMenuedTab);
                if((index == -1) || (ContextMenuedTab == null)) {
                    e.Cancel = true;
                }
                else {
                    tabControl1.SetContextMenuState(true);
                    contextMenuTab.SuspendLayout();
                    if(tabControl1.TabCount == 1) {
                        tsmiTabOrder.Enabled = tsmiCloseAllButThis.Enabled = tsmiCloseLeft.Enabled = tsmiCloseRight.Enabled = false;
                    }
                    else {
                        if(index == 0) {
                            tsmiCloseLeft.Enabled = false;
                            tsmiCloseRight.Enabled = true;
                        }
                        else if(index == (tabControl1.TabCount - 1)) {
                            tsmiCloseLeft.Enabled = true;
                            tsmiCloseRight.Enabled = false;
                        }
                        else {
                            tsmiCloseLeft.Enabled = tsmiCloseRight.Enabled = true;
                        }
                        tsmiTabOrder.Enabled = tsmiCloseAllButThis.Enabled = true;
                    }
                    tsmiClose.Enabled = !ContextMenuedTab.TabLocked;
                    tsmiLockThis.Text = ContextMenuedTab.TabLocked ? QTUtility.ResMain[20] : QTUtility.ResMain[6];
                    tsmiHistory.DropDown.SuspendLayout();
                    while(tsmiHistory.DropDownItems.Count > 0) {
                        tsmiHistory.DropDownItems[0].Dispose();
                    }
                    if((ContextMenuedTab.HistoryCount_Back + ContextMenuedTab.HistoryCount_Forward) > 1) {
                        tsmiHistory.DropDownItems.AddRange(CreateNavBtnMenuItems(false).ToArray());
                        tsmiHistory.DropDownItems.AddRange(CreateBranchMenu(false, components, tsmiBranchRoot_DropDownItemClicked).ToArray());
                        tsmiHistory.Enabled = true;
                    }
                    else {
                        tsmiHistory.Enabled = false;
                    }
                    tsmiHistory.DropDown.ResumeLayout();
                    contextMenuTab.Items.Remove(menuTextBoxTabAlias);
                    if(!Config.Tabs.RenameAmbTabs) {
                        contextMenuTab.Items.Insert(12, menuTextBoxTabAlias);
                        if(ContextMenuedTab.Comment.Length > 0) {
                            menuTextBoxTabAlias.Text = ContextMenuedTab.Comment;
                            menuTextBoxTabAlias.ForeColor = SystemColors.WindowText;
                        }
                        else {
                            menuTextBoxTabAlias.Text = QTUtility.ResMain[26];
                            menuTextBoxTabAlias.ForeColor = SystemColors.GrayText;
                        }
                        menuTextBoxTabAlias.Enabled = !tabControl1.AutoSubText;
                    }
                    if(tsmiTabOrder.DropDownItems.Count == 0) {
                        ((ToolStripDropDownMenu)tsmiTabOrder.DropDown).ShowImageMargin = false;
                        ToolStripMenuItem item2 = new ToolStripMenuItem(QTUtility.ResMain[28]);
                        ToolStripMenuItem item3 = new ToolStripMenuItem(QTUtility.ResMain[29]);
                        ToolStripMenuItem item4 = new ToolStripMenuItem(QTUtility.ResMain[30]);
                        ToolStripSeparator separator = new ToolStripSeparator();
                        ToolStripMenuItem item5 = new ToolStripMenuItem(QTUtility.ResMain[33]);
                        item2.Name = "Name";
                        item3.Name = "Drive";
                        item4.Name = "Active";
                        separator.Enabled = false;
                        item5.Name = "Rev";
                        tsmiTabOrder.DropDownItems.Add(item2);
                        tsmiTabOrder.DropDownItems.Add(item3);
                        tsmiTabOrder.DropDownItems.Add(item4);
                        tsmiTabOrder.DropDownItems.Add(separator);
                        tsmiTabOrder.DropDownItems.Add(item5);
                        tsmiTabOrder.DropDownItemClicked += menuitemTabOrder_DropDownItemClicked;
                    }
                    if((lstPluginMenuItems_Tab != null) && (lstPluginMenuItems_Tab.Count > 0)) {
                        foreach(ToolStripItem item6 in lstPluginMenuItems_Tab) {
                            item6.Dispose();
                        }
                        lstPluginMenuItems_Tab = null;
                    }
                    contextMenuTab.ResumeLayout();
                }
            }
            catch (Exception ex) { QTUtility2.MakeErrorLog(ex); }
        }

        private void Controls_GotFocus(object sender, EventArgs e) {
            OnGotFocus(e);
        }

        internal List<ToolStripItem> CreateBranchMenu(bool fCurrent, IContainer container, ToolStripItemClickedEventHandler itemClickedEvent) {
            QTabItem item = fCurrent ? CurrentTab : ContextMenuedTab;
            List<ToolStripItem> list = new List<ToolStripItem>();
            List<LogData> branches = item.Branches;
            if(branches.Count > 0) {
                ToolStripMenuItem item2 = new ToolStripMenuItem(QTUtility.ResMain[23]);
                item2.Tag = item;
                item2.DropDown = new DropDownMenuBase(container, true, true);
                item2.DropDown.ImageList = QTUtility.ImageListGlobal;
                item2.DropDownItemClicked += itemClickedEvent;
                int index = -1;
                foreach(LogData data in branches) {
                    index++;
                    if(IsSpecialFolderNeedsToTravel(data.Path)) {
                        if(LogEntryDic.ContainsKey(data.Hash)) {
                            goto Label_00B3;
                        }
                        continue;
                    }
                    if(!QTUtility2.PathExists(data.Path)) {
                        continue;
                    }
                Label_00B3:
                    item2.DropDownItems.Add(MenuUtility.CreateMenuItem(new MenuItemArguments(data.Path, false, index, MenuGenre.Branch)));
                }
                if(item2.DropDownItems.Count > 0) {
                    list.Add(new ToolStripSeparator());
                    list.Add(item2);
                }
            }
            return list;
        }

        private static Cursor CreateCursor(Bitmap bmpColor) {
            Cursor cursor;
            using(bmpColor) {
                using(Bitmap bitmap = new Bitmap(0x20, 0x20)) {
                    ICONINFO piconinfo = new ICONINFO();
                    piconinfo.fIcon = false;
                    piconinfo.hbmColor = bmpColor.GetHbitmap();
                    piconinfo.hbmMask = bitmap.GetHbitmap();
                    try {
                        cursor = new Cursor(PInvoke.CreateIconIndirect(ref piconinfo));
                    }
                    catch {
                        cursor = Cursors.Default;
                    }
                }
            }
            return cursor;
        }

        internal List<QMenuItem> CreateNavBtnMenuItems(bool fCurrent) {
            QTabItem item = fCurrent ? CurrentTab : ContextMenuedTab;
            List<QMenuItem> list = new List<QMenuItem>();
            string[] historyBack = item.GetHistoryBack();
            string[] historyForward = item.GetHistoryForward();
            if((historyBack.Length + historyForward.Length) > 1) {
                for(int i = historyBack.Length - 1; i >= 0; i--) {
                    QMenuItem item2 = MenuUtility.CreateMenuItem(new MenuItemArguments(historyBack[i], true, i, MenuGenre.Navigation));
                    if(IsSpecialFolderNeedsToTravel(historyBack[i])) {
                        item2.Enabled = LogEntryDic.ContainsKey(item.GetLogHash(true, i));
                    }
                    else if(!QTUtility2.PathExists(historyBack[i])) {
                        item2.Enabled = false;
                    }
                    if(item2.Enabled && (i == 0)) {
                        item2.BackColor = QTUtility2.MakeModColor(SystemColors.Highlight);
                    }
                    list.Add(item2);
                }
                for(int j = 0; j < historyForward.Length; j++) {
                    QMenuItem item3 = MenuUtility.CreateMenuItem(new MenuItemArguments(historyForward[j], false, j, MenuGenre.Navigation));
                    if(IsSpecialFolderNeedsToTravel(historyForward[j])) {
                        item3.Enabled = LogEntryDic.ContainsKey(item.GetLogHash(false, j));
                    }
                    else if(!QTUtility2.PathExists(historyForward[j])) {
                        item3.Enabled = false;
                    }
                    list.Add(item3);
                }
            }
            return list;
        }
        
        // �����µ�tabҳ
        private QTabItem CreateNewTab(IDLWrapper idlw) {
            string path = idlw.Path;
            QTabItem tab = new QTabItem(QTUtility2.MakePathDisplayText(path, false), path, tabControl1);
            tab.NavigatedTo(path, idlw.IDL, -1, false);
            tab.ToolTipText = QTUtility2.MakePathDisplayText(path, true);
            AddInsertTab(tab);
            return tab;
        }

        // ���� tab ͼƬ
        internal static Bitmap[] CreateTabImage() {
            if(File.Exists(Config.Skin.TabImageFile)) {
                try {
                    Bitmap[] bitmapArray = new Bitmap[3];
                    using(Bitmap bitmap = new Bitmap(Config.Skin.TabImageFile)) {
                        int height = bitmap.Height / 3;
                        bitmapArray[0] = bitmap.Clone(new Rectangle(0, 0, bitmap.Width, height), PixelFormat.Format32bppArgb);
                        bitmapArray[1] = bitmap.Clone(new Rectangle(0, height, bitmap.Width, height), PixelFormat.Format32bppArgb);
                        bitmapArray[2] = bitmap.Clone(new Rectangle(0, height * 2, bitmap.Width, height), PixelFormat.Format32bppArgb);
                    }
                    if(Path.GetExtension(Config.Skin.TabImageFile).PathEquals(".bmp")) {
                        bitmapArray[0].MakeTransparent(Color.Magenta);
                        bitmapArray[1].MakeTransparent(Color.Magenta);
                        bitmapArray[2].MakeTransparent(Color.Magenta);
                    }
                    return bitmapArray;
                }
                catch {
                }
            }
            return null;
        }

        // todo: handle links
        private static IEnumerable<string> CreateTMPPathsToOpenNew(Address[] addresses, string pathExclude) {
            List<string> list = new List<string>();
            QTUtility2.InitializeTemporaryPaths();
            for(int i = 0; i < addresses.Length; i++) {
                try {
                    using(IDLWrapper wrapper = new IDLWrapper(addresses[i].ITEMIDLIST)) {
                        if(wrapper.Available && wrapper.HasPath) {
                            string path = wrapper.Path;
                            if(path.Length > 0 && !path.PathEquals(pathExclude) && 
                                    !QTUtility2.IsShellPathButNotFileSystem(path) && 
                                    wrapper.IsFolder && !wrapper.IsLinkToDeadFolder) {
                                list.Add(path);
                            }
                        }
                    }
                }
                catch {
                }
            }
            return list;
        }

        private void ddmrUndoClose_ItemRightClicked(object sender, ItemRightClickedEventArgs e) {
            QMenuItem clickedItem = e.ClickedItem as QMenuItem;
            if(clickedItem != null) {
                using(IDLWrapper wrapper = new IDLWrapper(clickedItem.Path)) {
                    e.HRESULT = shellContextMenu.Open(wrapper, e.IsKey ? e.Point : MousePosition, ((DropDownMenuReorderable)sender).Handle, true);
                }
                if(e.HRESULT == 0xffff) {
                    StaticReg.ClosedTabHistoryList.Remove(clickedItem.Path);
                    e.ClickedItem.Dispose();
                }
            }
        }

        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool DoBindAction(BindAction action, bool fRepeat = false, QTabItem tab = null, IDLWrapper item = null) {
            if(fRepeat && !(
                    action == BindAction.GoBack ||
                    action == BindAction.GoForward ||
                    action == BindAction.TransparencyPlus ||
                    action == BindAction.TransparencyMinus)) {
                return false;
            }

            if(tab == null) tab = CurrentTab;

            // IntPtr ptr;
            switch(action) {
                case BindAction.GoBack:
                    NavigateCurrentTab(true);
                    break;

                case BindAction.GoForward:
                    NavigateCurrentTab(false);
                    break;

                case BindAction.GoFirst:
                    NavigateToFirstOrLast(true);
                    break;

                case BindAction.GoLast:
                    NavigateToFirstOrLast(false);
                    break;

                case BindAction.NextTab:
                    if(tabControl1.SelectedIndex == tabControl1.TabCount - 1) {
                        tabControl1.SelectedIndex = 0;
                    }
                    else {
                        tabControl1.SelectedIndex++;
                    }
                    break;

                case BindAction.PreviousTab:
                    if(tabControl1.SelectedIndex == 0) {
                        tabControl1.SelectedIndex = tabControl1.TabCount - 1;
                    }
                    else {
                        tabControl1.SelectedIndex--;
                    }
                    break;

                case BindAction.FirstTab:
                    tabControl1.SelectedIndex = 0;
                    break;

                case BindAction.LastTab:
                    tabControl1.SelectedIndex = tabControl1.TabCount - 1;
                    break;

                case BindAction.CloseCurrent:
                case BindAction.CloseTab:
                    NowTabDragging = false;
                    if(!tab.TabLocked) {
                        if(tabControl1.TabCount > 1) {
                            CloseTab(tab);
                        }
                        else {
                            WindowUtils.CloseExplorer(ExplorerHandle, 1);
                        }
                    }
                    break;

                case BindAction.CloseAllButCurrent: // �ر�����
                case BindAction.CloseAllButThis:
                    CloseAllTabsExcept(tab);
                    break;

                case BindAction.CloseLeft: // �����
                case BindAction.CloseLeftTab:
                    CloseLeftRight(true, tab.Index);
                    break;

                case BindAction.CloseRight: // �ر��ұ�
                case BindAction.CloseRightTab:
                    CloseLeftRight(false, tab.Index);
                    break;

                case BindAction.CloseWindow: // �رմ��� indiff
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root))
                    {
                        string[] list = (from QTabItem item2 in tabControl1.TabPages
                                         where item2.TabLocked
                                         select item2.CurrentPath).ToArray();

                        //MessageBox.Show(String.Join(",", list));
                        QTUtility2.WriteRegBinary(list, "TabsLocked", key);
                    }
                    WindowUtils.CloseExplorer(ExplorerHandle, 1);
                    break;

                case BindAction.RestoreLastClosed:
                    RestoreLastClosed();
                    break;

                case BindAction.CloneCurrent: // ���Ƶ�ǰ
                case BindAction.CloneTab:
                    CloneTabButton(tab, null, true, -1);
                    break;

                case BindAction.TearOffCurrent: //
                case BindAction.TearOffTab:
                    if(tabControl1.TabCount > 1) {
                        using(IDLWrapper wrapper = new IDLWrapper(tab.CurrentIDL)) {
                            OpenNewWindow(wrapper);
                        }
                        CloseTab(tab);
                    }
                    break;

                case BindAction.LockCurrent: // �رձ�ǩ
                case BindAction.LockTab:
                    tab.TabLocked = !tab.TabLocked;
                    break;

                case BindAction.LockAll: // ��������
                    bool lockState = tabControl1.TabPages.Any(t => t.TabLocked);
                    tabControl1.TabPages.ForEach(t => t.TabLocked = !lockState);
                    break;
                case BindAction.BrowseFolder: // ���ļ���
                    ChooseNewDirectory();
                    break;

                case BindAction.ShowOptions: // ��ʾѡ��
                    OptionsDialog.Open();
                    break;

                case BindAction.ShowToolbarMenu: // hmm.
                    Rectangle tabRect = tabControl1.GetTabRect(tabControl1.TabCount - 1, true);
                    contextMenuSys.Show(PointToScreen(new Point(tabRect.Right + 10, tabRect.Bottom - 10)));
                    break;

                case BindAction.ShowTabMenuCurrent:
                    if(tab.Index != -1) {
                        ContextMenuedTab = tab;
                        Rectangle rect = tabControl1.GetTabRect(tab.Index, true);
                        contextMenuTab.Show(PointToScreen(new Point(rect.Right + 10, rect.Bottom - 10)));
                    }
                    break;

                case BindAction.ShowTabMenu:
                    ContextMenuedTab = tab;
                    contextMenuTab.Show(MousePosition);
                    break;

                case BindAction.ShowRecentTabsMenu:
                    break;

                case BindAction.CopySelectedPaths:
                    if(listView.SubDirTipMenuIsShowing() || (subDirTip_Tab != null && subDirTip_Tab.MenuIsShowing)) {
                        return false;
                    }
                    DoFileTools(0);
                    break;

                case BindAction.CopySelectedNames:
                    if(listView.SubDirTipMenuIsShowing() || (subDirTip_Tab != null && subDirTip_Tab.MenuIsShowing)) {
                        return false;
                    }
                    DoFileTools(1);
                    break;

                case BindAction.CopyCurrentFolderPath:
                    DoFileTools(2);
                    break;

                case BindAction.CopyCurrentFolderName:
                    DoFileTools(3);
                    break;

                case BindAction.ChecksumSelected:
                    DoFileTools(4);
                    break;

                case BindAction.ToggleTopMost:
                    ToggleTopMost(); // todo: move v to < ?
                    break;

                case BindAction.TransparencyPlus:
                case BindAction.TransparencyMinus: {
                        // TODO!!!
                        int num9;
                        int num10;
                        byte num11;
                        if(0x80000 != ((int)PInvoke.Ptr_OP_AND(PInvoke.GetWindowLongPtr(ExplorerHandle, -20), 0x80000))) {
                            if(action == BindAction.TransparencyPlus) {
                                return true;
                            }
                            PInvoke.SetWindowLongPtr(ExplorerHandle, -20, PInvoke.Ptr_OP_OR(PInvoke.GetWindowLongPtr(ExplorerHandle, -20), 0x80000));
                            PInvoke.SetLayeredWindowAttributes(ExplorerHandle, 0, 0xff, 2);
                        }
                        if(PInvoke.GetLayeredWindowAttributes(ExplorerHandle, out num9, out num11, out num10)) {
                           // IntPtr ptr3;
                            if(action == BindAction.TransparencyPlus) {
                                if(num11 > 0xf3) {
                                    num11 = 0xff;
                                }
                                else {
                                    num11 = (byte)(num11 + 12);
                                }
                            }
                            else if(num11 < 0x20) {
                                num11 = 20;
                            }
                            else {
                                num11 = (byte)(num11 - 12);
                            }
                            PInvoke.SetLayeredWindowAttributes(ExplorerHandle, 0, num11, 2);
                            if(num11 == 0xff) {
                                PInvoke.SetWindowLongPtr(ExplorerHandle, -20, PInvoke.Ptr_OP_AND(PInvoke.GetWindowLongPtr(ExplorerHandle, -20), 0xfff7ffff));
                            }
                        }
                    }
                    break;

                case BindAction.FocusFileList:
                    listView.SetFocus();
                    break;

                case BindAction.FocusSearchBarReal:
                    if(QTUtility.IsXP) return false;
                    // todo, I don't think this works
                    PInvoke.SetFocus(GetSearchBand_Edit());
                    break;

                case BindAction.ShowSDTSelected:
                    if(!Config.Tips.ShowSubDirTips) return false;
                    listView.ShowAndClickSubDirTip();
                    break;

                case BindAction.SendToTray:
                    MinimizeToTray();
                    break;

                case BindAction.FocusTabBar:
                    tabControl1.Focus();
                    tabControl1.FocusNextTab(false, true, false);
                    break;

                case BindAction.NewTab:
                    using(IDLWrapper wrapper = new IDLWrapper(Config.Window.DefaultLocation)) {
                        OpenNewTab(wrapper, false, true);    
                    }
                    break;

                case BindAction.NewWindow:
                    using(IDLWrapper wrapper = new IDLWrapper(Config.Window.DefaultLocation)) {
                        OpenNewWindow(wrapper);
                    }
                    break;

                // TODO all the blank ones
                case BindAction.NewFolder:
                    break;
                case BindAction.NewFile:
                    createNewFile();
                    break;

                case BindAction.SwitchToLastActivated:
                    if(lstActivatedTabs.Count > 1 && tabControl1.TabPages.Contains(lstActivatedTabs[lstActivatedTabs.Count - 2])) {
                        try {
                            tabControl1.SelectTab(lstActivatedTabs[lstActivatedTabs.Count - 2]);
                        }
                        catch(ArgumentException) {
                        }
                    }
                    break;

                case BindAction.MergeWindows:
                    MergeAllWindows();
                    break;

                case BindAction.ShowRecentFilesMenu:
                    break;
                case BindAction.SortTabsByName:
                    break;
                case BindAction.SortTabsByPath:
                    break;
                case BindAction.SortTabsByActive:
                    break;

                case BindAction.UpOneLevelTab:
                    // QTUtility2.log("QTTabBarClass UpOneLevelTab");
                case BindAction.UpOneLevel:
                    QTUtility2.log("QTTabBarClass case UpOneLevel");
                    UpOneLevel(); // Hmm...
                    break;

                case BindAction.Refresh:
                    break;
                case BindAction.Paste:
                    break;
                case BindAction.Maximize:
                    break;
                case BindAction.Minimize:
                    break;

                case BindAction.CopyTabPath:
                    // ���Ʊ�ǩ��·��·��
                    string currentPath = tab.CurrentPath;
                    if(currentPath.IndexOf("???") != -1) {
                        currentPath = currentPath.Substring(0, currentPath.IndexOf("???"));
                    }
                    QTUtility2.SetStringClipboard(currentPath);
                    break;

                case BindAction.TabProperties:
                    ShellMethods.ShowProperties(tab.CurrentIDL);
                    break;

                case BindAction.ShowTabSubfolderMenu:
                    break;

                case BindAction.ItemOpenInNewTab:
                case BindAction.ItemOpenInNewTabNoSel:
                case BindAction.ItemOpenInNewWindow:
                    if(item.Available && item.HasPath && item.IsReadyIfDrive && !item.IsLinkToDeadFolder) {
                        using(IDLWrapper linkWrapper = item.ResolveTargetIfLink()) {
                            IDLWrapper actualItem = linkWrapper ?? item;
                            if(actualItem.IsFolder) {
                                if(action == BindAction.ItemOpenInNewWindow) {
                                    OpenNewWindow(actualItem);
                                }
                                else {
                                    OpenNewTab(actualItem, action == BindAction.ItemOpenInNewTabNoSel);
                                }
                            }
                        }
                    }
                    break;

                case BindAction.ItemCut:
                case BindAction.ItemCopy:      
                case BindAction.ItemDelete:
                    break;

                case BindAction.ItemProperties:
                    ShellMethods.ShowProperties(item.IDL);
                    break;

                case BindAction.CopyItemPath:
                case BindAction.CopyItemName:
                case BindAction.ChecksumItem:
                    break;
                /***** add by qwop start ***/
                case BindAction.OpenCmd:  // ��������ʾ��
                    OpenCmd( tab ); // add by qwop...
                    break;
                case BindAction.ItemsOpenInNewTabNoSel: // ��ѡ�е��ļ��� ����ǩҳ(������)
                    Address[] addressArray;
                    if ( ShellBrowser.TryGetSelection(out addressArray, false  )) {
                        foreach (Address address in addressArray)
                        {
                            if (address.Path != null && Directory.Exists(address.Path) ) {
                                OpenNewTab(address.Path, action == BindAction.ItemsOpenInNewTabNoSel);
                            }
                        }
                    }                   
                    break;
                /***** add by qwop end   ***/
            }
            return true;
        }

        /// <summary>
        /// �������ļ� add by indiff
        /// </summary>
        private void createNewFile()
        {

            // Create new file
            IShellView shellView = null;
            IntPtr pIDL = IntPtr.Zero;

            try
            {
                string path = "";

                if (String.IsNullOrEmpty(path) || !Directory.Exists(path))
                {
                    QTUtility.SoundPlay();
                    return;
                }

                // make new name

                int i = 2;
                // string name = "�½��ı��ĵ�";
                string name = QTUtility.DefaultNewFileName();
                string ext =  ".txt";
                string pathNew = path + "\\" + name + ext;

                while ( Directory.Exists(pathNew) || File.Exists(pathNew) )
                {
                    pathNew = path + "\\" + name + " (" + i + ")" + ext;
                    i++;
                }

                using (File.Create(pathNew))
                {
                }

                // Select and put into rename mode.
                if (0 == ShellBrowser.GetIShellBrowser().QueryActiveShellView(out shellView))
                {
                    shellView.Refresh();

                    pIDL = PInvoke.ILCreateFromPath(pathNew);
                    if (pIDL != IntPtr.Zero)
                    {
                        IntPtr pIDLRltv = PInvoke.ILFindLastID(pIDL);
                        if (pIDLRltv != IntPtr.Zero) {
                            // ѡ���ļ�
                             shellView.SelectItem(pIDLRltv, SVSIF.SELECT | SVSIF.DESELECTOTHERS | SVSIF.ENSUREVISIBLE | SVSIF.EDIT );
                           //  ShellBrowser.GetIShellBrowser().SelectItem(pIDLRltv, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_EDIT);
                            return;
                        }
                    }
                }
            }
            catch
            {
            }
            finally
            {
                if (ShellBrowser.GetIShellBrowser() != null)
                {
                    QTUtility2.log("ReleaseComObject ShellBrowser.GetIShellBrowser()");
                    Marshal.ReleaseComObject(ShellBrowser.GetIShellBrowser());
                }
                    

                if (pIDL != IntPtr.Zero)
                    Marshal.FreeCoTaskMem(pIDL);
            }

            QTUtility.SoundPlay();
       
        }
        /***** add by qwop end   ***/

        /// <summary>
        /// �ú��������ɲ�ͬ�̲߳����Ĵ��ڵ���ʾ״̬��
        /// </summary>
        /// <param name="hWnd">���ھ��</param>
        /// <param name="cmdShow">ָ�����������ʾ���鿴����ֵ�б������ShowWlndow������˵�����֡�</param>
        /// <returns>�������ԭ���ɼ�������ֵΪ���㣻�������ԭ�������أ�����ֵΪ�㡣</returns>
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);
        /// <summary>
        /// �ú���������ָ�����ڵ��߳����õ�ǰ̨�����Ҽ���ô��ڡ���������ת��ô��ڣ���Ϊ�û��ĸ��ֿ��ӵļǺš�ϵͳ������ǰ̨���ڵ��̷߳����Ȩ���Ը��������̡߳�
        /// </summary>
        /// <param name="hWnd">�������������ǰ̨�Ĵ��ھ����</param>
        /// <returns>�������������ǰ̨������ֵΪ���㣻�������δ������ǰ̨������ֵΪ�㡣</returns>
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32.dll")]
        private static extern bool SetFocus(IntPtr hWnd);

        private const int WS_SHOWNORMAL = 1;

        private void cmdPath(string currentPath)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = "/k cd " + currentPath;
            process.StartInfo.WorkingDirectory = currentPath;
            process.Start();

            ShowWindowAsync(process.MainWindowHandle, WS_SHOWNORMAL); //��ʾ������ע�͵�
            SetForegroundWindow(process.MainWindowHandle);            //�ŵ�ǰ��
            SetFocus(process.MainWindowHandle);
        }

        /************************************************************************/
        /* �򿪵�ǰĿ¼�� ������ʾ��                                            */
        /************************************************************************/
        private void OpenCmd( QTabItem tab )
        {
            /*            if ( tab == null ) {
                            return ;
                        }*/
            if (ShellBrowser.GetIShellBrowser() != null) { 
                string currentPath = "";
                if (tab != null)
                {
                    currentPath = tab.CurrentPath;
                }
                else
                {
                    currentPath = ContextMenuedTab.CurrentPath;
                }
            
                if (currentPath.IndexOf("???") != -1)
                {
                    currentPath = currentPath.Substring(0, currentPath.IndexOf("???"));
                }
                else if (currentPath.IndexOf("*?*?*") != -1)
                {
                    currentPath = currentPath.Substring(0, currentPath.IndexOf("*?*?*"));
                }
                // �ж��ļ��Ƿ����
                if (Directory.Exists(currentPath))
                {
                    /*
                    ProcessStartInfo startInfo = new ProcessStartInfo("cmd");
                    startInfo.WindowStyle = ProcessWindowStyle.Normal;
                    startInfo.Verb = "runas";
                    startInfo.CreateNoWindow = false;
                
                
                    startInfo.WorkingDirectory = currentPath;
                    // ��ӻ�ȡ����
                    Process instance = Process.Start(startInfo );
                    //  instance.WaitForInputIdle();
                    
                    if (!instance.WaitForInputIdle(10000)) // 10 s timout 
                    {
                        throw new ApplicationException("Process takes too much time to start");
                    }
                    */

                    cmdPath(currentPath);

                } // end for open cmd.
                else { 
                    // �Ҳ���·�����ϵͳ��
                    if (QTUtility2.PathExists("C:\\"))
                    {
                        cmdPath("C:\\");
                    } else if (QTUtility2.PathExists("D:\\"))
                    {
                        cmdPath("D:\\");
                    }
                    else if (QTUtility2.PathExists("E:\\"))
                    {
                        cmdPath("E:\\");
                    }
                    else if (QTUtility2.PathExists("F:\\"))
                    {
                        cmdPath("F:\\");
                    }
                }
            }
        }

        // todo: clean, enum.
        internal bool DoFileTools(int index) {
            try {
                Address[] addressArray;
                List<string> list;
                int num;
                string displayName = string.Empty;
                switch(index) {
                    case 0:
                    case 1:
                    case 4:
                        string str2;
                        if(!ShellBrowser.TryGetSelection(out addressArray, out str2, index == 1)) {
                            goto Label_019C;
                        }
                        list = new List<string>();
                        num = 0;
                        goto Label_00A1;

                    case 2:
                    case 3: {
                            using(IDLWrapper wrapper = ShellBrowser.GetShellPath()) {
                                if(wrapper.Available) {
                                    displayName = ShellMethods.GetDisplayName(wrapper.PIDL, index == 3);
                                }
                                goto Label_019C;
                            }
                        }
                    case 5:
                        foreach(QTabItem item in tabControl1.TabPages) {
                            string currentPath = item.CurrentPath;
                            int length = currentPath.IndexOf("???");
                            if(length != -1) {
                                currentPath = currentPath.Substring(0, length);
                            }
                            int num3 = currentPath.IndexOf("*?*?*");
                            if(num3 != -1) {
                                currentPath = currentPath.Substring(0, num3);
                            }
                            displayName = displayName + ((displayName.Length == 0) ? string.Empty : "\r\n") + currentPath;
                        }
                        goto Label_019C;

                    default:
                        goto Label_019C;
                }
            Label_004B:
                if(addressArray[num].Path != null) {
                    if(index != 4) {
                        displayName = displayName + ((displayName.Length == 0) ? string.Empty : "\r\n") + addressArray[num].Path;
                    }
                    else {
                        list.Add(addressArray[num].Path);
                    }
                }
                num++;
            Label_00A1:
                if(num < addressArray.Length) {
                    goto Label_004B;
                }
                if(index == 4) {
                    ShowMD5(list.ToArray());
                    return true;
                }
            Label_019C:
                if(displayName.Length > 0) {
                    QTUtility2.SetStringClipboard(displayName);
                    return true;
                }
            }
            catch (Exception e){
                QTUtility2.MakeErrorLog( e, "DoFileTools");
            }
            if(index == 4) {
                // .net4 ��������ǿ����������е�
                /*QTUtility2.log("md5 out tabbar.OpenNewTab");
                InstanceManager.BeginInvokeMain(tabbar =>
                {
                    QTUtility2.log("md5 inner tabbar.OpenNewTab");
                    tabbar.OpenNewTab("c:/");
                    tabbar.RestoreWindow();
                });*/

                ShowMD5(null);
                return true;
            }
            return false;
        }

        // ����ķ�ʽ  select 1 / factory 2 / other 3
        private int mCmdType = 0;

        // This function is either called by BeforeNavigate2 (on XP and Vista)
        // �˺�����BeforeNavigate2����(on XP and Vista)
        // or NavigateComplete2 (on 7)
        private void DoFirstNavigation(bool before, string path) {
            // TODO: sort out this mess  ����һ���������
            if(StaticReg.CreateWindowPaths.Count > 0 || StaticReg.CreateWindowIDLs.Count > 0) {
                QTUtility2.log("DoFirstNavigation StaticReg.CreateWindowPaths.Count " + StaticReg.CreateWindowPaths.Count + " StaticReg.CreateWindowIDLs.Count:" + StaticReg.CreateWindowIDLs.Count);
                foreach (string tpath in StaticReg.CreateWindowPaths.Where(str2 => !str2.PathEquals(path))) {
                    using(IDLWrapper wrapper = new IDLWrapper(tpath)) {
                        if(wrapper.Available) {
                            CreateNewTab(wrapper);
                        }
                    }
                }
                foreach(byte[] idl in StaticReg.CreateWindowIDLs) {
                    using(IDLWrapper wrapper2 = new IDLWrapper(idl)) {
                        OpenNewTab(wrapper2, true);
                    }
                }
                QTUtility2.InitializeTemporaryPaths();
                AddStartUpTabs(string.Empty, path);
                InitializeOpenedWindow();
            }
            else
            {
                AddStartUpTabs(string.Empty, path);
                InitializeOpenedWindow();
            }
        }

        private int dropTargetWrapper_DragFileDrop(out IntPtr hwnd, out byte[] idlReal) {
            HideToolTipForDD();
            hwnd = tabControl1.Handle;
            idlReal = null;
            QTabItem tabMouseOn = tabControl1.GetTabMouseOn();
            if((tabMouseOn == null) || !Config.Tabs.DragOverTabOpensSDT) {
                return 1;
            }
            if((tabMouseOn.CurrentIDL != null) && (tabMouseOn.CurrentIDL.Length > 0)) {
                idlReal = tabMouseOn.CurrentIDL;
                return 0;
            }
            return -1;
        }

        private DragDropEffects dropTargetWrapper_DragFileEnter(IntPtr hDrop, Point pnt, int grfKeyState) {
            if(Config.Tabs.DragOverTabOpensSDT) {
                int num = HandleDragEnter(hDrop, out strDraggingDrive, out strDraggingStartPath);
                fDrivesContainedDD = num == 2;
                if(num == -1) {
                    return DragDropEffects.None;
                }
                if(tabControl1.GetTabMouseOn() == null) {
                    return DragDropEffects.Copy;
                }
                switch(num) {
                    case 0:
                        return DropTargetWrapper.MakeEffect(grfKeyState, 0);

                    case 1:
                        return DropTargetWrapper.MakeEffect(grfKeyState, 1);

                    case 2:
                        return DragDropEffects.None;
                }
            }
            return DragDropEffects.Copy;
        }

        private void dropTargetWrapper_DragFileLeave(object sender, EventArgs e) {
            HideToolTipForDD();
            strDraggingDrive = null;
            strDraggingStartPath = null;
            tabControl1.Refresh();
        }

        private void dropTargetWrapper_DragFileOver(object sender, DragEventArgs e) {
            QTUtility2.log("QTTabBarClass dropTargetWrapper_DragFileOver");
            e.Effect = DragDropEffects.None;
            QTabItem mouseOnTab = tabControl1.GetTabMouseOn(); // ��������ı�ǩ 
            bool flag = true;
            if(mouseOnTab != tabForDD) {
                tabControl1.Refresh();
                HideSubDirTip_Tab_Menu();
                fToggleTabMenu = false;
                flag = false;
            }
            if(mouseOnTab == null) {
                e.Effect = DragDropEffects.Copy;
            }
            else if(mouseOnTab.CurrentPath.Length > 2) {
                if(fDrivesContainedDD || strDraggingStartPath.PathEquals(mouseOnTab.CurrentPath)) {
                    if(toolTipForDD != null) {
                        toolTipForDD.Hide(tabControl1);
                    }
                    ShowToolTipForDD(mouseOnTab, -1, e.KeyState); // ��ʾtip ��ʾ��Ϣ
                }
                else {
                    using(IDLWrapper wrapper = new IDLWrapper(mouseOnTab.CurrentIDL, !flag)) {
                        if(wrapper.Available && wrapper.IsDropTarget) {
                            string b = mouseOnTab.CurrentPath.Substring(0, 3);
                            int num = strDraggingDrive != null && strDraggingDrive.Equals(b, StringComparison.OrdinalIgnoreCase)
                                    ? 0 : 1;
                            ShowToolTipForDD(mouseOnTab, num, e.KeyState);
                            e.Effect = Config.Tabs.DragOverTabOpensSDT
                                    ? DropTargetWrapper.MakeEffect(e.KeyState, num)
                                    : DragDropEffects.Copy;
                        }
                        else {
                            HideToolTipForDD();
                        }
                    }
                }
            }
            else {
                HideToolTipForDD();
            }
        }

        private void Explorer_BeforeNavigate2(object pDisp, 
                                                ref object URL, 
                                                ref object Flags, 
                                                ref object TargetFrameName, 
                                                ref object PostData, 
                                                ref object Headers, 
                                                ref bool Cancel) {
            // DebugUtil.WriteLine("QTTabBarClass Explorer_BeforeNavigate2:" ); // add by qwop.
            QTUtility2.log("QTTabBarClass Explorer_BeforeNavigate2  pDisp :" + pDisp
                    + " URL :" + (string)URL
                    + " Flags :" + Flags
                    + " TargetFrameName :" + TargetFrameName
                    + " PostData :" + PostData
                    + " Headers :" + Headers
                    + " Cancel :" + Cancel
                );
            if(!IsShown) {
                DoFirstNavigation(true, (string)URL);
            }
        }



        private void Explorer_NavigateComplete2(object pDisp, ref object URL) {
            // add by indiff.
            /*QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2  pDisp :" 
                           + pDisp
                           + " URL :" + (string)URL
            );*/
            string path = (string)URL;
            lastCompletedBrowseObjectIDL = lastAttemptedBrowseObjectIDL;
            // ��������ʱ��ˢ���ļ�����ͼ
            QTUtility2.log("QTTabBarClass ShellBrowser.OnNavigateComplete reset field FolderView");
            ShellBrowser.OnNavigateComplete();
            
            // if(fFinalRelease && !IsShown) {
            if(!IsShown) {
                QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2  !IsShown");
                DoFirstNavigation(false, path);
            }            

            if(fNowQuitting)
            {
                fHideExplorer = true;

                if (mCmdType == 3)
                {
                    Explorer.Quit();
                    WindowUtils.HideExplorer(ExplorerHandle);
                }
            }
            else {
                int hash = -1;
                // �Ƿ�Ϊ����Ŀ¼
                bool flag = IsSpecialFolderNeedsToTravel(path);
                // �Ƿ�Ϊ shellPath 
                bool flag2 = QTUtility2.IsShellPathButNotFileSystem(path);
                // �Ƿ�Ϊ shellPath 
                bool flag3 = QTUtility2.IsShellPathButNotFileSystem(CurrentTab.CurrentPath);

                // If we're navigating on a locked tab, we simulate opening the target folder
                // in a new tab.  First we clone the tab at the old address and lock it.  Then
                // we move the current tab to the "new tab" position and unlock it.
                if(!flag2 && !flag3 && !NavigatedByCode && CurrentTab.TabLocked) { // �����ǩ����״̬
                    QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2  !flag2 && !flag3 && !NavigatedByCode && CurrentTab.TabLocked");
                    int pos = tabControl1.SelectedIndex;
                    tabControl1.SetRedraw(false);
                    QTabItem item = CloneTabButton(CurrentTab, null, false, pos);
                    item.TabLocked = true;
                    CurrentTab.TabLocked = false;
                    pos++;
                    int max = tabControl1.TabPages.Count - 1;

                    switch(Config.Tabs.NewTabPosition) {
                        case TabPos.Rightmost:
                            if(pos != max) {
                                tabControl1.TabPages.Relocate(pos, max);
                            }
                            break;
                        case TabPos.Leftmost:
                            tabControl1.TabPages.Relocate(pos, 0);
                            break;
                        case TabPos.Left:
                            tabControl1.TabPages.Relocate(pos, pos - 1);
                            break;
                    }
                    tabControl1.SetRedraw(true);

                    lstActivatedTabs.Remove(CurrentTab);
                    lstActivatedTabs.Add(item);
                    lstActivatedTabs.Add(CurrentTab);
                    if(lstActivatedTabs.Count > 15) {
                        lstActivatedTabs.RemoveAt(0);
                    }
                }
                // ���������Ŀ¼,��code��ʽ����
                if(!NavigatedByCode && flag) {
                    QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2  !NavigatedByCode && flag");
                    hash = DateTime.Now.GetHashCode();
                    LogEntryDic[hash] = GetCurrentLogEntry();
                }
                ClearTravelLogs();
                try {
                    tabControl1.SetRedraw(false);
                    if(fNowTravelByTree) { // ���α���, �м�������������࿪һ����ǩ��
                        // ��ȡ��ǰ�򿪵�·��
                        using(IDLWrapper wrapper = GetCurrentPIDL()) {
                            // ����һ���±�ǩ����ѡ�У����õ�ǰ�ı�ǩ
                            QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2  fNowTravelByTree CreateNewTab");
                            QTabItem tabPage = CreateNewTab(wrapper);
                            tabControl1.SelectTabDirectly(tabPage);
                            CurrentTab = tabPage;
                        }
                    }
                    if(tabControl1.AutoSubText && !fNavigatedByTabSelection) {
                        CurrentTab.Comment = string.Empty;
                    }
                    CurrentAddress = path;
                    CurrentTab.Text = Explorer.LocationName;
                    CurrentTab.CurrentIDL = null;
                    CurrentTab.ShellToolTip = null;
                    byte[] idl;
                    using(IDLWrapper wrapper2 = GetCurrentPIDL()) {
                        CurrentTab.CurrentIDL = idl = wrapper2.IDL;
                        if(flag) {
                            if((!NavigatedByCode && (idl != null)) && (idl.Length > 0)) {
                                path = path + "*?*?*" + hash;
                                QTUtility.ITEMIDLIST_Dic_Session[path] = idl;
                                CurrentTab.CurrentPath = CurrentAddress = path;
                            }
                        }
                        else if((flag2 && wrapper2.Available) && !CurrentTab.CurrentPath.Contains("???")) {
                            string str2;
                            int num2;
                            if(IDLWrapper.GetIDLHash(wrapper2.PIDL, out num2, out str2)) {
                                hash = num2;
                                CurrentTab.CurrentPath = CurrentAddress = path = str2;
                            }
                            else if((idl != null) && (idl.Length > 0)) {
                                hash = num2;
                                path = path + "???" + hash;
                                IDLWrapper.AddCache(path, idl);
                                CurrentTab.CurrentPath = CurrentAddress = path;
                            }
                        }
                        if(!NavigatedByCode) {
                            QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2  !NavigatedByCode");
                            CurrentTab.NavigatedTo(CurrentAddress, idl, hash, fAutoNavigating);
                        }
                    }
                    // ͬ��������״̬
                    SyncTravelState();
                    if (QTUtility.IsXP) // XPϵͳ
                    {
                        if (CurrentAddress.StartsWith(QTUtility.PATH_SEARCHFOLDER))
                        {
                            QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2 ShowSearchBar(true)");
                            ShowSearchBar(true);
                        }
                        else if (QTUtility.fExplorerPrevented)
                        {
                            QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2 ShowFolderTree(true)");
                            ShowFolderTree(true);
                            QTUtility.fExplorerPrevented = false;
                        }
                    }                    
                    if(CurrentAddress.StartsWith("::")) { // ������ʾ���ƣ�����Ŀ¼
                        CurrentTab.ToolTipText = CurrentTab.Text;
                        QTUtility.DisplayNameCacheDic[CurrentAddress] = CurrentTab.Text;
                    }
                    else if(flag2) {
                        CurrentTab.ToolTipText = (string)URL;
                    }
                    else if(((CurrentAddress.Length == 3) 
                             || CurrentAddress.StartsWith(@"\\")) 
                             || (CurrentAddress.StartsWith("http://", StringComparison.OrdinalIgnoreCase) 
                             || CurrentAddress.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase))) {
                        CurrentTab.ToolTipText = CurrentTab.CurrentPath;
                        QTUtility.DisplayNameCacheDic[CurrentAddress] = CurrentTab.Text;
                    }
                    else {
                        CurrentTab.ToolTipText = CurrentTab.CurrentPath;
                    }
                    if(NavigatedByCode && !NowTabCreated) {
                        string str3;
                        Address[] selectedItemsAt = CurrentTab.GetSelectedItemsAt(CurrentAddress, out str3);
                        if(selectedItemsAt != null) {
                            QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2 ShellBrowser.TrySetSelection " + str3);
                            ShellBrowser.TrySetSelection(selectedItemsAt, str3, true);
                        }
                    }
                    if(QTUtility.RestoreFolderTree_Hide) {
                        QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2 QTUtility.RestoreFolderTree_Hide");
                        new WaitTimeoutCallback(WaitTimeout).BeginInvoke(150, AsyncComplete_FolderTree, false);
                    }
                    if(fNowRestoring) {
                        fNowRestoring = false;
                        if(StaticReg.LockedTabsToRestoreList.Contains(path)) {
                            CurrentTab.TabLocked = true;
                        }
                    }
                    if( (!QTUtility.IsXP 
                         || FirstNavigationCompleted) && 
                        (!PInvoke.IsWindowVisible(ExplorerHandle) 
                         || PInvoke.IsIconic(ExplorerHandle))) {
                        QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2 WindowUtils.BringExplorerToFront");
                        WindowUtils.BringExplorerToFront(ExplorerHandle);
                    }
                    if(buttonNavHistoryMenu.DropDown.Visible)
                    {
                        QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2 buttonNavHistoryMenu.DropDown.Visible");
                        buttonNavHistoryMenu.DropDown.Close(ToolStripDropDownCloseReason.AppFocusChange);
                    }
                }
                catch(Exception exception) {
                    QTUtility2.MakeErrorLog(exception);
                }
                finally {
                    QTUtility.RestoreFolderTree_Hide = 
                            NavigatedByCode = 
                            fNavigatedByTabSelection = 
                            NowTabCreated = 
                            fNowTravelByTree = false;
                    QTUtility2.log("QTTabBarClass Explorer_NavigateComplete2 tabControl1.SetRedraw(true)");
                    tabControl1.SetRedraw(true);
                    FirstNavigationCompleted = true;
                    // this.listView
                    // compatibleView.SetBackColor(ColorTranslator.ToWin32(SystemColors.Window),   ColorTranslator.ToWin32(System.Drawing.Color.FromArgb(244, 248, 253)));
                    // listView.RefreshViewWatermark(false);
                }
            }
        }


        // ��Ϣ����
        private bool explorerController_MessageCaptured(ref Message msg) {
            if (msg.Msg != WM.CLOSE) {
                iSequential_WM_CLOSE = 0;
            }

            if(msg.Msg == WM_BROWSEOBJECT) {
                SBSP flags = (SBSP)Marshal.ReadInt32(msg.WParam);
                if((flags & SBSP.NAVIGATEBACK) != 0) {
                    msg.Result = (IntPtr)1;
                    QTUtility2.log("explorerController_MessageCaptured WM_BROWSEOBJECT: NAVIGATEBACK");
                    if(!NavigateCurrentTab(true) && CloseTab(CurrentTab, true) && tabControl1.TabCount == 0) {
                        WindowUtils.CloseExplorer(ExplorerHandle, 2);
                    }
                }
                else if((flags & SBSP.NAVIGATEFORWARD) != 0) {
                    QTUtility2.log("explorerController_MessageCaptured WM_BROWSEOBJECT: NAVIGATEFORWARD");
                    msg.Result = (IntPtr)1;
                    NavigateCurrentTab(false);
                }
                else {
                    QTUtility2.log("explorerController_MessageCaptured PInvoke.ILClone: ");
                    var commandLine = GetCommandLine();
                    IntPtr pidl = IntPtr.Zero;
                    if(msg.LParam != IntPtr.Zero) {
                        pidl = PInvoke.ILClone(msg.LParam);
                    }
                    bool autonav = (flags & SBSP.AUTONAVIGATE) != 0;
                    using(IDLWrapper wrapper = new IDLWrapper(pidl)) {
                        msg.Result = (IntPtr)(BeforeNavigate(wrapper, autonav) ? 1 : 0);
                    }
                }
                return true;
            }
            else if(msg.Msg == WM_HEADERINALLVIEWS) {
                msg.Result = (IntPtr)(Config.Tweaks.AlwaysShowHeaders ? 1 : 0);
                return true;
            }
            else if(msg.Msg == WM_SHOWHIDEBARS) {
                // Todo: hardcoding = bad
                object pvaTabBar = new Guid("{d2bf470e-ed1c-487f-a333-2bd8835eb6ce}").ToString("B");
                object pvarShow = (msg.WParam != IntPtr.Zero);
                object pvarSize = null;
                try {
                    Explorer.ShowBrowserBar(pvaTabBar, pvarShow, pvarSize);
                    msg.Result = (IntPtr)1;

                    QTUtility2.flog("QTTabBarClass WM_SHOWHIDEBARS ShowBrowsserBar tabBar buttonBar");
                }
                catch(COMException e) {
                    QTUtility2.MakeErrorLog(e, "WM_SHOWHIDEBARS ShowBrowserBar");
                }
                return true;
            }
            else if(msg.Msg == WM_CHECKPULSE) {
                if(fNeedsNewWindowPulse && msg.LParam != IntPtr.Zero) {
                    Marshal.WriteIntPtr(msg.LParam, Marshal.GetIDispatchForObject(Explorer));
                    msg.Result = (IntPtr)1;
                    fNeedsNewWindowPulse = false;
                }
                return true;
            }
            else if (msg.Msg == WM_SELECTFILE)
            {
                // MessageBox.Show("select file 2");
                QTUtility2.log(" select file 2  wparam " + msg.WParam + " lparam " + msg.LParam);
                return true;
            }

            switch(msg.Msg) {
                case WM.SETTINGCHANGE:
                    if(QTUtility.IsXP) {
                        QTUtility.GetShellClickMode();
                    }
                    if(Marshal.PtrToStringUni(msg.LParam) == "Environment") {
                        // todo: hmmmm....
                        //QTUtility.fRequiredRefresh_App = true;
                        SyncTaskBarMenu();
                    }
                    return false;

                case WM.NCLBUTTONDOWN:
                case WM.NCRBUTTONDOWN:
                    HideTabSwitcher(false);
                    return false;

                case WM.MOVE:
                case WM.SIZE:
                    listView.HideThumbnailTooltip(0);
                    listView.HideSubDirTip(0);
                    return false;

                case WM.ACTIVATE: {
                    int num3 = ((int) msg.WParam) & 0xffff;
                    if(num3 > 0) {
                        // I can't figure out why, but calling these methods synchronously
                        // prevents the window from being restored when minimized.
                        // I guess neither has to synchronous anyway, so...
                        BeginInvoke(new Action(() => {
                            InstanceManager.PushTabBarInstance(this);
                            InstanceManager.RemoveFromTrayIcon(Handle);
                        }));
                    }
                    else {
                        listView.HideThumbnailTooltip(1);
                        listView.HideSubDirTip_ExplorerInactivated();
                        HideTabSwitcher(false);
                        if(tabControl1.Focused) {
                            listView.SetFocus();
                        }
                        if((Config.Tabs.ShowCloseButtons &&
                                Config.Tabs.CloseBtnsWithAlt) &&
                                        tabControl1.EnableCloseButton) {
                            tabControl1.EnableCloseButton = false;
                            tabControl1.Refresh();
                        }
                    }
                    return false;
                }
                case WM.CLOSE:
                    if(iSequential_WM_CLOSE > 0) {
                        return true;
                    }
                    iSequential_WM_CLOSE++;
                    return HandleCLOSE(msg.LParam);

                case WM.NCMBUTTONDOWN:
                case WM.NCXBUTTONDOWN:
                    HideTabSwitcher(false);
                    return false;

                case WM.SYSCOMMAND:
                    if((((int) msg.WParam) & 0xfff0) == 0xf020) {
                        if(Config.Window.TrayOnMinimize) {
                            MinimizeToTray();
                            return true;
                        }
                        return false;
                    }
                    if((((int) msg.WParam) & 0xfff0) == 0xf030) {
                        return false;
                    }
                    if((((int) msg.WParam) & 0xfff0) == 0xf120) {
                        return false;
                    }
                    if((Config.Window.TrayOnClose &&
                            ((((int) msg.WParam) == 0xf060) || (((int) msg.WParam) == 0xf063))) &&
                                    (ModifierKeys != Keys.Shift)) {
                        MinimizeToTray();
                        return true;
                    }
                    if(!QTUtility.IsXP || ((((int) msg.WParam) != 0xf060) && (((int) msg.WParam) != 0xf063))) {
                        return false;
                    }
                    WindowUtils.CloseExplorer(ExplorerHandle, 3);
                    return true;

                case WM.POWERBROADCAST:
                    if(((int) msg.WParam) == 7) {
                        OnAwake();
                    }
                    return false;

                case WM.DEVICECHANGE:
                    if(((int) msg.WParam) == 0x8004) {
                        DEV_BROADCAST_HDR dev_broadcast_hdr = 
                            (DEV_BROADCAST_HDR)Marshal.PtrToStructure(msg.LParam, typeof(DEV_BROADCAST_HDR));
                        if(dev_broadcast_hdr.dbch_devicetype == 2) {
                            DEV_BROADCAST_VOLUME dev_broadcast_volume = (DEV_BROADCAST_VOLUME)Marshal.PtrToStructure(msg.LParam, typeof(DEV_BROADCAST_VOLUME));
                            uint num4 = dev_broadcast_volume.dbcv_unitmask;
                            ushort num5 = 0;
                            while(num5 < 0x1a) {
                                if((num4 & 1) != 0) {
                                    break;
                                }
                                num4 = num4 >> 1;
                                num5 = (ushort) (num5 + 1);
                            }
                            num5 = (ushort) (num5 + 0x41);
                            string str = ((char) num5) + @":\";
                            CloseTabs(tabControl1.TabPages.Where(item =>
                                    item.CurrentPath.PathStartsWith(str)).ToList(), true);
                            if(tabControl1.TabCount == 0) {
                                WindowUtils.CloseExplorer(ExplorerHandle, 2);
                            }
                        }
                    }
                    return false;

                case WM.PARENTNOTIFY:
                    switch((((int)msg.WParam) & 0xffff)) {
                        case WM.LBUTTONDOWN:
                        case WM.RBUTTONDOWN:
                        case WM.MBUTTONDOWN:
                        case WM.XBUTTONDOWN:
                            HideTabSwitcher(false);
                            break;
                    }
                    return false;


                case WM.APPCOMMAND:
                    // some mouse drivers and utilities have the extra buttons send WM_APPCOMMAND.
                    const int APPCOMMAND_BROWSER_BACKWARD = 1;
                    const int APPCOMMAND_BROWSER_FORWARD = 2;
                    const int APPCOMMAND_CLOSE = 31;
                    const int FAPPCOMMAND_MOUSE = 0x8000;
                    const int FAPPCOMMAND_MASK = 0xF000;

                    //GET_APPCOMMAND_LPARAM(lParam) ((short)(HIWORD(lParam) & ~FAPPCOMMAND_MASK))
                    //GET_DEVICE_LPARAM(lParam)     ((WORD)(HIWORD(lParam) & FAPPCOMMAND_MASK))
                    int command = ((((int)(long)msg.LParam) >> 16) & 0xFFFF) & ~FAPPCOMMAND_MASK;
                    int device = ((((int)(long)msg.LParam) >> 16) & 0xFFFF) & FAPPCOMMAND_MASK;
                    bool fProcess = device != FAPPCOMMAND_MOUSE;
                    BindAction action;

                    switch(command) {
                        case APPCOMMAND_BROWSER_BACKWARD:
                            QTUtility2.log("APPCOMMAND_BROWSER_BACKWARD");
                            if(fProcess) {
                                MouseChord chord = QTUtility.MakeMouseChord(MouseChord.X1, ModifierKeys);
                               // QTUtility2.log("APPCOMMAND_BROWSER_BACKWARD fProcess chord " + chord);
                                if(Config.Mouse.GlobalMouseActions.TryGetValue(chord, out action)) {
                                  //  QTUtility2.log("APPCOMMAND_BROWSER_BACKWARD fProcess DoBindAction " + action);
                                    DoBindAction(action);
                                }
                            }
                            else
                            {

                            }
                            return true;

                        case APPCOMMAND_BROWSER_FORWARD:
                            QTUtility2.log("APPCOMMAND_BROWSER_FORWARD");
                            if(fProcess) {
                                MouseChord chord = QTUtility.MakeMouseChord(MouseChord.X2, ModifierKeys);
                                if(Config.Mouse.GlobalMouseActions.TryGetValue(chord, out action)) {
                                    DoBindAction(action);
                                }
                            }
                            return true;

                        case APPCOMMAND_CLOSE:
                            QTUtility2.log("APPCOMMAND_CLOSE");
                            WindowUtils.CloseExplorer(ExplorerHandle, 0);
                            return true;
                    }
                    break;
            }
            return false;
        }

        public override void GetBandInfo(uint dwBandID, uint dwViewMode, ref DESKBANDINFO dbi) {
            if((dbi.dwMask & DBIM.ACTUAL) != (0)) {
                dbi.ptActual.X = Size.Width;
                dbi.ptActual.Y = BandHeight;
            }
            if((dbi.dwMask & DBIM.INTEGRAL) != (0)) {
                dbi.ptIntegral.X = -1;
                dbi.ptIntegral.Y = 10;
            }
            if((dbi.dwMask & DBIM.MAXSIZE) != (0)) {
                dbi.ptMaxSize.X = -1;
                dbi.ptMaxSize.Y = BandHeight;
            }
            if((dbi.dwMask & DBIM.MINSIZE) != (0)) {
                dbi.ptMinSize.X = MinSize.Width;
                dbi.ptMinSize.Y = BandHeight;
            }
            if((dbi.dwMask & DBIM.MODEFLAGS) != (0)) {
                dbi.dwModeFlags = DBIMF.NORMAL;
            }
            if((dbi.dwMask & DBIM.BKCOLOR) != (0)) {
                dbi.dwMask &= ~DBIM.BKCOLOR;
            }
            if((dbi.dwMask & DBIM.TITLE) != (0)) {
                dbi.wszTitle = null;
            }
        }

        private ITravelLogEntry GetCurrentLogEntry() {
            IEnumTravelLogEntry ppenum = null;
            ITravelLogEntry rgElt = null;
            ITravelLogEntry entry3;
            try {
                if(TravelLog.EnumEntries(1, out ppenum) == 0) {
                    ppenum.Next(1, out rgElt, 0);
                }
                entry3 = rgElt;
            }
            catch(Exception exception) {
                QTUtility2.MakeErrorLog(exception);
                entry3 = null;
            }
            finally {
                if(ppenum != null) {
                    QTUtility2.log("ReleaseComObject ppenum");
                    Marshal.ReleaseComObject(ppenum);
                }
            }
            return entry3;
        }

        internal IDLWrapper GetCurrentPIDL() {
            IDLWrapper wrapper = ShellBrowser.GetShellPath();
            if(!wrapper.Available) {
                wrapper.Dispose();
                wrapper = new IDLWrapper(ShellMethods.ShellGetPath2(ExplorerHandle));
                if(!wrapper.Available) {
                    wrapper.Dispose();
                    wrapper = new IDLWrapper(lastCompletedBrowseObjectIDL);
                }
            }
            return wrapper;
        }

        private Cursor GetCursor(bool fDragging) {
            return fDragging ?
                    curTabDrag ?? (curTabDrag = CreateCursor(Resources_Image.imgCurTabDrag)) :
                    curTabCloning ?? (curTabCloning = CreateCursor(Resources_Image.imgCurTabCloning));
        }
        /**
         * new �Ƿ��������أ�
         */
        private  static string GetCommandLine()
        {
            Process cprocess = Process.GetCurrentProcess();
            
            int currentProcessId2 = cprocess.Id;

            int currentProcessId = (int)PInvoke.GetCurrentProcessId();
            var process = Process.GetProcessById( currentProcessId );
            QTUtility2.log(" process command line 0 : " + cprocess.StartInfo.Arguments);
            QTUtility2.log(" process command line 1 : " + process.StartInfo.Arguments);
            

            string result = null;
            try
            {
                var cpid = currentProcessId;
                if (currentProcessId2 != currentProcessId)
                {
                    cpid = currentProcessId2;
                }
                string wmiQuery = string.Format("select CommandLine from Win32_Process where ProcessID ={0}", cpid);
                ManagementObjectSearcher managementObjectSearcher =
                    new ManagementObjectSearcher(wmiQuery);
                ManagementObjectCollection managementObjectCollection = managementObjectSearcher.Get();
                
                foreach (ManagementObject managementObject in managementObjectCollection.Cast<ManagementObject>())
                {
                    result = managementObject["CommandLine"] == null ? "" : managementObject["CommandLine"].ToString();
                }
                QTUtility2.log(" process command line 3 : " + result);
            }
            catch (Exception ex)
            {
                result = "";
            }
            string str = Marshal.PtrToStringUni(PInvoke.GetCommandLine());
            QTUtility2.log(" process command line 2 : " + str);
            return str;
        }

        private static string GetNameToSelectFromCommandLineArg(string str) {
            QTUtility2.log("Marshal.PtrToStringUni ��" + str   );
            if(!string.IsNullOrEmpty(str)) {
                int index = str.IndexOf("/select,", StringComparison.CurrentCultureIgnoreCase);
                if(index == -1) {
                    index = str.IndexOf(",select,", StringComparison.CurrentCultureIgnoreCase);
                }
                if(index != -1) {
                    index += 8;
                    if(str.Length < index) {
                        return string.Empty;
                    }
                    string path = str.Substring(index).Split(new char[] { ',' })[0].Trim().Trim(new char[] { ' ', '"' });
                    try {
                        if(File.Exists(path) || Directory.Exists(path)) {
                            return Path.GetFileName(path);
                        }
                    }
                    catch {
                    }
                }
            }
            return string.Empty;
        }

        private IntPtr GetSearchBand_Edit() {
            IntPtr hwndSearchBand = WindowUtils.FindChildWindow(ExplorerHandle, hwnd => PInvoke.GetClassName(hwnd) == "UniversalSearchBand");
            if(hwndSearchBand != IntPtr.Zero) {
                hwndSearchBand = WindowUtils.FindChildWindow(hwndSearchBand, hwnd =>
                        PInvoke.GetClassName(hwnd) == "Edit" && ((int)PInvoke.GetWindowLongPtr(hwnd, -16) & 0x10000000) != 0);
            }
            return hwndSearchBand;
        }

        public ShellBrowserEx GetShellBrowser() {
            return ShellBrowser;
        }

        private IntPtr GetTravelToolBarWindow32()
        {
            IntPtr hwndTravelBand = WindowUtils.FindChildWindow(ExplorerHandle, hwnd => PInvoke.GetClassName(hwnd) == "TravelBand");
            return hwndTravelBand != IntPtr.Zero
                    ? PInvoke.FindWindowEx(hwndTravelBand, IntPtr.Zero, "ToolbarWindow32", null)
                    : IntPtr.Zero;
        }

        internal struct TVITEM
        {
            public int mask;
            public IntPtr hItem;
            public int state;
            public int stateMask;
            public IntPtr pszText;
            public int cchTextMax;
            public int iImage;
            public int iSelectedImage;
            public int cChildren;
            public IntPtr lParam;
        }

        #region �м������Դ��������ർ��
        unsafe private void Handle_MButtonUp_Tree(MSG msg) 
        {
            QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree msg");
            // IntPtr ptr;
            // if (ShellBrowser.IsFolderTreeVisible(out ptr)   && msg.hwnd == ptr)
            if ( !Explorer.Busy && msg.hwnd != null)
            {
                QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree hwnd  " + msg.hwnd);
                TVHITTESTINFO structure = new TVHITTESTINFO { pt = QTUtility2.PointFromLPARAM(msg.lParam) };
                // HitTest
                IntPtr wParam = PInvoke.SendMessage(msg.hwnd, 0x1111, IntPtr.Zero, ref structure);
                QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree wParam  " + wParam);
                if (wParam != IntPtr.Zero)
                {
                    QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree SendMessage  start " + 4362);
                    Stack<IntPtr> numStack = new Stack<IntPtr>();
                    do
                    {
                        numStack.Push(wParam);
                    }
                    while ((wParam = PInvoke.SendMessage(msg.hwnd, 4362, (IntPtr)3, wParam)) != IntPtr.Zero);
                    QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree SendMessage  end " + 4362);
                    TVITEM* ptrTvitem = stackalloc TVITEM[1];
                    ptrTvitem->mask = 4;
                    int num1 = 0;
                    IntPtr pidl = IntPtr.Zero;
                    QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree numStack.Count " + numStack.Count);
                    while (numStack.Count > 0)
                    {
                        // var hItem = numStack.Pop();
                        ptrTvitem->hItem = numStack.Pop();
                        QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree numStack.Pop hItem " + ptrTvitem->hItem);
                        bool flag1 = num1 == 2 && numStack.Count == 0 && ShellMethods.GetPath(pidl) == "::{031E4825-7B94-4DC3-B131-E946B44C8DD5}";
                        bool flag2 = num1 == 1 && numStack.Count == 0 && ShellMethods.GetPath(pidl) == "::{679F85CB-0220-4080-B29B-5540CC05AAB6}";
                        QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree numStack.Pop flag1 " + flag1 + " flag2 " + flag2);
                        if (!(IntPtr.Zero != PInvoke.SendMessage(msg.hwnd, 4414, (void*)null, (void*)ptrTvitem)))
                        {
                            QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree  while return1 " );
                            return;
                        }

                        if (!(ptrTvitem->lParam != IntPtr.Zero))
                        {
                            QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree  while return2 ");
                            return;
                        }
                        IntPtr num2 = IntPtr.Zero;
                        try
                        {
                            num2 = PInvoke.ILCombine(pidl, *(IntPtr*)*(IntPtr*)(void*)ptrTvitem->lParam);
                            QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree numStack.Pop ILCombine  " + num2 + " pidl " + pidl);
                            if (pidl != IntPtr.Zero)
                            {
                                PInvoke.CoTaskMemFree(pidl);
                                QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree CoTaskMemFree pidl " + pidl);
                            }
                        }
                        catch (Exception e)
                        {
                            QTUtility2.MakeErrorLog(e, "QTTabBarClass Handle_MButtonUp_Tree Exception");
                        }
                       
                            

                        if (flag1 | flag2)
                        {
                            QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree ShellMethods.GetPath " + num2);
                            string path = ShellMethods.GetPath(num2);
                            if (!string.IsNullOrEmpty(path))
                            {

                                QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree path  " + path);
                                uint theOut = 0;
                                PInvoke.SHParseDisplayName(path, IntPtr.Zero, out pidl, 0, out theOut);
                                if (num2 != IntPtr.Zero)
                                    PInvoke.CoTaskMemFree(num2);
                            }
                            else
                            {
                                pidl = num2;
                            }
                        }
                        else
                        {
                            QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree flag1 " + flag1 + " flag2" + flag2);
                            pidl = num2;
                        }
                        ++num1;
                        /*IShellFolder ppv = null;
                        IntPtr ptr1;
                        if (PInvoke.SHBindToParent(pidl, ExplorerGUIDs.IID_IShellFolder, out ppv, out ptr1) ==
                            0)
                        {
                           // IntPtr num2 = PInvoke.ILCombine(pidl, ptr1);
                        }*/
                    }

                    if (pidl != null && pidl != IntPtr.Zero) {
                        QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree pidl " + pidl);
                        MouseChord chord = QTUtility.MakeMouseChord(MouseChord.Middle, ModifierKeys);
                        BindAction action;
                        
                        Keys modKeys = ModifierKeys;
                        // QTUtility2.InitializeTemporaryPaths();
                        bool fBlockSelecting = modKeys == Keys.Shift;
                        bool fCtrl = modKeys == Keys.Control;
                        using (IDLWrapper wrapper = new IDLWrapper(pidl))
                        {
                            if (!wrapper.Available) return;
                            if (wrapper.IsFolder && wrapper.IsReadyIfDrive)
                            {
                                // blockSelecting
                                NavigatedByCode = true;
                                fNowTravelByTree = false; // �ᵼ�¶࿪һ����ǩ by indiff
                               // OpenNewTab(wrapper, true);

                               if (Config.Mouse.ItemActions.TryGetValue(chord, out action))
                               {
                                   if (action == BindAction.ItemOpenInNewTab)
                                   {
                                       OpenNewTab(wrapper, false);
                                       QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree BindAction.ItemOpenInNewTab");
                                   }
                                   else if (action == BindAction.ItemOpenInNewTabNoSel)
                                   {
                                       OpenNewTab(wrapper, true);
                                       QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree BindAction.ItemOpenInNewTabNoSel");
                                   }
                               }
                               else
                               {
                                   if (fCtrl)
                                   {
                                       OpenNewTab(wrapper, true);
                                   }
                                   else
                                   {
                                       OpenNewTab(wrapper, false);
                                   }
                               }
                               QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree IsFolder IsReadyIfDrive " + wrapper.Path);
                            } else if (wrapper.IsLink){
                                if (wrapper.IsLinkToDeadFolder)
                                {
                                    QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree wrapper.IsLinkToDeadFolder");
                                    return;
                                }
                                using (IDLWrapper idlwTarget = new IDLWrapper(ShellMethods.GetLinkTargetIDL(wrapper.Path)))
                                {
                                    if (idlwTarget.IsFolder && idlwTarget.IsReadyIfDrive)
                                    {
                                        NavigatedByCode = true;
                                        fNowTravelByTree = false; // �ᵼ�¶࿪һ����ǩ by indiff
                                        if (Config.Mouse.ItemActions.TryGetValue(chord, out action))
                                        {
                                            if (action == BindAction.ItemOpenInNewTab)
                                            {
                                                OpenNewTab(wrapper, false);
                                                QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree BindAction.ItemOpenInNewTab");
                                            }
                                            else if (action == BindAction.ItemOpenInNewTabNoSel)
                                            {
                                                OpenNewTab(wrapper, true);
                                                QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree BindAction.ItemOpenInNewTabNoSel");
                                            }
                                        }
                                        else
                                        {
                                            if (fCtrl)
                                            {
                                                // StaticReg.CreateWindowIDLs.Add(wrapper.IDL);
                                                OpenNewTab(wrapper, true);
                                            }
                                            else
                                            {
                                                OpenNewTab(wrapper, false);
                                            }
                                        }
                                        QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree IsLink GetLinkTargetIDL" + wrapper.Path);
                                    }
                                }
                            }
                        }
                    }
                    /*QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree wParam " + wParam);
                    int num = (int)PInvoke.SendMessage(msg.hwnd, 0x1127, wParam, (IntPtr)2);
                    if ((num & 2) == 0)
                    {
                        NavigatedByCode = fNowTravelByTree = true;
                        PInvoke.SendMessage(msg.hwnd, 0x110b, (IntPtr)9, wParam);
                        QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree SendMessage  " + wParam);
                    }*/
                }
            }
            else
            {
                QTUtility2.log("QTTabBarClass Explorer.Busy msg.hwnd " + msg.hwnd);
            }
        }
        #endregion

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }

        /*
        public static string GetTreeItemText(IntPtr treeViewHwnd, IntPtr hItem)
        {
            string itemText;

            uint pid;
            uint MY_MAXLVITEMTEXT = 1024;
            PInvoke.GetWindowThreadProcessId(treeViewHwnd, out pid);


            IntPtr process = PInvoke.OpenProcess(ProcessAccessFlags.VirtualMemoryOperation |
                                                 ProcessAccessFlags.VirtualMemoryRead |
                                                 ProcessAccessFlags.VirtualMemoryWrite | 
                                                 ProcessAccessFlags.QueryInformation, false, pid);
            if (process == IntPtr.Zero)
                throw new Exception("Could not open handle to owning process of TreeView", new Win32Exception());

            try
            {
                uint tviSize = Marshal.SizeOf(typeof(TVITEM));
                
                uint textSize = MY_MAXLVITEMTEXT;
                bool isUnicode = IsWindowUnicode(treeViewHwnd);
                if (isUnicode)
                    textSize *= 2;

                IntPtr tviPtr = PInvoke.VirtualAllocEx(process, IntPtr.Zero, tviSize + textSize, AllocationType.Commit, MemoryProtection.ReadWrite);
                if (tviPtr == IntPtr.Zero)
                    throw new Exception("Could not allocate memory in owning process of TreeView", new Win32Exception());

                try
                {
                    IntPtr textPtr = IntPtr.Add(tviPtr, tviSize);

                    TVITEM tvi = new TVITEM();
                    tvi.mask = TVIF_TEXT;
                    tvi.hItem = hItem;
                    tvi.cchTextMax = MY_MAXLVITEMTEXT;
                    tvi.pszText = textPtr;

                    IntPtr ptr = Marshal.AllocHGlobal(tviSize);
                    try
                    {
                        Marshal.StructureToPtr(tvi, ptr, false);
                        if (!PInvoke.WriteProcessMemory(process, tviPtr, ptr,(int) tviSize, IntPtr.Zero))
                            throw new Exception("Could not write to memory in owning process of TreeView", new Win32Exception());
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }

                    if (SendMessage(treeViewHwnd, isUnicode ? TVM_GETITEMW : TVM_GETITEMA, 0, tviPtr) != 1)
                        throw new Exception("Could not get item data from TreeView");

                    ptr = Marshal.AllocHGlobal(textSize);
                    try
                    {
                        int bytesRead;
                        if (!PInvoke.ReadProcessMemory(process, textPtr, ptr, textSize, out bytesRead))
                            throw new Exception("Could not read from memory in owning process of TreeView", new Win32Exception());

                        if (isUnicode)
                            itemText = Marshal.PtrToStringUni(ptr, bytesRead / 2);
                        else
                            itemText = Marshal.PtrToStringAnsi(ptr, bytesRead);
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }
                }
                finally
                {
                    PInvoke.VirtualFreeEx(process, tviPtr, 0, FreeType.Release);
                }
            }
            finally
            {
                PInvoke.CloseHandle(process);
            }

            //char[] arr = itemText.ToCharArray(); //<== use this array to look at the bytes in debug mode

            return itemText;
        } */

        private bool Handle_MButtonUp_Tree(IntPtr hwnd, IntPtr lParam)
        {
            QTUtility2.log("QTTabBarClass Handle_MButtonUp_Tree");
            IntPtr ptr;
            if(ShellBrowser.IsFolderTreeVisible(out ptr) && hwnd == ptr) {
                TVHITTESTINFO structure = new TVHITTESTINFO {pt = QTUtility2.PointFromLPARAM(lParam)};
                QTUtility2.log("QTTabBarClass structure " + structure);
                IntPtr wParam = PInvoke.SendMessage(ptr, 0x1111, IntPtr.Zero, ref structure);
                if(wParam != IntPtr.Zero) {
                    int num = (int)PInvoke.SendMessage(ptr, 0x1127, wParam, (IntPtr)2);
                    if((num & 2) == 0) {
                        NavigatedByCode = fNowTravelByTree = true;
                        PInvoke.SendMessage(ptr, 0x110b, (IntPtr)9, wParam);
                        return true;
                    }
                }
            }
            return false;
        }
        /**
         * ����رղ���
         */
        private bool HandleCLOSE(IntPtr lParam) {
            bool flag = Config.Window.CloseBtnClosesSingleTab;
            bool flag2 = Config.Window.CloseBtnClosesUnlocked;
            List<string> closingPaths = new List<string>();
            int num = (int)lParam;
            switch(num) {
                case 1:
                    closingPaths = CloseAllTabsExcept(null, flag2);
                    if(tabControl1.TabCount > 0) {
                        return true;
                    }
                    break;
                    
                case 2:
                    return false;

                default: {
                        bool flag3 = QTUtility2.PathExists(CurrentTab.CurrentPath);
                        if((QTUtility.IsXP && flag3) && (num == 0)) {
                            return true;
                        }
                        if(!flag3) {
                            CloseTab(CurrentTab, true);
                            return (tabControl1.TabCount > 0);
                        }
                        if(flag2 && !flag) {
                            closingPaths = CloseAllTabsExcept(null);
                            if(tabControl1.TabCount > 0) {
                                return true;
                            }
                            QTUtility.SaveClosing(closingPaths);
                            return false;
                        }
                        Keys modifierKeys = ModifierKeys;
                        if((modifierKeys == (Keys.Control | Keys.Shift)) || !flag) {
                            foreach(QTabItem item2 in tabControl1.TabPages) {
                                closingPaths.Add(item2.CurrentPath);
                                AddToHistory(item2);
                            }
                            QTUtility.SaveClosing(closingPaths);
                            return false;
                        }
                        if(modifierKeys == Keys.Control) {
                            closingPaths = CloseAllTabsExcept(null);
                        }
                        else {
                            closingPaths.Add(CurrentTab.CurrentPath);
                            CloseTab(CurrentTab, false);
                        }
                        if(tabControl1.TabCount > 0) {
                            return true;
                        }
                        QTUtility.SaveClosing(closingPaths);
                        return false;
                    }
            }
            QTUtility.SaveClosing(closingPaths);
            return false;
        }

        internal static int HandleDragEnter(IntPtr hDrop, out string strDraggingDrive, out string strDraggingStartPath) {
            QTUtility2.log("QTTabBarClass HandleDragEnter IsFolder hDrop " + hDrop + 
                           " out string strDraggingDrive, out string strDraggingStartPath" );
            strDraggingDrive = (strDraggingStartPath = null);
            int capacity = (int)PInvoke.DragQueryFile(hDrop, uint.MaxValue, null, 0);
            if(capacity < 1) {
                return -1;
            }
            List<string> list = new List<string>(capacity);
            for(int i = 0; i < capacity; i++) {
                StringBuilder lpszFile = new StringBuilder(260);
                PInvoke.DragQueryFile(hDrop, (uint)i, lpszFile, lpszFile.Capacity);
                if(lpszFile.Length > 0) {
                    list.Add(lpszFile.ToString());
                }
            }
            if(list.Count <= 0) {
                return -1;
            }
            if(list[0].Length < 4) {
                return 2;
            }
            bool flag = true;
            string b = QTUtility2.MakeRootName(list[0]);
            foreach(string str2 in list) {
                if(File.Exists(str2) || Directory.Exists(str2)) {
                    if(str2.Length <= 3) {
                        return 2;
                    }
                    if(!QTUtility2.MakeRootName(str2).PathEquals(b)) {
                        flag = false;
                    }
                    continue;
                }
                return -1;
            }
            if(flag) {
                strDraggingDrive = b;
                strDraggingStartPath = Path.GetDirectoryName(list[0]);
                return 0;
            }
            return 1;
        }

        private static void HandleF5() {
        }

        private void HandleFileDrop(IntPtr hDrop) {
            HideToolTipForDD();
            int capacity = (int)PInvoke.DragQueryFile(hDrop, uint.MaxValue, null, 0);
            if(capacity >= 1) {
                // ��ȡ��ק���ļ��б�
                List<string> listDroppedPaths = new List<string>(capacity);
                for(int i = 0; i < capacity; i++) {
                    StringBuilder lpszFile = new StringBuilder(260);
                    PInvoke.DragQueryFile(hDrop, (uint)i, lpszFile, lpszFile.Capacity);
                    listDroppedPaths.Add(lpszFile.ToString());
                }
                OpenDroppedFolder(listDroppedPaths);
            }
        }

        private bool HandleKEYDOWN(Keys key, bool fRepeat) {
            Keys mkey = key | ModifierKeys;

            switch(key) {
                case Keys.Enter:
                    return false;

                case Keys.Menu:
                    if(!fRepeat && Config.Tabs.ShowCloseButtons && Config.Tabs.CloseBtnsWithAlt) {
                        tabControl1.ShowCloseButton(true);
                    }
                    return false;

                case Keys.ControlKey:
                    if(!fRepeat && NowTabDragging && DraggingTab != null && tabControl1.GetTabMouseOn() == null) {
                        Cursor = GetCursor(false);
                    }
                    break;

                case Keys.Tab:
                    if(Config.Keys.UseTabSwitcher && (mkey & Keys.Control) != Keys.None) {
                        return ShowTabSwitcher((mkey & Keys.Shift) != Keys.None, fRepeat);
                    }
                    break;
            }

            switch(mkey) {
                case Keys.Back:
                    if(!QTUtility.IsXP) {
                        if(listView.HasFocus()) {
                            if(!fRepeat) {
                                if(Config.Tweaks.BackspaceUpLevel) {
                                    QTUtility2.log("QTTabBarClass BackspaceUpLevel UpOneLevel");
                                    UpOneLevel();
                                }
                                else {
                                    NavigateCurrentTab(true);
                                }
                            }
                            return true;
                        }
                    }
                    return false;

                case Keys.Alt | Keys.Left:
                    NavigateCurrentTab(true);
                    return true;

                case Keys.Alt | Keys.Right:
                    NavigateCurrentTab(false);
                    return true;

                case Keys.Alt | Keys.F4:  // ��ݼ���ʽ�رմ���
                    if(!fRepeat) {
                        using (RegistryKey key1 = Registry.CurrentUser.CreateSubKey(RegConst.Root))
                        {
                            string[] list = (from QTabItem item2 in tabControl1.TabPages
                                             where item2.TabLocked
                                             select item2.CurrentPath).ToArray();
                           //  MessageBox.Show(String.Join(",", list));
                            QTUtility2.WriteRegBinary(list, "TabsLocked", key1);
                        }
                        WindowUtils.CloseExplorer(ExplorerHandle, 1);
                    }
                    return true;

                case Keys.F2:
                    if(!Config.Tweaks.F2Selection) {
                        listView.HandleF2();
                    }
                    return false;

            }

            // Ctrl+number = switch to tab #n
            if(((Keys.Control | Keys.NumPad0) <= mkey && mkey <= (Keys.Control | Keys.NumPad9)) ||
                    ((Keys.Control | Keys.D0) <= mkey && mkey <= (Keys.Control | Keys.D9))) {
                int digit;
                if(mkey >= (Keys.Control | Keys.NumPad0)) {
                    digit = (mkey - (Keys.Control | Keys.NumPad0));
                }
                else {
                    digit = (mkey - (Keys.Control | Keys.D0));
                }
                if(digit == 0) {
                    digit = 10;
                }
                if(tabControl1.TabCount >= digit) {
                    tabControl1.SelectTab(digit - 1);
                }
                return true;
            }

            // Check for hotkeys
            int imkey = (int)mkey | QTUtility.FLAG_KEYENABLED;
            for(int i = 0; i < Config.Keys.Shortcuts.Length; ++i) {
                if(Config.Keys.Shortcuts[i] == imkey) {
                    QTUtility2.log("QTTabBarClass imkey " + (BindAction)i);
                    return DoBindAction((BindAction)i);
                }
            }

            // This is important I guess?  Not sure
            if(mkey == (Keys.Control | Keys.W)) return true;

            return false;
        }

        // TODO
        private void HandleLBUTTON_Tree(MSG msg, bool fMouseDown) {
            IntPtr ptr;
            if(ShellBrowser.IsFolderTreeVisible(out ptr) && msg.hwnd == ptr) {
                TVHITTESTINFO structure = new TVHITTESTINFO {pt = QTUtility2.PointFromLPARAM(msg.lParam)};
                IntPtr wParam = PInvoke.SendMessage(ptr, 0x1111, IntPtr.Zero, ref structure);
                if(wParam != IntPtr.Zero) {
                    bool flag;
                    if(fMouseDown) {
                        flag = (((structure.flags != 1) && (structure.flags != 0x10)) && ((structure.flags & 2) == 0)) && ((structure.flags & 4) == 0);
                    }
                    else {
                        flag = ((structure.flags & 2) != 0) || ((structure.flags & 4) != 0);
                    }
                    if(flag) {
                        int num = (int)PInvoke.SendMessage(ptr, 0x1127, wParam, (IntPtr)2);
                        if((num & 2) == 0) {
                            NavigatedByCode = fNowTravelByTree = true;
                        }
                    }
                }
            }
        }        

        private bool HandleMOUSEWHEEL(IntPtr lParam) {
            if(!IsHandleCreated) {
                return false;
            }
            MOUSEHOOKSTRUCTEX mousehookstructex = (MOUSEHOOKSTRUCTEX)Marshal.PtrToStructure(lParam, typeof(MOUSEHOOKSTRUCTEX));
            int y = mousehookstructex.mouseData >> 0x10;
            IntPtr handle = PInvoke.WindowFromPoint(mousehookstructex.mhs.pt);
            Control control = FromHandle(handle);
            bool flag = false;
            if(control != null) {
                IntPtr ptr2;
                DropDownMenuReorderable reorderable = control as DropDownMenuReorderable;
                if(reorderable != null) {
                    if(reorderable.CanScroll) {
                        PInvoke.SendMessage(handle, WM.MOUSEWHEEL, QTUtility2.Make_LPARAM(0, y), QTUtility2.Make_LPARAM(mousehookstructex.mhs.pt));
                    }
                    return true;
                }
                flag = (control == tabControl1) || (handle == Handle);
            }
            if(!flag) {
                Keys modifierKeys = ModifierKeys;
                if((QTUtility.IsXP && modifierKeys == Keys.Control) ||
                        (Config.Tweaks.HorizontalScroll && modifierKeys == Keys.Shift)) {
                    if(listView.MouseIsOverListView()) {
                        switch(modifierKeys) {
                            case Keys.Shift:
                                listView.ScrollHorizontal(y);
                                return true;

                            case Keys.Control:
                                ChangeViewMode(y > 0);
                                return true;
                        }
                    }
                }
                return false;
            }
            if(((tabControl1.TabCount < 2) || (ExplorerHandle != PInvoke.GetForegroundWindow())) || Explorer.Busy) {
                return false;
            }
            int selectedIndex = tabControl1.SelectedIndex;
            if(y < 0) {
                if(selectedIndex == (tabControl1.TabCount - 1)) {
                    tabControl1.SelectedIndex = 0;
                }
                else {
                    tabControl1.SelectedIndex = selectedIndex + 1;
                }
            }
            else if(selectedIndex < 1) {
                tabControl1.SelectedIndex = tabControl1.TabCount - 1;
            }
            else {
                tabControl1.SelectedIndex = selectedIndex - 1;
            }
            return true;
        }

        // todo: clean this crap up...
        private bool HandleItemActivate(Keys modKeys, bool fEnqExec) {
            IntPtr zero = IntPtr.Zero;
            IntPtr ppidl = IntPtr.Zero;
            try {
                Address[] addressArray;
                IDLWrapper wrapper1;
                bool fOpenFirstInTab;
                string str;
                if(ShellBrowser.TryGetSelection(out addressArray, out str, false) && (addressArray.Length > 0)) {
                    List<Address> list = new List<Address>(addressArray);
                    wrapper1 = new IDLWrapper(list[0]);
                    list.RemoveAt(0);
                    addressArray = list.ToArray();
                    fOpenFirstInTab = (addressArray.Length > 0) || (modKeys == Keys.Shift);
                }
                else {
                    return false;
                }
                using(IDLWrapper wrapper = wrapper1) {
                    if((wrapper.Available && wrapper.HasPath) && wrapper.IsReadyIfDrive) {
                        if(wrapper.IsFolder) {
                            if(modKeys == Keys.Control) {
                                if(!wrapper.IsLinkToDeadFolder) {
                                    StaticReg.CreateWindowPaths.AddRange(CreateTMPPathsToOpenNew(addressArray, wrapper.Path));
                                    OpenNewWindow(wrapper);
                                }
                                else {
                                    QTUtility.SoundPlay();
                                }
                            }
                            else if(modKeys == (Keys.Alt | Keys.Control | Keys.Shift)) {
                                DirectoryInfo info = new DirectoryInfo(wrapper.Path);
                                if(info.Exists) {
                                    DirectoryInfo[] directories = info.GetDirectories();
                                    if((directories.Length + tabControl1.TabCount) < 0x41) {
                                        tabControl1.SetRedraw(false);
                                        foreach(DirectoryInfo info2 in directories) {
                                            if(info2.Name != "System Volume Information") {
                                                using(IDLWrapper wrapper2 = new IDLWrapper(info2.FullName)) {
                                                    if(wrapper2.Available && (!wrapper2.IsLink || Directory.Exists(ShellMethods.GetLinkTargetPath(info2.FullName)))) {
                                                        // MessageBox.Show("Open New Tab");
                                                        OpenNewTab(wrapper2, true);
                                                    }
                                                }
                                            }
                                        }
                                        tabControl1.SetRedraw(true);
                                    }
                                    else {
                                        QTUtility.SoundPlay();
                                    }
                                }
                            }
                            else {
                                if(addressArray.Length > 1) {
                                    tabControl1.SetRedraw(false);
                                }
                                try {
                                    if(fOpenFirstInTab) {
                                        OpenNewTab(wrapper, (modKeys & Keys.Shift) == Keys.Shift);
                                    }
                                    else if(!wrapper.IsFileSystemFile) {
                                        ShellBrowser.Navigate(wrapper);
                                    }
                                    else {
                                        return false;
                                    }
                                    for(int i = 0; i < addressArray.Length; i++) {
                                        using(IDLWrapper wrapper3 = new IDLWrapper(addressArray[i].ITEMIDLIST)) {
                                            if(((wrapper3.Available && wrapper3.HasPath) && (wrapper3.IsReadyIfDrive && wrapper3.IsFolder)) && !wrapper3.IsLinkToDeadFolder) {
                                                string path = wrapper3.Path;
                                                if(((path != wrapper.Path) && (path.Length > 0)) && !QTUtility2.IsShellPathButNotFileSystem(path)) {
                                                    OpenNewTab(wrapper3, true);
                                                }
                                            }
                                        }
                                    }
                                }
                                finally {
                                    if(addressArray.Length > 1) {
                                        tabControl1.SetRedraw(true);
                                    }
                                }
                            }
                            return true;
                        }
                        if(wrapper.IsLink) {
                            using(IDLWrapper wrapper4 = new IDLWrapper(ShellMethods.GetLinkTargetIDL(wrapper.Path))) {
                                if(((wrapper4.Available && wrapper4.HasPath) && (wrapper4.IsReadyIfDrive && wrapper4.IsFolder)) && !wrapper.IsLinkToDeadFolder) {
                                    if(modKeys == Keys.Control) {
                                        StaticReg.CreateWindowPaths.AddRange(CreateTMPPathsToOpenNew(addressArray, wrapper.Path));
                                        OpenNewWindow(wrapper4);
                                    }
                                    else {
                                        if(fOpenFirstInTab) {
                                            OpenNewTab(wrapper4, (modKeys & Keys.Shift) == Keys.Shift);
                                        }
                                        else {
                                            ShellBrowser.Navigate(wrapper4);
                                        }
                                        for(int j = 0; j < addressArray.Length; j++) {
                                            using(IDLWrapper wrapper5 = new IDLWrapper(addressArray[j].ITEMIDLIST)) {
                                                if(((wrapper5.Available && wrapper5.HasPath) && (wrapper5.IsReadyIfDrive && wrapper5.IsFolder)) && !wrapper5.IsLinkToDeadFolder) {
                                                    string str3 = wrapper5.Path;
                                                    if(((str3 != wrapper4.Path) && (str3.Length > 0)) && !QTUtility2.IsShellPathButNotFileSystem(str3)) {
                                                        OpenNewTab(wrapper5, true);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    return true;
                                }
                            }
                        }
                        if(fEnqExec) {
                            List<string> list2 = new List<string>();
                            list2.Add(wrapper.Path);
                            foreach(Address address in addressArray) {
                                using(IDLWrapper wrapper6 = new IDLWrapper(address.ITEMIDLIST)) {
                                    if(wrapper6.IsFolder) {
                                        return true;
                                    }
                                    if(wrapper6.HasPath && !wrapper6.IsLinkToDeadFolder) {
                                        list2.Add(wrapper6.Path);
                                    }
                                }
                            }
                            foreach(string str4 in list2) {
                                StaticReg.ExecutedPathsList.Add(str4);
                            }
                        }
                    }
                }
            }
            catch(Exception exception) {
                QTUtility2.MakeErrorLog(exception);
            }
            finally {
                if(zero != IntPtr.Zero) {
                    PInvoke.CoTaskMemFree(zero);
                }
                if(ppidl != IntPtr.Zero) {
                    PInvoke.CoTaskMemFree(ppidl);
                }
            }
            return false;
        }

        private void HideSubDirTip_Tab_Menu() {
            if(subDirTip_Tab != null) {
                subDirTip_Tab.HideMenu();
            }
        }

        private void HideTabSwitcher(bool fSwitch) {
            if((tabSwitcher != null) && tabSwitcher.IsShown) {
                tabSwitcher.HideSwitcher(fSwitch);
                tabControl1.SetPseudoHotIndex(-1);
            }
        }

        private void HideToolTipForDD() {
            tabForDD = null;
            iModKeyStateDD = 0;
            if(toolTipForDD != null) {
                toolTipForDD.Hide(tabControl1);
            }
            if(timerOnTab != null) {
                timerOnTab.Enabled = false;
            }
        }

        private void InitializeComponent() {
            components = new Container();
            buttonNavHistoryMenu = new ToolStripDropDownButton();
            tabControl1 = new QTabControl();
            CurrentTab = new QTabItem(string.Empty, string.Empty, tabControl1);
            contextMenuTab = new ContextMenuStripEx(components, false);
            contextMenuSys = new ContextMenuStripEx(components, false);
            tabControl1.SuspendLayout();
            contextMenuSys.SuspendLayout();
            contextMenuTab.SuspendLayout();
            SuspendLayout();

            bool flag = Config.Tabs.ShowOptionsButton;
            if (flag)
            {
                InitializeOptionsButton();
            }

            buttonNavHistoryMenu.AutoSize = false;
            buttonNavHistoryMenu.DisplayStyle = ToolStripItemDisplayStyle.None;
            buttonNavHistoryMenu.Enabled = false;
            buttonNavHistoryMenu.Size = new Size(13, 0x15);
            buttonNavHistoryMenu.DropDown = new DropDownMenuBase(components, true, true, true);
            buttonNavHistoryMenu.DropDown.ItemClicked += NavigationButton_DropDownMenu_ItemClicked;
            buttonNavHistoryMenu.DropDownOpening += NavigationButtons_DropDownOpening;
            buttonNavHistoryMenu.DropDown.ImageList = QTUtility.ImageListGlobal;

            tabControl1.SetRedraw(false);
            tabControl1.TabPages.Add(CurrentTab);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.ContextMenuStrip = contextMenuTab;
            tabControl1.RefreshOptions(true);
            tabControl1.RowCountChanged += tabControl1_RowCountChanged;
            tabControl1.Deselecting += tabControl1_Deselecting;
            tabControl1.Selecting += tabControl1_Selecting;
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
            tabControl1.GotFocus += Controls_GotFocus;
            tabControl1.MouseEnter += tabControl1_MouseEnter;
            tabControl1.MouseLeave += tabControl1_MouseLeave;
            tabControl1.MouseDown += tabControl1_MouseDown;
            tabControl1.MouseUp += tabControl1_MouseUp;
            tabControl1.MouseMove += tabControl1_MouseMove;
            tabControl1.MouseDoubleClick += tabControl1_MouseDoubleClick;
            tabControl1.ItemDrag += tabControl1_ItemDrag;
            tabControl1.PointedTabChanged += tabControl1_PointedTabChanged;
            tabControl1.TabCountChanged += tabControl1_TabCountChanged;
            tabControl1.CloseButtonClicked += tabControl1_CloseButtonClicked;
            tabControl1.TabIconMouseDown += tabControl1_TabIconMouseDown;
            tabControl1.PlusButtonClicked += tabControl1_PlusButtonClicked;
            
            contextMenuTab.Items.Add(new ToolStripMenuItem());
            contextMenuTab.ShowImageMargin = false;
            contextMenuTab.ItemClicked += contextMenuTab_ItemClicked;
            contextMenuTab.Opening += contextMenuTab_Opening;
            contextMenuTab.Closed += contextMenuTab_Closed;
            contextMenuSys.Items.Add(new ToolStripMenuItem());
            contextMenuSys.ShowImageMargin = false;
            contextMenuSys.ItemClicked += contextMenuSys_ItemClicked;
            contextMenuSys.Opening += contextMenuSys_Opening;
            Controls.Add(tabControl1);
            if(flag) {
                Controls.Add(optionsStrip);
            }
            MinSize = new Size(150, Config.Skin.TabHeight + 2);
            Height = Config.Skin.TabHeight + 2;
            ContextMenuStrip = contextMenuSys;
            // ע�����˫���¼�
            MouseDoubleClick += QTTabBarClass_MouseDoubleClick;
            MouseUp += QTTabBarClass_MouseUp;
            tabControl1.ResumeLayout(false);
            contextMenuSys.ResumeLayout(false);
            contextMenuTab.ResumeLayout(false);
            if(flag) {
                optionsStrip.ResumeLayout(false);
                optionsStrip.PerformLayout();
            }
            ResumeLayout(false);
        }

        private void InitializeInstallation() {
            InitializeOpenedWindow();
            object locationURL = Explorer.LocationURL;
            if(ShellBrowser != null) {
                using(IDLWrapper wrapper = ShellBrowser.GetShellPath()) {
                    if(wrapper.Available) {
                        locationURL = wrapper.Path;
                    }
                }
            }
            QTUtility2.log("QTTabBarClass InitializeInstallation  pDisp :" + null + " locationURL :" + (string)locationURL);
            Explorer_NavigateComplete2(null, ref locationURL);
        }
        /**
         * ��ʼ��������
         */
        private void InitializeOptionsButton() {
            optionsStrip = new ToolStripClasses();
            buttonOptions = new ToolStripButton();
            optionsStrip.SuspendLayout();
            if(!QTUtility.ImageListGlobal.Images.ContainsKey("options")) {
                QTUtility.ImageListGlobal.Images.Add("options", Resources_Image.imgOptionsButton_White);
            }
            optionsStrip.Dock = DockStyle.Right;
            optionsStrip.GripStyle = ToolStripGripStyle.Hidden;
            optionsStrip.Items.Add(buttonOptions);
            optionsStrip.Renderer = new ToolbarRenderer();
            optionsStrip.Width = 50;
            optionsStrip.BackColor = Color.Transparent;

            buttonOptions.AutoSize = false;
            buttonOptions.ForeColor = Config.Skin.ToolBarTextColor;
            buttonOptions.BackColor = Config.Skin.RebarColor;
            buttonOptions.DisplayStyle = ToolStripItemDisplayStyle.Image;
            buttonOptions.Text = buttonOptions.ToolTipText = "Options";
            buttonOptions.Enabled = true;
            buttonOptions.Image = QTUtility.ImageListGlobal.Images["options"];
            buttonOptions.Click += OptionsButton_Click;
        }
        /**
         * ��ʼ���Ѿ��򿪵Ĵ���
         */
        private void InitializeOpenedWindow() {
            IsShown = true;
            InstanceManager.PushTabBarInstance(this);
            //  ��װ����
            QTUtility2.log("QTTabBarClass InitializeOpenedWindow  InstallHooks");
            InstallHooks();
            
            if(QTUtility.WindowAlpha < 0xff) {
                QTUtility2.log("QTTabBarClass SetWindowLongPtr SetLayeredWindowAttributes");
                PInvoke.SetWindowLongPtr(ExplorerHandle, -20, PInvoke.Ptr_OP_OR(PInvoke.GetWindowLongPtr(ExplorerHandle, -20), 0x80000));
                PInvoke.SetLayeredWindowAttributes(ExplorerHandle, 0, QTUtility.WindowAlpha, 2);
            }

            QTUtility2.log("QTTabBarClass ListViewMonitor ");
            listViewManager = new ListViewMonitor(ShellBrowser, ExplorerHandle, Handle);
            listViewManager.ListViewChanged += ListViewMonitor_ListViewChanged;
            listViewManager.Initialize();

            IntPtr hwndBreadcrumbBar = WindowUtils.FindChildWindow(ExplorerHandle, hwnd => PInvoke.GetClassName(hwnd) == "Breadcrumb Parent");
            if(hwndBreadcrumbBar != IntPtr.Zero) {
                hwndBreadcrumbBar = PInvoke.FindWindowEx(hwndBreadcrumbBar, IntPtr.Zero, "ToolbarWindow32", null);
                if(hwndBreadcrumbBar != IntPtr.Zero) {
                    breadcrumbBar = new BreadcrumbBar(hwndBreadcrumbBar);
                    QTUtility2.log("QTTabBarClass BreadcrumbBar set FolderLinkClicked ");
                    breadcrumbBar.ItemClicked += FolderLinkClicked;
                }
            }
            // SysTreeView32
        }

        private static void InitializeStaticFields() {
            fInitialized = true;
            // ����DPI���� indiff
            PInvoke.SetProcessDPIAware();
            Application.EnableVisualStyles();
        }

        private void InitializeSysMenu(bool fText) {
            bool flag = false;
            if(tsmiUndoClose == null) {
                flag = true;
                // ����ر�
                tsmiUndoClose = new ToolStripMenuItem(QTUtility.ResMain[13]);
                // ����ر�
                tsmiLastActiv = new ToolStripMenuItem(QTUtility.ResMain[14]);
                tsmiExecuted = new ToolStripMenuItem(QTUtility.ResMain[15]);
                tsmiBrowseFolder = new ToolStripMenuItem(QTUtility.ResMain[16] + "...");
                tsmiCloseAllButCurrent = new ToolStripMenuItem(QTUtility.ResMain[17]);
                tsmiCloseWindow = new ToolStripMenuItem(QTUtility.ResMain[18]);
                tsmiOption = new ToolStripMenuItem(QTUtility.ResMain[19]);
                tsmiLockToolbar = new ToolStripMenuItem(QTUtility.ResMain[31]);
                tsmiMergeWindows = new ToolStripMenuItem(QTUtility.ResMain[32]);
                tssep_Sys1 = new ToolStripSeparator();
                tssep_Sys2 = new ToolStripSeparator();
                if (contextMenuSys != null)
                {
                    contextMenuSys.SuspendLayout();
                    contextMenuSys.Items[0].Dispose();
                    contextMenuSys.Items.AddRange(new ToolStripItem[]
                    {
                        tsmiUndoClose, tsmiLastActiv, tsmiExecuted, tssep_Sys1, tsmiBrowseFolder, tsmiCloseAllButCurrent, tsmiCloseWindow, tsmiMergeWindows, tsmiLockToolbar, tssep_Sys2, tsmiOption
                    });
                }

                DropDownMenuReorderable reorderable = new DropDownMenuReorderable(components, true, false);
                reorderable.ImageList = QTUtility.ImageListGlobal;
                DropDownMenuReorderable reorderable2 = new DropDownMenuReorderable(components);
                reorderable2.ReorderEnabled = false;
                reorderable2.MessageParent = Handle;
                reorderable2.ImageList = QTUtility.ImageListGlobal;
                reorderable2.ItemRightClicked += ddmrUndoClose_ItemRightClicked;
                tsmiUndoClose.DropDown = reorderable2;
                tsmiUndoClose.DropDownItemClicked += menuitemUndoClose_DropDownItemClicked;
                DropDownMenuReorderable reorderable3 = new DropDownMenuReorderable(components);
                reorderable3.MessageParent = Handle;
                reorderable3.ItemRightClicked += menuitemExecuted_ItemRightClicked;
                reorderable3.ItemClicked += menuitemExecuted_DropDownItemClicked;
                reorderable3.ImageList = QTUtility.ImageListGlobal;
                tsmiExecuted.DropDown = reorderable3;
                tssep_Sys1.Enabled = false;
                tssep_Sys2.Enabled = false;
                if (contextMenuSys != null)
                {
                    contextMenuSys.ResumeLayout(false);
                }
            }
            if(!flag && fText) {
                tsmiUndoClose.Text = QTUtility.ResMain[13];
                tsmiLastActiv.Text = QTUtility.ResMain[14];
                tsmiExecuted.Text = QTUtility.ResMain[15];
                tsmiBrowseFolder.Text = QTUtility.ResMain[16] + "...";
                tsmiCloseAllButCurrent.Text = QTUtility.ResMain[17];
                tsmiCloseWindow.Text = QTUtility.ResMain[18];
                tsmiOption.Text = QTUtility.ResMain[19];
                tsmiLockToolbar.Text = QTUtility.ResMain[31];
                tsmiMergeWindows.Text = QTUtility.ResMain[32];
            }
        }

        // ��ʼ����ǩ���Ҽ��˵�
        private void InitializeTabMenu(bool fText) {
            try
            {
               // MessageBox.Show( "InitializeTabMenu");
                bool flag = false;
                if (tsmiClose == null)
                {
                    flag = true;
                    tsmiClose = new ToolStripMenuItem(QTUtility.ResMain[0]);
                    tsmiCloseRight = new ToolStripMenuItem(QTUtility.ResMain[1]);
                    tsmiCloseLeft = new ToolStripMenuItem(QTUtility.ResMain[2]);
                    tsmiCloseAllButThis = new ToolStripMenuItem(QTUtility.ResMain[3]);
                    tsmiLockThis = new ToolStripMenuItem(QTUtility.ResMain[6]);
                    tsmiCloneThis = new ToolStripMenuItem(QTUtility.ResMain[7]);
                    tsmiCreateWindow = new ToolStripMenuItem(QTUtility.ResMain[8]);
                    tsmiCopy = new ToolStripMenuItem(QTUtility.ResMain[9]);
                    tsmiProp = new ToolStripMenuItem(QTUtility.ResMain[10]);
                    tsmiHistory = new ToolStripMenuItem(QTUtility.ResMain[11]);
                    tsmiTabOrder = new ToolStripMenuItem(QTUtility.ResMain[27]);

                    /** add by qwop 2012-07-13.*/
                    int len = QTUtility.ResMain.Length;
                    tsmiOpenCmd = new ToolStripMenuItem(QTUtility.ResMain[len - 1]);
                    enableApiHook = new ToolStripMenuItem("Enable Image Hook");
                    /** add by qwop 2012-07-13.*/


                    menuTextBoxTabAlias = new ToolStripTextBox();
                    tssep_Tab1 = new ToolStripSeparator();
                    tssep_Tab2 = new ToolStripSeparator();
                    tssep_Tab3 = new ToolStripSeparator();
                    contextMenuTab.SuspendLayout();
                    contextMenuTab.Items[0].Dispose();
                    contextMenuTab.Items.AddRange(new ToolStripItem[] { 
                        tsmiClose, tsmiCloseRight, tsmiCloseLeft, tsmiCloseAllButThis, 
                        tssep_Tab1, tssep_Tab2, tsmiLockThis, 
                        tsmiCloneThis, tsmiCreateWindow, tsmiCopy, tsmiTabOrder, tssep_Tab3, tsmiProp,
                        tsmiHistory,
                        tsmiOpenCmd,
                        // enableApiHook
                    });

                    tsmiHistory.DropDown = new DropDownMenuBase(components, true, true, true);
                    tsmiHistory.DropDownItemClicked += menuitemHistory_DropDownItemClicked;
                    (tsmiHistory.DropDown).ImageList = QTUtility.ImageListGlobal;
                    menuTextBoxTabAlias.Text = menuTextBoxTabAlias.ToolTipText = QTUtility.ResMain[26];
                    menuTextBoxTabAlias.GotFocus += menuTextBoxTabAlias_GotFocus;
                    menuTextBoxTabAlias.LostFocus += menuTextBoxTabAlias_LostFocus;
                    menuTextBoxTabAlias.KeyPress += menuTextBoxTabAlias_KeyPress;
                    tsmiTabOrder.DropDown = new ContextMenuStripEx(components, false);
                    tssep_Tab1.Enabled = false;
                    tssep_Tab2.Enabled = false;
                    tssep_Tab3.Enabled = false;
                    contextMenuTab.ResumeLayout(false);
                }
                if (!flag && fText)
                {
                    tsmiClose.Text = QTUtility.ResMain[0];
                    tsmiCloseRight.Text = QTUtility.ResMain[1];
                    tsmiCloseLeft.Text = QTUtility.ResMain[2];
                    tsmiCloseAllButThis.Text = QTUtility.ResMain[3];
                    tsmiLockThis.Text = QTUtility.ResMain[6];
                    tsmiCloneThis.Text = QTUtility.ResMain[7];
                    tsmiCreateWindow.Text = QTUtility.ResMain[8];
                    tsmiCopy.Text = QTUtility.ResMain[9];
                    tsmiProp.Text = QTUtility.ResMain[10];
                    tsmiHistory.Text = QTUtility.ResMain[11];
                    tsmiTabOrder.Text = QTUtility.ResMain[27];
                    menuTextBoxTabAlias.Text = menuTextBoxTabAlias.ToolTipText = QTUtility.ResMain[26];
                }
            }
            catch (Exception e) {
                QTUtility2.MakeErrorLog(e);
            }
        }

        // ��װ����
        private void InstallHooks() {
            hookProc_Key = new HookProc(CallbackKeyboardProc);
            hookProc_Mouse = new HookProc(CallbackMouseProc);
            hookProc_GetMsg = new HookProc(CallbackGetMsgProc);
            int currentThreadId = PInvoke.GetCurrentThreadId();
            hHook_Key = PInvoke.SetWindowsHookEx(2, hookProc_Key, IntPtr.Zero, currentThreadId);
            hHook_Mouse = PInvoke.SetWindowsHookEx(7, hookProc_Mouse, IntPtr.Zero, currentThreadId);
            hHook_Msg = PInvoke.SetWindowsHookEx(3, hookProc_GetMsg, IntPtr.Zero, currentThreadId);
            explorerController = new NativeWindowController(ExplorerHandle);
            explorerController.MessageCaptured += explorerController_MessageCaptured;
            if(ReBarHandle != IntPtr.Zero) {
                rebarController = new RebarController(this, ReBarHandle, BandObjectSite as IOleCommandTarget);
            }
            if(!QTUtility.IsXP)
            {
                TravelToolBarHandle = GetTravelToolBarWindow32();
                if(TravelToolBarHandle != IntPtr.Zero)
                {
                    travelBtnController = new NativeWindowController(TravelToolBarHandle);
                    travelBtnController.MessageCaptured += travelBtnController_MessageCaptured;
                }
            }
            dropTargetWrapper = new DropTargetWrapper(this);
            dropTargetWrapper.DragFileEnter += dropTargetWrapper_DragFileEnter;
            dropTargetWrapper.DragFileOver += dropTargetWrapper_DragFileOver;
            dropTargetWrapper.DragFileLeave += dropTargetWrapper_DragFileLeave;
            dropTargetWrapper.DragFileDrop += dropTargetWrapper_DragFileDrop;
        }



        

        private bool ListView_SelectionActivated(Keys modKeys) {
            QTUtility2.log("ListView_SelectionActivated");
            if(timerSelectionChanged != null) {
                timerSelectionChanged.Enabled = false;
            }
            int num = ShellBrowser.GetSelectedCount();
            bool fEnqExec = Config.Misc.KeepRecentFiles;
            return (fEnqExec || num != 1 || (modKeys != Keys.None && modKeys != Keys.Alt)) &&
                    HandleItemActivate(modKeys, fEnqExec);
        }

        private bool ListView_MiddleClick(Point pt) {
            MouseChord chord = QTUtility.MakeMouseChord(MouseChord.Middle, ModifierKeys);
            BindAction action;
            if(Config.Mouse.MarginActions.TryGetValue(chord, out action)) {
                QTUtility2.log("ListView_MiddleClick " + action);
                if(listView.PointIsBackground(pt, false)) {
                    return DoBindAction(action);
                }
            }
            if(Config.Mouse.ItemActions.TryGetValue(chord, out action)) {
                int index = listView.HitTest(pt, false);
                if(index <= -1) {
                    return false;
                }
                using(IDLWrapper wrapper = ShellBrowser.GetItem(index)) {
                    QTUtility2.log("QTTabBarClass ListView_MiddleClick " + action);
                    return DoBindAction(action, false, null, wrapper);
                }
            }
            return false;
        }

        private bool ListView_MouseActivate(ref int result) {
            // The purpose of this is to prevent accidentally
            // renaming an item when clicking out of a SubDirTip menu.
            bool ret = false;
            if(listView.SubDirTipMenuIsShowing() || (subDirTip_Tab != null && subDirTip_Tab.MenuIsShowing)) {
                if(ShellBrowser.GetSelectedCount() == 1 && listView.HotItemIsSelected()) {
                    result = 2;
                    listView.HideSubDirTipMenu();
                    HideSubDirTip_Tab_Menu();
                    listView.SetFocus();
                    ret = true;
                }
            }
            listView.RefreshSubDirTip(true);
            return ret;
        }

        private bool ListView_DoubleClick(Point pt) {
            MouseChord chord = QTUtility.MakeMouseChord(MouseChord.Double, ModifierKeys);
            BindAction action;
            if(Config.Mouse.MarginActions.TryGetValue(chord, out action) && listView.PointIsBackground(pt, false)) {
                QTUtility2.log("ListView_DoubleClick " + action);
                DoBindAction(action);
                return true;
            }
            return false;
        }

        private void ListView_EndLabelEdit(LVITEM item) {
            if(item.pszText == IntPtr.Zero) return;
            using(IDLWrapper wrapper = ShellBrowser.GetItem(item.iItem)) {
                if(wrapper.DisplayName != Marshal.PtrToStringUni(item.pszText)) {
                    HandleF5();
                }
            }
        }

        private void ListViewMonitor_ListViewChanged(object sender, EventArgs args) {
            if (listViewManager != null) // �޸���ָ������ by indiff
            {
                listView = listViewManager.CurrentListView;
                ExtendedListViewCommon elvc = listView as ExtendedListViewCommon;
                if (elvc != null)
                {
                    //elvc.ItemCountChanged += ListView_ItemCountChanged;
                    elvc.SelectionActivated += ListView_SelectionActivated;
                    //elvc.SelectionChanged += ListView_SelectionChanged;
                    elvc.MiddleClick += ListView_MiddleClick;
                    elvc.DoubleClick += ListView_DoubleClick;
                    elvc.EndLabelEdit += ListView_EndLabelEdit;
                    elvc.MouseActivate += ListView_MouseActivate;
                    elvc.SubDirTip_MenuItemClicked += subDirTip_MenuItemClicked;
                    elvc.SubDirTip_MenuItemRightClicked += subDirTip_MenuItemRightClicked;
                    elvc.SubDirTip_MultipleMenuItemsClicked += subDirTip_MultipleMenuItemsClicked;
                    elvc.SubDirTip_MultipleMenuItemsRightClicked += subDirTip_MultipleMenuItemsRightClicked;
                }
            }
            HandleF5();
        }
          
        private string MakeTravelBtnTooltipText(bool fBack) {
            string path = string.Empty;
            if(fBack) {
                string[] historyBack = CurrentTab.GetHistoryBack();
                if(historyBack.Length > 1) {
                    path = historyBack[1];
                }
            }
            else {
                string[] historyForward = CurrentTab.GetHistoryForward();
                if(historyForward.Length > 0) {
                    path = historyForward[0];
                }
            }
            if(path.Length > 0) {
                string str2 = QTUtility2.MakePathDisplayText(path, false);
                if(!string.IsNullOrEmpty(str2)) {
                    return str2;
                }
            }
            return path;
        }

        private void menuitemExecuted_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            try {
                string toolTipText = e.ClickedItem.ToolTipText;
                ProcessStartInfo startInfo = new ProcessStartInfo(toolTipText);
                startInfo.WorkingDirectory = Path.GetDirectoryName(toolTipText);
                startInfo.ErrorDialog = true;
                startInfo.ErrorDialogParentHandle = ExplorerHandle;
                Process.Start(startInfo);
                StaticReg.ExecutedPathsList.Add(toolTipText);
            }
            catch {
                QTUtility.SoundPlay();
            }
        }

        private void menuitemExecuted_ItemRightClicked(object sender, ItemRightClickedEventArgs e) {
            using(IDLWrapper wrapper = new IDLWrapper(e.ClickedItem.ToolTipText)) {
                e.HRESULT = shellContextMenu.Open(wrapper, e.IsKey ? e.Point : MousePosition, ((DropDownMenuReorderable)sender).Handle, true);
            }
            if(e.HRESULT == 0xffff) {
                StaticReg.ExecutedPathsList.Remove(e.ClickedItem.ToolTipText);
                e.ClickedItem.Dispose();
            }
        }

        private void menuitemHistory_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            QMenuItem clickedItem = e.ClickedItem as QMenuItem;
            if((ContextMenuedTab != null) && (clickedItem != null)) {
                MenuItemArguments menuItemArguments = clickedItem.MenuItemArguments;
                switch(ModifierKeys) {
                    case Keys.Shift:
                        CloneTabButton(ContextMenuedTab, null, true, -1);
                        NavigateToHistory(menuItemArguments.Path, menuItemArguments.IsBack, menuItemArguments.Index);
                        return;

                    case Keys.Control: {
                            using(IDLWrapper wrapper = new IDLWrapper(menuItemArguments.Path)) {
                                OpenNewWindow(wrapper);
                                return;
                            }
                        }
                    default:
                        tabControl1.SelectTab(ContextMenuedTab);
                        NavigateToHistory(menuItemArguments.Path, menuItemArguments.IsBack, menuItemArguments.Index);
                        return;
                }
            }
        }

        private void menuitemTabOrder_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            if(e.ClickedItem.Name == "Name") {
                ReorderTab(0, false);
            }
            else if(e.ClickedItem.Name == "Drive") {
                ReorderTab(1, false);
            }
            else if(e.ClickedItem.Name == "Active") {
                ReorderTab(2, false);
            }
            else if(e.ClickedItem.Name == "Rev") {
                ReorderTab(3, false);
            }
        }

        private void menuitemUndoClose_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            QMenuItem clickedItem = (QMenuItem)e.ClickedItem;
            if(ModifierKeys != Keys.Control) {
                OpenNewTab(clickedItem.Path);
            }
            else {
                using(IDLWrapper wrapper = new IDLWrapper(clickedItem.Path)) {
                    OpenNewWindow(wrapper);
                }
            }
        }

        private void menuTextBoxTabAlias_GotFocus(object sender, EventArgs e) {
            menuTextBoxTabAlias.ForeColor = SystemColors.WindowText;
            if(menuTextBoxTabAlias.TextBox.ImeMode != ImeMode.On) {
                menuTextBoxTabAlias.TextBox.ImeMode = ImeMode.On;
            }
            if(menuTextBoxTabAlias.Text == QTUtility.ResMain[26]) {
                menuTextBoxTabAlias.Text = string.Empty;
            }
        }

        private void menuTextBoxTabAlias_KeyPress(object sender, KeyPressEventArgs e) {
            if(e.KeyChar == '\r') {
                e.Handled = true;
                contextMenuTab.Close(ToolStripDropDownCloseReason.ItemClicked);
            }
        }

        private void menuTextBoxTabAlias_LostFocus(object sender, EventArgs e) {
            string text = menuTextBoxTabAlias.Text;
            if(text.Length == 0) {
                menuTextBoxTabAlias.Text = QTUtility.ResMain[26];
            }
            if((text != QTUtility.ResMain[26]) && (ContextMenuedTab != null)) {
                ContextMenuedTab.Comment = text;
                ContextMenuedTab.RefreshRectangle();
                tabControl1.Refresh();
            }
            menuTextBoxTabAlias.TextBox.SelectionStart = 0;
        }

        private void MergeAllWindows() {
            InstanceManager.PushTabBarInstance(this);
            InstanceManager.TabBarBroadcast(tabbar => {
                QTabItem[] tabs = tabbar.tabControl1.TabPages.Select(tab => tab.Clone(true)).ToArray();
                InstanceManager.InvokeMain(main => {
                    try {
                        main.tabControl1.SetRedraw(false);
                        foreach(QTabItem tab in tabs) {
                            tab.ResetOwner(main.tabControl1);
                            tab.ImageKey = tab.ImageKey; // this is so dumb, but unfortunately necessary
                            // todo: make it less dumb.
                        }
                        QTabItem.CheckSubTexts(main.tabControl1);
                    }
                    finally {
                        main.tabControl1.SetRedraw(true);
                    }
                });
                WindowUtils.CloseExplorer(tabbar.ExplorerHandle, 2, true);
            }, false);
        }

        private void MinimizeToTray() {
            InstanceManager.AddToTrayIcon(Handle, ExplorerHandle, CurrentAddress,
                tabControl1.TabPages.Select(t => t.Text).ToArray(),
                tabControl1.TabPages.Select(t => t.CurrentPath).ToArray());
        }

        private void NavigateBackToTheFuture() {
            IEnumTravelLogEntry ppenum = null;
            ITravelLogEntry rgElt = null;
            try {
                int num;
                if(((TravelLog.EnumEntries(0x20, out ppenum) == 0) && (TravelLog.GetCount(0x20, out num) == 0)) && (num > 0)) {
                    while(ppenum.Next(1, out rgElt, 0) == 0) {
                        if(--num == 0) {
                            break;
                        }
                        if(rgElt != null) {
                            QTUtility2.log("ReleaseComObject rgElt");
                            Marshal.ReleaseComObject(rgElt);
                            rgElt = null;
                        }
                    }
                    if(rgElt != null) {
                        TravelLog.TravelTo(rgElt);
                    }
                }
            }
            catch(Exception exception) {
                QTUtility2.MakeErrorLog(exception);
            }
            finally {
                if(ppenum != null) {
                    QTUtility2.log("ReleaseComObject ppenum");
                    Marshal.ReleaseComObject(ppenum);
                }
                if(rgElt != null) {
                    QTUtility2.log("ReleaseComObject rgElt");
                    Marshal.ReleaseComObject(rgElt);
                }
            }
        }

        internal void NavigateBranchCurrent(int index) {
            NavigateBranches(CurrentTab, index);
        }

        private void NavigateBranches(QTabItem tab, int index) {
            LogData log = tab.Branches[index];
            Keys modifierKeys = ModifierKeys;
            if(modifierKeys == Keys.Control) {
                using(IDLWrapper wrapper = new IDLWrapper(log.IDL)) {
                    if(!wrapper.Available) {
                        ShowMessageNavCanceled(log.Path, false);
                    }
                    else {
                        OpenNewWindow(wrapper);
                    }
                }
            }
            else if(modifierKeys == Keys.Shift) {
                CloneTabButton(tab, log);
            }
            else {
                tabControl1.SelectTab(tab);
                if(IsSpecialFolderNeedsToTravel(log.Path)) {
                    SaveSelectedItems(CurrentTab);
                    NavigatedByCode = true;
                    NavigateToPastSpecialDir(log.Hash);
                }
                else {
                    NavigatedByCode = false;
                    using(IDLWrapper wrapper2 = new IDLWrapper(log.IDL)) {
                        if(!wrapper2.Available) {
                            ShowMessageNavCanceled(log.Path, false);
                        }
                        else {
                            SaveSelectedItems(CurrentTab);
                            ShellBrowser.Navigate(wrapper2);
                        }
                    }
                }
            }
        }

        private bool NavigateCurrentTab(bool fBack) {
            string currentPath = CurrentTab.CurrentPath;
            LogData data = fBack ? CurrentTab.GoBackward() : CurrentTab.GoForward();
            if(string.IsNullOrEmpty(data.Path)) {
                return false;
            }
            if((CurrentTab.TabLocked && !data.Path.Contains("*?*?*")) && !currentPath.Contains("*?*?*")) {
                try {
                    NowTabCloned = true;
                    QTabItem tab = CurrentTab.Clone();
                    AddInsertTab(tab);
                    if(fBack) {
                        CurrentTab.GoForward();
                    }
                    else {
                        CurrentTab.GoBackward();
                    }
                    tabControl1.SelectTab(tab);
                }
                catch(Exception exception) {
                    QTUtility2.MakeErrorLog(exception);
                }
                return true;
            }
            string path = data.Path;
            if(IsSpecialFolderNeedsToTravel(path) && LogEntryDic.ContainsKey(data.Hash)) {
                SaveSelectedItems(CurrentTab);
                NavigatedByCode = true;
                return NavigateToPastSpecialDir(data.Hash);
            }
            using(IDLWrapper wrapper = new IDLWrapper(data.IDL)) {
                if(!wrapper.Available) {
                    CancelFailedNavigation(path, fBack, 1);
                    return false;
                }
                SaveSelectedItems(CurrentTab);
                NavigatedByCode = true;
                return (0 == ShellBrowser.Navigate(wrapper));
            }
        }

        private void NavigateToFirstOrLast(bool fBack) {
            string[] historyBack;
            if(fBack) {
                historyBack = CurrentTab.GetHistoryBack();
            }
            else {
                historyBack = CurrentTab.GetHistoryForward();
            }
            if(historyBack.Length > (fBack ? 1 : 0)) {
                NavigateToHistory(historyBack[historyBack.Length - 1], fBack, historyBack.Length - 1);
            }
        }

        internal void NavigateToHistory(string displayPath, bool fBack, int steps) {
            LogData data = new LogData();
            int countRollback = fBack ? steps : (steps + 1);
            if(fBack) {
                for(int i = 0; i < steps; i++) {
                    data = CurrentTab.GoBackward();
                }
            }
            else {
                for(int j = 0; j < steps + 1; j++) {
                    data = CurrentTab.GoForward();
                }
            }
            if(string.IsNullOrEmpty(data.Path)) {
                CancelFailedNavigation("( Unknown Path )", fBack, countRollback);
            }
            else if(CurrentTab.TabLocked) {
                NowTabCloned = true;
                QTabItem tab = CurrentTab.Clone();
                AddInsertTab(tab);
                if(fBack) {
                    for(int k = 0; k < steps; k++) {
                        CurrentTab.GoForward();
                    }
                }
                else {
                    for(int m = 0; m < (steps + 1); m++) {
                        CurrentTab.GoBackward();
                    }
                }
                tabControl1.SelectTab(tab);
            }
            else if(IsSpecialFolderNeedsToTravel(displayPath)) {
                SaveSelectedItems(CurrentTab);
                NavigatedByCode = true;
                NavigateToPastSpecialDir(data.Hash);
            }
            else {
                using(IDLWrapper wrapper = new IDLWrapper(data.IDL)) {
                    if(!wrapper.Available) {
                        CancelFailedNavigation(displayPath, fBack, countRollback);
                    }
                    else {
                        SaveSelectedItems(CurrentTab);
                        NavigatedByCode = true;
                        ShellBrowser.Navigate(wrapper);
                    }
                }
            }
        }

        private bool NavigateToIndex(bool fBack, int index) {
            string[] historyBack;
            if(index == 0) {
                return false;
            }
            if(fBack) {
                historyBack = CurrentTab.GetHistoryBack();
                if((historyBack.Length - 1) < index) {
                    return false;
                }
            }
            else {
                historyBack = CurrentTab.GetHistoryForward();
                if(historyBack.Length < index) {
                    return false;
                }
            }
            string str = fBack ? historyBack[index] : historyBack[index - 1];
            if(!fBack) {
                index--;
            }
            NavigateToHistory(str, fBack, index);
            return true;
        }

        

        private void NavigationButton_DropDownMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            QMenuItem clickedItem = e.ClickedItem as QMenuItem;
            if(clickedItem != null) {
                MenuItemArguments menuItemArguments = clickedItem.MenuItemArguments;
                switch(ModifierKeys) {
                    case Keys.Shift:
                        CloneTabButton(CurrentTab, null, true, -1);
                        NavigateToHistory(menuItemArguments.Path, menuItemArguments.IsBack, menuItemArguments.Index);
                        return;

                    case Keys.Control: {
                            using(IDLWrapper wrapper = new IDLWrapper(menuItemArguments.Path)) {
                                OpenNewWindow(wrapper);
                                return;
                            }
                        }
                    default:
                        NavigateToHistory(menuItemArguments.Path, menuItemArguments.IsBack, menuItemArguments.Index);
                        return;
                }
            }
        }

        private void OptionsButton_Click(object sender, EventArgs e) {
            OptionsDialog.Open();
        }

        private void NavigationButtons_DropDownOpening(object sender, EventArgs e)
        {
            buttonNavHistoryMenu.DropDown.SuspendLayout();
            while(buttonNavHistoryMenu.DropDownItems.Count > 0)
            {
                buttonNavHistoryMenu.DropDownItems[0].Dispose();
            }
            if((CurrentTab.HistoryCount_Back + CurrentTab.HistoryCount_Forward) > 1)
            {
                buttonNavHistoryMenu.DropDownItems.AddRange(CreateNavBtnMenuItems(true).ToArray());
                buttonNavHistoryMenu.DropDownItems.AddRange(CreateBranchMenu(true, components, tsmiBranchRoot_DropDownItemClicked).ToArray());
            }
            else
            {
                ToolStripMenuItem item = new ToolStripMenuItem("none");
                item.Enabled = false;
                buttonNavHistoryMenu.DropDownItems.Add(item);
            }
            buttonNavHistoryMenu.DropDown.ResumeLayout();
        }

        private void OnAwake() {
        }

        internal void OnMouseDoubleClick() {
            OnMouseDoubleClick(new MouseEventArgs(MouseButtons.Left, 0, 0, 0, 0));
        }

        protected override void OnExplorerAttached() {
            ExplorerHandle = (IntPtr)Explorer.HWND;
            try {
                object obj2;
                object obj3;
                _IServiceProvider bandObjectSite = (_IServiceProvider)BandObjectSite;
                bandObjectSite.QueryService(ExplorerGUIDs.IID_IShellBrowser, ExplorerGUIDs.IID_IUnknown, out obj2);
                ShellBrowser = new ShellBrowserEx((IShellBrowser)obj2);
                HookLibManager.InitShellBrowserHook(ShellBrowser.GetIShellBrowser());
                if(Config.Tweaks.ForceSysListView) {
                    ShellBrowser.SetUsingListView(true);
                }
                bandObjectSite.QueryService(ExplorerGUIDs.IID_ITravelLogStg, ExplorerGUIDs.IID_ITravelLogStg, out obj3);
                TravelLog = (ITravelLogStg)obj3;
            }
            catch(COMException exception) {
                QTUtility2.MakeErrorLog(exception);
            }
            
            Explorer.BeforeNavigate2 += Explorer_BeforeNavigate2;
            Explorer.NavigateComplete2 += Explorer_NavigateComplete2;
            QTUtility2.log("QTTabBarClass set BeforeNavigate2 NavigateComplete2");
        }

        protected override void OnPaintBackground(PaintEventArgs e) {
            if(VisualStyleRenderer.IsSupported) {
                if(bgRenderer == null) {
                    bgRenderer = new VisualStyleRenderer(VisualStyleElement.Rebar.Band.Normal);
                }
                bgRenderer.DrawParentBackground(e.Graphics, e.ClipRectangle, this);
            }
            else {
                if(ReBarHandle != IntPtr.Zero) {
                    int colorref = (int)PInvoke.SendMessage(ReBarHandle, 0x414, IntPtr.Zero, IntPtr.Zero);
                    using(SolidBrush brush = new SolidBrush(QTUtility2.MakeColor(colorref))) {
                        e.Graphics.FillRectangle(brush, e.ClipRectangle);
                        return;
                    }
                }
                base.OnPaintBackground(e);
            }
        }

        private void OpenDroppedFolder(IList<string> listDroppedPaths) {
            // ��ק�ļ��еĲ�����������ק�ļ�
            Keys modKeys = ModifierKeys;
            QTUtility2.InitializeTemporaryPaths();
            bool fBlockSelecting = modKeys == Keys.Shift;
            bool fCtrl = modKeys == Keys.Control;
            bool fOpened = false;

            tabControl1.SetRedraw(false);
            try {
                // ����Ŀ¼����һ���±�ǩ,����ѡ�д�
                foreach(string path in listDroppedPaths.Where(path => !string.IsNullOrEmpty(path))) {
                    try {
                        using(IDLWrapper wrapper = new IDLWrapper(path)) {
                            if(!wrapper.Available) continue;
                            if(wrapper.IsLink) {  // ����Ŀ¼
                                if(wrapper.IsLinkToDeadFolder) continue;
                                using(IDLWrapper idlwTarget = new IDLWrapper(ShellMethods.GetLinkTargetIDL(path))) {
                                    if(idlwTarget.IsFolder && idlwTarget.IsReadyIfDrive) {
                                        IDLWrapper idlwToNavigate = wrapper.IsFolder ? wrapper : idlwTarget;
                                        if(fCtrl) {
                                            StaticReg.CreateWindowIDLs.Add(idlwToNavigate.IDL);
                                        }
                                        else {
                                            OpenNewTab(idlwToNavigate, fBlockSelecting);
                                            fBlockSelecting = true;
                                        }
                                        fOpened = true;
                                    }
                                }
                            }
                            else if(wrapper.IsFolder && wrapper.IsReadyIfDrive) {
                                //  ������Ŀ¼�ļ���
                                if(fCtrl) {
                                    StaticReg.CreateWindowIDLs.Add(wrapper.IDL);
                                }
                                else {
                                    OpenNewTab(wrapper, fBlockSelecting);
                                    fBlockSelecting = true;
                                }
                                fOpened = true;
                            }
                        }
                    }
                    catch(Exception e) {
                        QTUtility2.MakeErrorLog(e, "OpenDroppedFolder");
                    }
                }
            }
            finally {
                // ��������Զ�ˢ�� bRedraw
                tabControl1.SetRedraw(true);
            }

            if(fCtrl) {
                // ctrl+�϶���һ���´���
                if(StaticReg.CreateWindowIDLs.Count > 0) {
                    byte[] first = StaticReg.CreateWindowIDLs[0];
                    StaticReg.CreateWindowIDLs.RemoveAt(0);
                    using(IDLWrapper idlw = new IDLWrapper(first)) {
                        ShellBrowser.Navigate(idlw, SBSP.NEWBROWSER);
                    }
                }
            }
        }

        private bool OpenNewTab(string path, bool blockSelecting = false, bool fForceNew = false) {
            using(IDLWrapper wrapper = new IDLWrapper(path)) {
                if(wrapper.Available) {
                    return OpenNewTab(wrapper, blockSelecting, fForceNew);
                }
            }
            return false;
        }

        internal bool OpenNewTab(IDLWrapper idlwGiven, bool blockSelecting = false, bool fForceNew = false) {
            // Check that the folder exists and is navigable.
            if(idlwGiven == null || !idlwGiven.Available || !idlwGiven.HasPath || !idlwGiven.IsReadyIfDrive || idlwGiven.IsLinkToDeadFolder) {
                QTUtility.SoundPlay();
                return false;
            }

            // If the IDL is a link, resolve it.  Otherwise keep using the one we're given.
            using(IDLWrapper idlwLink = idlwGiven.ResolveTargetIfLink()) {
                IDLWrapper idlw = idlwLink ?? idlwGiven;

                // Recheck a few things
                if(!idlw.Available || !idlw.HasPath || !idlw.IsReadyIfDrive || !idlw.IsFolder) {
                    QTUtility.SoundPlay();
                    return false;
                }

                if(blockSelecting) {
                    NowTabsAddingRemoving = true;
                }
                try {
                    // Check if it's already open
                    if(!fForceNew && Config.Tabs.NeverOpenSame) {
                        QTabItem tabPage = tabControl1.TabPages.FirstOrDefault(
                                item2 => item2.CurrentPath.PathEquals(idlw.Path));
                        if(tabPage != null) {
                            if(Config.Tabs.ActivateNewTab) {
                                tabControl1.SelectTab(tabPage);
                            }
                            return false;
                        }
                    }

                    // TODO
                    // This entire block is a mystery to me, and I think it should be
                    // removed. It's gone in Quizo's version.
                    string path = idlw.Path;
                    if(!idlw.Special && !path.StartsWith("::")) {
                        string directoryName = Path.GetDirectoryName(path);
                        if(!string.IsNullOrEmpty(directoryName)) {
                            using(IDLWrapper wrapper = new IDLWrapper(directoryName)) {
                                if(wrapper.Special && idlw.Available) {
                                    IShellFolder ppv = null;
                                    try {
                                        IntPtr ptr;
                                        if(PInvoke.SHBindToParent(idlw.PIDL, ExplorerGUIDs.IID_IShellFolder, out ppv, out ptr) == 0) {
                                            using(IDLWrapper wrapper2 = new IDLWrapper(PInvoke.ILCombine(wrapper.PIDL, ptr))) {
                                                if(wrapper2.Available && wrapper2.HasPath) {
                                                    if(!blockSelecting && Config.Tabs.ActivateNewTab) {
                                                        NowTabCreated = true;
                                                        tabControl1.SelectTab(CreateNewTab(wrapper2));
                                                    }
                                                    else {
                                                        CreateNewTab(wrapper2);
                                                        QTabItem.CheckSubTexts(tabControl1);
                                                    }
                                                    return true;
                                                }
                                            }
                                        }
                                    }
                                    catch {
                                    }
                                    finally {
                                        if(ppv != null) {
                                            QTUtility2.log("ReleaseComObject ppv");
                                            Marshal.ReleaseComObject(ppv);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // This should work for everything...
                    if(!blockSelecting && Config.Tabs.ActivateNewTab) {
                        NowTabCreated = true;
                        tabControl1.SelectTab(CreateNewTab(idlw));
                    }
                    else {
                        CreateNewTab(idlw);
                        QTabItem.CheckSubTexts(tabControl1);
                    }
                }
                finally {
                    if(blockSelecting) {
                        NowTabsAddingRemoving = false;
                    }
                }
            }
            return true;
        }

        internal void OpenNewTabOrWindow(IDLWrapper idlw, bool fNeedsPulse = false) {
            Keys modKeys = ModifierKeys;
            if((modKeys & Keys.Control) == 0) {
                OpenNewTab(idlw, (modKeys & Keys.Shift) == Keys.Shift);
                WindowUtils.BringExplorerToFront(ExplorerHandle);
                if(fNeedsPulse) {
                    fNeedsNewWindowPulse = true;
                }
            }
            else {
                OpenNewWindow(idlw);
            }
        }

        internal void OpenNewWindow(IDLWrapper idlwGiven) {
            // Check that the folder exists and is navigable.
            if(idlwGiven == null || !idlwGiven.Available || !idlwGiven.HasPath || !idlwGiven.IsReadyIfDrive || idlwGiven.IsLinkToDeadFolder) {
                QTUtility.SoundPlay();
                return;
            }
            
            // If the IDL is a link, resolve it.  Otherwise keep using the one we're given.
            using(IDLWrapper idlwLink = idlwGiven.ResolveTargetIfLink()) {
                IDLWrapper idlw = idlwLink ?? idlwGiven;

                // Recheck a few things
                if(!idlw.Available || !idlw.HasPath || !idlw.IsReadyIfDrive || !idlw.IsFolder) {
                    QTUtility.SoundPlay();
                    return;
                }

                bool isFolderTreeVisible = ShellBrowser.IsFolderTreeVisible();    
                bool fSameAsCurrent;
                using(IDLWrapper wrapper = ShellBrowser.GetShellPath()) {
                    fSameAsCurrent = (wrapper == idlw);
                }

                // There's some weird magic going on here, but it's apparently necessary.
                // TODO: understand it
                SBSP wFlags = SBSP.NEWBROWSER;
                if(fSameAsCurrent) {
                    if(isFolderTreeVisible) {
                        if(CheckProcessID(ExplorerHandle, WindowUtils.GetShellTrayWnd()) || WindowUtils.IsExplorerProcessSeparated()) {
                            PInvoke.SetRedraw(ExplorerHandle, false);
                            ShowFolderTree(false);
                            wFlags |= SBSP.EXPLOREMODE;
                            new WaitTimeoutCallback(WaitTimeout).BeginInvoke(200, AsyncComplete_FolderTree, true);
                        }
                        else {
                            QTUtility.fRestoreFolderTree = true;
                        }
                    }
                    else {
                        if(QTUtility.IsXP) {
                            QTUtility.RestoreFolderTree_Hide = true;
                        }
                        wFlags |= SBSP.EXPLOREMODE;
                    }
                }
                else if(isFolderTreeVisible) {
                    QTUtility.fRestoreFolderTree = true;
                }

                QTUtility.fRestoreFolderTree = false;
            }
        }

        private void QTTabBarClass_MouseDoubleClick(object sender, MouseEventArgs e) {
            MouseChord chord = QTUtility.MakeMouseChord(MouseChord.Double, ModifierKeys);
            BindAction action;
            if(Config.Mouse.BarActions.TryGetValue(chord, out action)) {
                QTUtility2.log("QTTabBarClass_MouseDoubleClick " + action);
                DoBindAction(action);
            }
        }

        private void QTTabBarClass_MouseUp(object sender, MouseEventArgs e) {
            MouseChord chord;
            if(e.Button == MouseButtons.Left) {
                chord = QTUtility.MakeMouseChord(MouseChord.Left, ModifierKeys);
            }
            else if(e.Button == MouseButtons.Middle) {
                chord = QTUtility.MakeMouseChord(MouseChord.Middle, ModifierKeys);
            }
            else {
                return;
            }
            BindAction action;
            
            if(Config.Mouse.BarActions.TryGetValue(chord, out action)) {
                QTUtility2.log("QTTabBarClass_MouseUp " + action );
                DoBindAction(action);
            }
        }

        /// <summary>
        /// ˢ����������
        /// </summary>
        internal void RefreshOptions() {
            QTUtility2.log(  "QTTabBarClass RefreshOptions" );
            SuspendLayout();
            tabControl1.SuspendLayout();
            tabControl1.RefreshOptions(false);
            if(Config.Tabs.ShowOptionsButton) {
                if(optionsStrip == null) {
                    InitializeOptionsButton();
                    Controls.Add(optionsStrip);
                }
                else {
                    optionsStrip.SuspendLayout();
                }
                optionsStrip.Dock = DockStyle.Right;
                optionsStrip.ResumeLayout(false);
                optionsStrip.PerformLayout();
            }
            else if(optionsStrip != null) {
                optionsStrip.Dock = DockStyle.None;
            }
            int iType = 0;
            if(Config.Tabs.MultipleTabRows) {
                iType = Config.Tabs.ActiveTabOnBottomRow ? 1 : 2;
            }
            SetBarRows(tabControl1.SetTabRowType(iType));
            rebarController.RefreshBG();
            foreach(QTabItem item in tabControl1.TabPages) {
                item.RefreshRectangle();
            }
            ShellBrowser.SetUsingListView(Config.Tweaks.ForceSysListView);
            tabControl1.ResumeLayout();
            ResumeLayout(true);
        }

        [ComRegisterFunction]
        private static void Register(Type t) {
            // QTUtility2.log(  "QTTabBarClass Register" );
            string name = t.GUID.ToString("B");
            using(RegistryKey key2 = Registry.ClassesRoot.CreateSubKey(@"CLSID\" + name)) {
                key2.SetValue(null, "QTTabBar");
                key2.SetValue("MenuText", "QTTabBar");
                key2.SetValue("HelpText", "QTTabBar");
            }
            using(RegistryKey key3 = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Internet Explorer\Toolbar")) {
                key3.SetValue(name, "QTTabBar");
            }
        }

        private void ReorderTab(int index, bool fDescending) {
            tabControl1.SetRedraw(false);
            try {
                if(index == 3) {
                    if(tabControl1.TabCount > 1) {
                        int indexSource = 0;
                        for(int i = tabControl1.TabCount - 1; indexSource < i; i--) {
                            tabControl1.TabPages.Relocate(indexSource, i);
                            tabControl1.TabPages.Relocate(i - 1, indexSource);
                            indexSource++;
                        }
                    }
                }
                else {
                    int num3 = fDescending ? -1 : 1;
                    for(int j = 0; j < (tabControl1.TabCount - 1); j++) {
                        for(int k = tabControl1.TabCount - 1; k > j; k--) {
                            string strA;
                            string strB;
                            if(index == 0) {
                                strA = tabControl1.TabPages[j].Text;
                                strB = tabControl1.TabPages[k].Text;
                            }
                            else if(index == 1) {
                                strA = tabControl1.TabPages[j].CurrentPath;
                                strB = tabControl1.TabPages[k].CurrentPath;
                            }
                            else {
                                int num6 = lstActivatedTabs.IndexOf(tabControl1.TabPages[j]);
                                int num7 = lstActivatedTabs.IndexOf(tabControl1.TabPages[k]);
                                if(((num6 - num7) * num3) < 0) {
                                    tabControl1.TabPages.Relocate(j, k);
                                }
                                continue;
                            }
                            if((string.Compare(strA, strB) * num3) > 0) {
                                tabControl1.TabPages.Relocate(j, k);
                            }
                        }
                    }
                }
            }
            finally {
                tabControl1.SetRedraw(true);
            }
        }

        private void RestoreLastClosed() {
            if(StaticReg.ClosedTabHistoryList.Count <= 0) {
                return;
            }
            Stack<string> stack = new Stack<string>(StaticReg.ClosedTabHistoryList);
            string path = null;
            while(stack.Count > 0) {
                path = stack.Pop();
                if(!tabControl1.TabPages.Any(item => item.CurrentPath.PathEquals(path))) {
                    OpenNewTab(path);
                    return;
                }
            }
            if(!path.PathEquals(CurrentAddress)) {
                OpenNewTab(path);
            }
        }

        // �ָ���ǩ
        private void RestoreTabsOnInitialize(int iIndex, string openingPath) {
            QTUtility2.log(  "QTTabBarClass RestoreTabsOnInitialize" );
            QTUtility.RefreshLockedTabsList();
            // TODO: unjank
            TabPos num = Config.Tabs.NewTabPosition;
            Config.Tabs.NewTabPosition = TabPos.Rightmost;
            try {
                if(iIndex == 1) {
                    /*
                    foreach(string str in StaticReg.LockedTabsToRestoreList) {
                        bool flag = false;
                        foreach(QTabItem item2 in tabControl1.TabPages) {
                            if(item2.CurrentPath == str) {
                                if(item2 == CurrentTab) {
                                    fNowRestoring = true;
                                }
                                else {
                                    item2.TabLocked = true;
                                    flag = true;
                                }
                                break;
                            }
                        }
                        if (flag)
                        {  // �ж��Ƿ����� !flag
                           // if(str != openingPath) {
                                using(IDLWrapper wrapper = new IDLWrapper(str)) {
                                    if(wrapper.Available) {
                                        CreateNewTab(wrapper).TabLocked = true;
                                    }
                                    continue;
                                }
                           // }
                            tabControl1.TabPages.Relocate(0, tabControl1.TabCount - 1);
                            fNowRestoring = true;
                        }
                    }
                    */
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegConst.Root, false))
                    {
                        if (key != null)
                        {
                            string[] strArray = ((string)key.GetValue("TabsLocked2", string.Empty)).Split(QTUtility.SEPARATOR_CHAR);
                            if ((strArray.Length > 0) && (strArray[0].Length > 0))
                            {
                                foreach (string str2 in strArray.Where(str2 => str2.Length > 0
                                        && tabControl1.TabPages.All(item3 => item3.CurrentPath != str2)))
                                {
                                    if (str2 == openingPath)
                                    {
                                        tabControl1.TabPages.Relocate(0, tabControl1.TabCount - 1);
                                    }
                                    else
                                    {
                                        using (IDLWrapper wrapper2 = new IDLWrapper(str2))
                                            {
                                                if (wrapper2.Available)
                                                {
                                                   // ֻ�ָ������� indiff
                                                    QTabItem item4 = CreateNewTab(wrapper2);
                                                    item4.TabLocked = true;
                                                }
                                            }
                                         // end of using 
                                    }
                                }
                                fNowRestoring = true;
                            }
                        }
                    }
                }
                else if(iIndex == 0) {
                    using(RegistryKey key = Registry.CurrentUser.OpenSubKey(RegConst.Root, false)) {
                        if(key != null) {
                            string[] strArray = ((string)key.GetValue("TabsOnLastClosedWindow", string.Empty)).Split(QTUtility.SEPARATOR_CHAR);
                            if((strArray.Length > 0) && (strArray[0].Length > 0)) {
                                foreach(string str2 in strArray.Where(str2 => str2.Length > 0
                                        && tabControl1.TabPages.All(item3 => item3.CurrentPath != str2))) {
                                    if(str2 == openingPath) {
                                        tabControl1.TabPages.Relocate(0, tabControl1.TabCount - 1);
                                    }
                                    else {
                                        using(IDLWrapper wrapper2 = new IDLWrapper(str2)) {
                                            if(wrapper2.Available) {
                                                QTabItem item4 = CreateNewTab(wrapper2);
                                                if(StaticReg.LockedTabsToRestoreList.Contains(str2)) {
                                                    item4.TabLocked = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                fNowRestoring = true;
                            }
                        }
                    }
                }
            }
            finally {
                Config.Tabs.NewTabPosition = num;
            }
        }
        

       

        protected override bool ShouldHaveBreak() {
            bool breakBar = true;
            using(RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root)) {
                if(key != null) {
                    breakBar = ((int)key.GetValue("BreakTabBar", 1) == 1);
                }
            }
            return breakBar;
        }

        internal void ShowContextMenu(bool fByKey) {
            contextMenuSys.Show(fByKey ? PointToScreen(Point.Empty) : MousePosition);
        }

        public override void ShowDW(bool fShow) {
            base.ShowDW(fShow);
            if((fShow && !FirstNavigationCompleted) && ((Explorer != null) && (Explorer.ReadyState == tagREADYSTATE.READYSTATE_COMPLETE))) {
                InitializeInstallation();
            }
            if(!fShow) {
                using(RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root)) {
                    key.SetValue("BreakTabBar", BandHasBreak() ? 1 : 0);
                }
            }
        }
        // ��ʾĿ¼��
        private void ShowFolderTree(bool fShow) {
            if(QTUtility.IsXP &&
               (fShow != ShellBrowser.IsFolderTreeVisible())) {
                object pvaClsid = "{EFA24E64-B078-11d0-89E4-00C04FC9E26E}";
                object pvarShow = fShow;
                object pvarSize = null;
                Explorer.ShowBrowserBar(ref pvaClsid, ref pvarShow, ref pvarSize);
            }
        }
        
        // ��ʾ·����MD5
        internal static void ShowMD5(string[] paths) {
            if(md5Form == null) {
                md5Form = new FileHashComputerForm();
            }
            List<string> list = new List<string>();
            if(paths != null) {
                list.AddRange(paths.Where(File.Exists));
            }
            string[] strArray = null;
            if(list.Count > 0) {
                strArray = list.ToArray();
            }
            if(md5Form.InvokeRequired) {
                md5Form.Invoke(new FormMethodInvoker(ShowMD5FormCore), new object[] { strArray });
            }
            else {
                ShowMD5FormCore(strArray);
            }
        }

        private static void ShowMD5FormCore(object paths) {
            md5Form.ShowFileHashForm((string[])paths);
        }

      

        private void ShowSearchBar(bool fShow) {
            QTUtility2.log(  "QTTabBarClass ShowSearchBar fShow: " + fShow);
            if(!QTUtility.IsXP) {
                if(!fShow) {
                    return;
                }
                using(IDLWrapper wrapper = new IDLWrapper(QTUtility.PATH_SEARCHFOLDER)) {
                    if(wrapper.Available) {
                        ShellBrowser.Navigate(wrapper, SBSP.NEWBROWSER);
                    }
                    return;
                }
            }
            object pvaClsid = "{C4EE31F3-4768-11D2-BE5C-00A0C9A83DA1}";
            object pvarShow = fShow;
            object pvarSize = null;
            Explorer.ShowBrowserBar(ref pvaClsid, ref pvarShow, ref pvarSize);
        }

        public AbstractListView GetListView() {
            return listView;
        }
        
        // ��ʾ��Ŀ¼��ʾ��Ϣ
        private void ShowSubdirTip_Tab(QTabItem tab, bool fShow, int offsetX, bool fKey, bool fParent) {
            try {
                if(fShow) {
                    if(Explorer.Busy || string.IsNullOrEmpty(tab.CurrentPath)) {
                        tabControl1.SetSubDirTipShown(false);
                    }
                    else if (QTUtility.IsNetPath(tab.CurrentPath))
                    {
                        tabControl1.SetSubDirTipShown(false);
                    }
                    else {
                        string currentPath = tab.CurrentPath;
                        if(fParent || ShellMethods.TryMakeSubDirTipPath(ref currentPath)) {
                            if(subDirTip_Tab == null) {
                                subDirTip_Tab = new SubDirTipForm(Handle, true, listView);
                                subDirTip_Tab.MenuItemClicked += subDirTip_MenuItemClicked;
                                subDirTip_Tab.MultipleMenuItemsClicked += subDirTip_MultipleMenuItemsClicked;
                                subDirTip_Tab.MenuItemRightClicked += subDirTip_MenuItemRightClicked;
                                subDirTip_Tab.MenuClosed += subDirTip_Tab_MenuClosed;
                                subDirTip_Tab.MultipleMenuItemsRightClicked += subDirTip_MultipleMenuItemsRightClicked;
                            }
                            ContextMenuedTab = tab;
                            Point pnt = tabControl1.PointToScreen(new Point(tab.TabBounds.X + offsetX, fParent ? tab.TabBounds.Top : (tab.TabBounds.Bottom - 3)));
                            if(tab != CurrentTab) {
                                pnt.X += 2;
                            }
                            tabControl1.SetSubDirTipShown(subDirTip_Tab.ShowMenuWithoutShowForm(currentPath, pnt, fParent));
                        }
                        else {
                            tabControl1.SetSubDirTipShown(false);
                            HideSubDirTip_Tab_Menu();
                        }
                    }
                }
                else {
                    HideSubDirTip_Tab_Menu();
                }
            }
            catch(Exception exception) {
                QTUtility2.MakeErrorLog(exception, "tabsubdir");
            }
        }

        private bool ShowTabSwitcher(bool fShift, bool fRepeat) {
            listView.HideSubDirTip();
            listView.HideThumbnailTooltip();
            if(tabControl1.TabCount < 2) {
                return false;
            }
            if(tabSwitcher == null) {
                tabSwitcher = new TabSwitchForm();
                tabSwitcher.Switched += tabSwitcher_Switched;
            }
            if(!tabSwitcher.IsShown) {
                List<PathData> lstPaths = new List<PathData>();
                string str = Config.Tabs.RenameAmbTabs ? " @ " : " : ";
                foreach(QTabItem item in tabControl1.TabPages) {
                    string strDisplay = item.Text;
                    if(!string.IsNullOrEmpty(item.Comment)) {
                        strDisplay += str + item.Comment;
                    }
                    lstPaths.Add(new PathData(strDisplay, item.CurrentPath, item.ImageKey));
                }
                tabSwitcher.ShowSwitcher(ExplorerHandle, tabControl1.SelectedIndex, lstPaths);
            }
            int index = tabSwitcher.Switch(fShift);
            if(!fRepeat || tabControl1.TabCount < 13) {
                tabControl1.SetPseudoHotIndex(index);
            }
            return true;
        }

        /**
         * ��ʾ������Ϣ
         *  shift ��ʾ��ϸ��Ϣ
         */
        private void ShowToolTipForDD(QTabItem tab, int iState, int grfKeyState) {
            if(((tabForDD == null) || (tabForDD != tab)) || (iModKeyStateDD != grfKeyState)) {
                tabForDD = tab;
                iModKeyStateDD = grfKeyState;
                if(timerOnTab == null) {
                    timerOnTab = new Timer(components);
                    timerOnTab.Tick += timerOnTab_Tick;
                }
                timerOnTab.Enabled = false;
                timerOnTab.Interval = Config.Tabs.DragOverTabOpensSDT ? INTERVAL_SHOWMENU : INTERVAL_SELCTTAB;
                timerOnTab.Enabled = true;
                if(Config.Tabs.DragOverTabOpensSDT && (iState != -1)) {
                    Rectangle tabRect = tabControl1.GetTabRect(tab);
                    Point lpPoints = new Point(tabRect.X + ((tabRect.Width * 3) / 4), tabRect.Bottom + 0x10);
                    string[] strArray = QTUtility.TextResourcesDic["DragDropToolTip"];
                    string str;
                    switch((grfKeyState & 12)) {
                        case 4:
                            str = strArray[1];
                            break;

                        case 8:
                            str = strArray[0];
                            break;

                        case 12:
                            str = strArray[2];
                            break;

                        default:
                            if(iState == 1) {
                                str = strArray[0];
                            }
                            else {
                                str = strArray[1];
                            }
                            break;
                    }
                    if(toolTipForDD == null) {
                        toolTipForDD = new ToolTip(components);
                        toolTipForDD.UseAnimation = toolTipForDD.UseFading = false;
                    }
                    toolTipForDD.ToolTipTitle = str;
                    if(PInvoke.GetForegroundWindow() != ExplorerHandle) {
                        Type type = typeof(ToolTip);
                        const BindingFlags bindingAttr = BindingFlags.NonPublic | BindingFlags.Instance;
                        MethodInfo method = type.GetMethod("SetTrackPosition", bindingAttr);
                        MethodInfo info2 = type.GetMethod("SetTool", bindingAttr);
                        PInvoke.MapWindowPoints(tabControl1.Handle, IntPtr.Zero, ref lpPoints, 1);
                        method.Invoke(toolTipForDD, new object[] { lpPoints.X, lpPoints.Y });
                        info2.Invoke(toolTipForDD, new object[] { tabControl1, tab.CurrentPath, 2, lpPoints });
                    }
                    else {
                        toolTipForDD.Active = true;
                        toolTipForDD.Show(tab.CurrentPath, tabControl1, lpPoints);
                    }
                }
            }
        }

        private void subDirTip_MenuItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            QMenuItem clickedItem = (QMenuItem)e.ClickedItem;
            if(clickedItem.Target == MenuTarget.Folder) {
                if(clickedItem.IDLData != null) {
                    using(IDLWrapper wrapper = new IDLWrapper(clickedItem.IDLData)) {
                        ShellBrowser.Navigate(wrapper);
                    }
                    return;
                }
                string targetPath = clickedItem.TargetPath;
                Keys modifierKeys = ModifierKeys;
                bool flag = (subDirTip_Tab != null) && (sender == subDirTip_Tab);
                if((modifierKeys & Keys.Control) == Keys.Control) {
                    using(IDLWrapper wrapper2 = new IDLWrapper(targetPath)) {
                        OpenNewWindow(wrapper2);
                        return;
                    }
                }
                if((modifierKeys & Keys.Shift) == Keys.Shift) {
                    using(IDLWrapper wrapper3 = new IDLWrapper(targetPath)) {
                        OpenNewTab(wrapper3, false, true);
                        return;
                    }
                }
                if((!flag || (ContextMenuedTab == CurrentTab)) && CurrentTab.TabLocked)
                {
                    QTUtility2.log("Clone Tab Button1");
                    CloneTabButton(CurrentTab, targetPath, true, TabIndex());
                    return;
                }
                if(flag && (ContextMenuedTab != CurrentTab)) {
                    if(ContextMenuedTab != null) {
                        if(ContextMenuedTab.TabLocked) {
                            // QTUtility2.log("Clone Tab Button2");
                            var index = TabIndex();
                            // ����߼���һ����ǩ
                            // tabControl1.TabPages.IndexOf(ContextMenuedTab) + 1
                            CloneTabButton(
                                ContextMenuedTab, 
                                targetPath, 
                                true,
                                index
                            );
                            return;
                        }

                        NowTabCloned = targetPath == CurrentAddress;
                        ContextMenuedTab.NavigatedTo(targetPath, null, 1, false);
                        // ContextMenuedTab.NavigatedTo(targetPath, null, -1, false);
                        tabControl1.SelectTab(ContextMenuedTab);
                        QTUtility2.log("NavigatedTo SelectTab");
                    }
                    return;
                }
                using(IDLWrapper wrapper4 = new IDLWrapper(targetPath)) {
                    ShellBrowser.Navigate(wrapper4);
                    QTUtility2.log("ShellBrowser.Navigate");
                    return;
                }
            }
            try {
                Process.Start(new ProcessStartInfo(clickedItem.Path) {
                    WorkingDirectory = Path.GetDirectoryName(clickedItem.Path) ?? "",
                    ErrorDialog = true,
                    ErrorDialogParentHandle = ExplorerHandle
                });
                QTUtility2.log("Process.Start");
                if(Config.Misc.KeepRecentFiles) {
                    StaticReg.ExecutedPathsList.Add(clickedItem.Path);
                    QTUtility2.log("StaticReg.ExecutedPathsList.Add");
                }
            }
            catch {
            }
        }

        // �޸�Ԥ��Ŀ¼��ת����ȷ�ı�ǩλ��
        private int TabIndex()
        {
            var index = 1;
            if (Config.Tabs.NewTabPosition == TabPos.Rightmost)
            {
                index = tabControl1.TabPages.Count;
            }
            else if (Config.Tabs.NewTabPosition == TabPos.Left)
            {
                index = tabControl1.SelectedIndex - 1;
            }
            else if (Config.Tabs.NewTabPosition == TabPos.Right)
            {
                index = tabControl1.SelectedIndex + 1;
            }
            else
            {
                index = 0;
            }

            return index;
        }

        private void subDirTip_MenuItemRightClicked(object sender, ItemRightClickedEventArgs e) {
            QMenuItem clickedItem = e.ClickedItem as QMenuItem;
            if(clickedItem != null) {
                using(IDLWrapper wrapper = new IDLWrapper(clickedItem.Path)) {
                    e.HRESULT = shellContextMenu.Open(wrapper, e.IsKey ? e.Point : MousePosition, ((SubDirTipForm)sender).Handle, false);
                }
            }
        }

        private void subDirTip_MultipleMenuItemsClicked(object sender, EventArgs e) {
            List<string> executedDirectories = ((SubDirTipForm)sender).ExecutedDirectories;
            if((ModifierKeys & Keys.Control) == Keys.Control) {
                QTUtility2.InitializeTemporaryPaths();
                StaticReg.CreateWindowPaths.AddRange(executedDirectories);
                using(IDLWrapper wrapper = new IDLWrapper(executedDirectories[0])) {
                    OpenNewWindow(wrapper);
                    return;
                }
            }
            bool flag = true;
            foreach(string str in executedDirectories) {
                OpenNewTab(str, !flag);
                flag = false;
            }
        }

        private void subDirTip_MultipleMenuItemsRightClicked(object sender, ItemRightClickedEventArgs e) {
            List<string> executedDirectories = ((SubDirTipForm)sender).ExecutedDirectories;
            // TODO: Replace ExecutedDirectories with ExecutedIDLs.
            List<byte[]> executedIDLs = executedDirectories.Select(path => {
                using(IDLWrapper wrapper = new IDLWrapper(path)) {
                    return wrapper.IDL;
                }
            }).ToList();
            e.HRESULT = shellContextMenu.Open(executedIDLs, e.IsKey ? e.Point : MousePosition, ((SubDirTipForm)sender).Handle);
        }

        private void subDirTip_Tab_MenuClosed(object sender, EventArgs e) {
            tabControl1.SetSubDirTipShown(false);
            tabControl1.RefreshFolderImage();
        }

        internal static void SyncTaskBarMenu() {
            // todo
            /*
            using(RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root)) {
                IntPtr hWnd = QTUtility2.ReadRegHandle("TaskBarHandle", key);
                if((hWnd != IntPtr.Zero) && PInvoke.IsWindow(hWnd)) {
                    QTUtility2.SendCOPYDATASTRUCT(hWnd, (IntPtr)3, string.Empty, IntPtr.Zero);
                }
            }*/
        }

        private void SyncToolbarTravelButton()
        {
            if(!QTUtility.IsXP)
            {
                IntPtr ptr = (IntPtr)0x10001;
                IntPtr ptr2 = (IntPtr)0x10000;
                bool flag = (navBtnsFlag & 1) != 0;
                bool flag2 = (navBtnsFlag & 2) != 0;
                PInvoke.SendMessage(TravelToolBarHandle, 0x401, (IntPtr)0x100, flag ? ptr : ptr2);
                PInvoke.SendMessage(TravelToolBarHandle, 0x401, (IntPtr)0x101, flag2 ? ptr : ptr2);
                PInvoke.SendMessage(TravelToolBarHandle, 0x401, (IntPtr)0x102, (flag || flag2) ? ptr : ptr2);
            }
        }

        private void SyncTravelState()
        {
            if(CurrentTab != null)
            {
                navBtnsFlag = ((CurrentTab.HistoryCount_Back > 1) ? 1 : 0) | ((CurrentTab.HistoryCount_Forward > 0) ? 2 : 0);
                QTabItem.CheckSubTexts(tabControl1);
                SyncToolbarTravelButton();
            }
        }

        private void tabControl1_CloseButtonClicked(object sender, QTabCancelEventArgs e) {
            if(NowTabDragging) {
                Cursor = Cursors.Default;
                NowTabDragging = false;
                DraggingTab = null;
                DraggingDestRect = Rectangle.Empty;
                e.Cancel = true;
            }
            else if(!Explorer.Busy) {
                if(tabControl1.TabCount > 1) {
                    e.Cancel = !CloseTab(e.TabPage);
                }
                else {
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegConst.Root))
                    {
                        string[] list = (from QTabItem item2 in tabControl1.TabPages
                                         where item2.TabLocked
                                         select item2.CurrentPath).ToArray();

                        // MessageBox.Show(String.Join(",", list));
                        QTUtility2.WriteRegBinary(list, "TabsLocked", key);
                    }
                    WindowUtils.CloseExplorer(ExplorerHandle, 1);
                }
            }
        }



        private void tabControl1_ItemDrag(object sender, ItemDragEventArgs e) {
            QTabItem item = (QTabItem)e.Item;
            string currentPath = item.CurrentPath;
            if(Directory.Exists(currentPath)) {
                ShellMethods.DoDragDrop(currentPath, this);
            }
        }

        /**
         * bug ��ֻ��һ����ǩ��ʱ�򣬵����ǩ�հ״�ʶ��Ϊ��ǩ
         */
        private void tabControl1_MouseDoubleClick(object sender, MouseEventArgs e) {
            if((ModifierKeys != Keys.Control) && (e.Button == MouseButtons.Left)) {
                QTabItem tabMouseOn = tabControl1.GetTabMouseOn();
                if(tabMouseOn != null) {
                    MouseChord chord = QTUtility.MakeMouseChord(MouseChord.Double, ModifierKeys);
                    BindAction action;
                    if(Config.Mouse.TabActions.TryGetValue(chord, out action)) {
                        QTUtility2.log("QTTabBarClass tabControl1_MouseDoubleClick " + action);
                        DoBindAction(action, false, DraggingTab);
                    }
                }
                else {
                    OnMouseDoubleClick(e);
                }
            }
        }

        private void tabControl1_MouseDown(object sender, MouseEventArgs e) {
            QTabItem tabMouseOn = tabControl1.GetTabMouseOn();
            DraggingTab = null;
            if(tabMouseOn != null) {
                if(e.Button == MouseButtons.Left) {
                    NowTabDragging = true;
                    DraggingTab = tabMouseOn;
                }
                else if(e.Button == MouseButtons.Right) {
                    ContextMenuedTab = tabMouseOn;
                }
            }
        }

        private void tabControl1_MouseEnter(object sender, EventArgs e) {
        }

        private void tabControl1_MouseLeave(object sender, EventArgs e) {
        }

        private void tabControl1_MouseMove(object sender, MouseEventArgs e) {
            RECT rect;
            if((tabControl1.Capture && (((e.X < 0) || (e.Y < 0)) || ((e.X > tabControl1.Width) || (e.Y > tabControl1.Height)))) && (PInvoke.GetWindowRect(ReBarHandle, out rect) && !PInvoke.PtInRect(ref rect, tabControl1.PointToScreen(e.Location)))) {
                Cursor = Cursors.Default;
                tabControl1.Capture = false;
            }
            else if((NowTabDragging && (DraggingTab != null)) && ((ModifierKeys & Keys.Shift) != Keys.Shift)) {
                if(Explorer.Busy || (MouseButtons != MouseButtons.Left)) {
                    NowTabDragging = false;
                    // Leave DraggingTab set so MouseUp doesn't get confused.
                    // It will be unset in MouseUp.
                }
                else {
                    int num;
                    QTabItem tabMouseOn = tabControl1.GetTabMouseOn(out num);
                    int index = tabControl1.TabPages.IndexOf(DraggingTab);
                    if((num > (tabControl1.TabCount - 1)) || (num < 0)) {
                        if((num == -1) && (ModifierKeys == Keys.Control)) {
                            Cursor = GetCursor(false);
                            DraggingDestRect = new Rectangle(1, 0, 0, 0);
                        }
                        else {
                            Cursor = Cursors.Default;
                        }
                    }
                    else if((index <= (tabControl1.TabCount - 1)) && (index >= 0)) {
                        Rectangle tabRect = tabControl1.GetTabRect(num, false);
                        Rectangle rectangle2 = tabControl1.GetTabRect(index, false);
                        if(tabMouseOn != null) {
                            if(tabMouseOn != DraggingTab) {
                                if(!DraggingDestRect.Contains(tabControl1.PointToClient(MousePosition))) {
                                    Cursor = GetCursor(true);
                                    bool flag = tabMouseOn.Row != DraggingTab.Row;
                                    bool flag2 = tabControl1.SelectedTab != DraggingTab;
                                    tabControl1.TabPages.Relocate(index, num);
                                    if(num < index) {
                                        DraggingDestRect = new Rectangle(tabRect.X + rectangle2.Width, tabRect.Y, tabRect.Width - rectangle2.Width, tabRect.Height);
                                    }
                                    else {
                                        DraggingDestRect = new Rectangle(tabRect.X, tabRect.Y, tabRect.Width - rectangle2.Width, tabRect.Height);
                                    }
                                    if((flag && !flag2) && !Config.Tabs.MultipleTabRows) {
                                        Rectangle rectangle3 = tabControl1.GetTabRect(num, false);
                                        Point p = new Point(rectangle3.X + (rectangle3.Width / 2), rectangle3.Y + (Config.Skin.TabHeight / 2));
                                        Cursor.Position = tabControl1.PointToScreen(p);
                                    }
                                }
                            }
                            else if((curTabCloning != null) && (Cursor == curTabCloning)) {
                                Cursor = GetCursor(true);
                            }
                        }
                    }
                }
            }
        }

        // ����ڱ�ǩ�ϲ���
        private void tabControl1_MouseUp(object sender, MouseEventArgs e) {
            if (null == tabControl1 || tabControl1.IsDisposed)
            {
                // ��������һ����ǩ�������bug
                return;
            }
            QTabItem tabMouseOn = tabControl1.GetTabMouseOn();
            if(NowTabDragging && e.Button == MouseButtons.Left) {
                Keys modifierKeys = ModifierKeys;
                if(tabMouseOn == null) {
                    if(DraggingTab != null && (modifierKeys == Keys.Control || modifierKeys == (Keys.Control | Keys.Shift))) {
                        bool cloning = false;
                        Point pt = tabControl1.PointToScreen(e.Location);
                        if(!QTUtility.IsXP) {
                            RECT rect;
                            PInvoke.GetWindowRect(ReBarHandle, out rect);
                            cloning = PInvoke.PtInRect(ref rect, pt);
                        }
                        else {
                            RECT rect2;
                            IntPtr ptr;
                            PInvoke.GetWindowRect(Handle, out rect2);
                            if(PInvoke.PtInRect(ref rect2, pt)) {
                                cloning = true;
                            }
                        }
                        if(cloning) {
                            CloneTabButton(DraggingTab, null, false, tabControl1.TabCount);
                        }
                    }
                } 
                else if(tabMouseOn == DraggingTab && DraggingDestRect == Rectangle.Empty) {
                    MouseChord chord = QTUtility.MakeMouseChord(MouseChord.Left, ModifierKeys);
                    BindAction action;
                    if(Config.Mouse.TabActions.TryGetValue(chord, out action)) {
                        QTUtility2.log("QTTabBarClass DraggingTab " + action);
                        DoBindAction(action, false, DraggingTab);
                    }
                }
                NowTabDragging = false;
                DraggingTab = null;
                DraggingDestRect = Rectangle.Empty;
            }
            else if(e.Button == MouseButtons.Middle && !Explorer.Busy && tabMouseOn != null) {
                DraggingTab = null;
                NowTabDragging = false;
                MouseChord chord = QTUtility.MakeMouseChord(MouseChord.Middle, ModifierKeys);
                BindAction action;
                if(Config.Mouse.TabActions.TryGetValue(chord, out action)) {
                    QTUtility2.log("QTTabBarClass MouseButtons.Middle " + action);
                    DoBindAction(action, false, tabMouseOn);
                }
            }
            else if(tabMouseOn == null) {
                NowTabDragging = false;
                if(DraggingTab == null) OnMouseUp(e); // This will prevent the bar's MouseUp from 
                DraggingTab = null;                   // firing if the MouseDown was on a tab.
            }
            Cursor = Cursors.Default;
        }

        private void tabControl1_PointedTabChanged(object sender, QTabCancelEventArgs e) {
        }


       



        private void tabControl1_TabCountChanged(object sender, QTabCancelEventArgs e) {
        }

        private void tabControl1_TabIconMouseDown(object sender, QTabCancelEventArgs e) {
            ShowSubdirTip_Tab(e.TabPage, e.Action == TabControlAction.Selecting, e.TabPageIndex, false, e.Cancel);
        }

       

        private void tabSwitcher_Switched(object sender, ItemCheckEventArgs e) {
            tabControl1.SelectedIndex = e.Index;
        }

        private void timerOnTab_Tick(object sender, EventArgs e) {
            timerOnTab.Enabled = false;
            QTabItem tabMouseOn = tabControl1.GetTabMouseOn();
            if(((tabMouseOn != null) && (tabMouseOn == tabForDD)) && tabControl1.TabPages.Contains(tabMouseOn)) {
                if(Config.Tabs.DragOverTabOpensSDT) {
                    WindowUtils.BringExplorerToFront(ExplorerHandle);
                    ShowSubdirTip_Tab(tabMouseOn, true, tabControl1.TabOffset, false, fToggleTabMenu);
                    fToggleTabMenu = !fToggleTabMenu;
                    timerOnTab.Enabled = true;
                    if(toolTipForDD != null) {
                        toolTipForDD.Active = false;
                    }
                }
                else {
                    tabControl1.SelectTab(tabMouseOn);
                }
            }
        }

        private void timerSelectionChanged_Tick(object sender, EventArgs e) {
            try {
                timerSelectionChanged.Enabled = false;
            }
            catch( Exception e1 ) {
                QTUtility2.MakeErrorLog( e1, "QTTabBarClass timerSelectionChanged_Tick");
            }
        }

        // ���ô����ö�����
        private void ToggleTopMost() {
            QTUtility2.log("QTTabBarClass ToggleTopMost");
            if(PInvoke.Ptr_OP_AND(PInvoke.GetWindowLongPtr(ExplorerHandle, -20), 8) != IntPtr.Zero) {
                PInvoke.SetWindowPos(ExplorerHandle, (IntPtr)(-2), 0, 0, 0, 0, 3);
                NowTopMost = false;
            }
            else {
                PInvoke.SetWindowPos(ExplorerHandle, (IntPtr)(-1), 0, 0, 0, 0, 3);
                NowTopMost = true;
            }
        }

        public override int TranslateAcceleratorIO(ref MSG msg) {
            if(msg.message == WM.KEYDOWN) {
                Keys wParam = (Keys)((int)((long)msg.wParam));
                bool flag = (((int)((long)msg.lParam)) & 0x40000000) != 0;
                switch(wParam) {
                    case Keys.Delete: {
                            if(!tabControl1.Focused || ((subDirTip_Tab != null) && subDirTip_Tab.MenuIsShowing)) {
                                break;
                            }
                            int focusedTabIndex = tabControl1.GetFocusedTabIndex();
                            if((-1 < focusedTabIndex) && (focusedTabIndex < tabControl1.TabCount)) {
                                bool flag3 = focusedTabIndex == (tabControl1.TabCount - 1);
                                if(CloseTab(tabControl1.TabPages[focusedTabIndex]) && flag3) {
                                    tabControl1.FocusNextTab(true, false, false);
                                }
                            }
                            return 0;
                        }

                    case Keys.F6:
                    case Keys.Tab:
                    case Keys.Left:
                    case Keys.Right: {
                            if(!tabControl1.Focused || ((subDirTip_Tab != null) && subDirTip_Tab.MenuIsShowing)) {
                                break;
                            }
                            bool fBack = (ModifierKeys == Keys.Shift) || (wParam == Keys.Left);
                            if(!tabControl1.FocusNextTab(fBack, false, false)) {
                                break;
                            }
                            return 0;
                        }
                    case Keys.Back:
                        return 0;

                    case Keys.Return:
                    case Keys.Space:
                        if(!flag && !tabControl1.SelectFocusedTab()) {
                            break;
                        }
                        listView.SetFocus();
                        return 0;

                    case Keys.Escape:
                        if(tabControl1.Focused && ((subDirTip_Tab == null) || !subDirTip_Tab.MenuIsShowing)) {
                            listView.SetFocus();
                        }
                        break;

                    case Keys.End:
                    case Keys.Home:
                        if((!tabControl1.Focused || ((subDirTip_Tab != null) && subDirTip_Tab.MenuIsShowing)) || !tabControl1.FocusNextTab(wParam == Keys.Home, false, true)) {
                            break;
                        }
                        return 0;

                    case Keys.Up:
                    case Keys.Down:
                        if(((!Config.Tabs.ShowSubDirTipOnTab || !tabControl1.Focused) || ((subDirTip_Tab != null) && subDirTip_Tab.MenuIsShowing)) || (!flag && !tabControl1.PerformFocusedFolderIconClick(wParam == Keys.Up))) {
                            break;
                        }
                        return 0;
                }
            }
            return base.TranslateAcceleratorIO(ref msg);
        }

        private bool travelBtnController_MessageCaptured(ref Message m)
        {
            if(CurrentTab == null)
            {
                QTUtility2.log("QTTabBarClass travelBtnController_MessageCaptured CurrentTab == null");
                return false;
            }
            switch(m.Msg)
            {
                case WM.LBUTTONDOWN:
                case WM.LBUTTONUP:
                    {
                        Point pt = QTUtility2.PointFromLPARAM(m.LParam);
                        int num = (int)PInvoke.SendMessage(travelBtnController.Handle, 0x445, IntPtr.Zero, ref pt);
                        bool flag = CurrentTab.HistoryCount_Back > 1;
                        bool flag2 = CurrentTab.HistoryCount_Forward > 0;
                        if(m.Msg != 0x202)
                        {
                            PInvoke.SetCapture(travelBtnController.Handle);
                            if(((flag && (num == 0)) || (flag2 && (num == 1))) || ((flag || flag2) && (num == 2)))
                            {
                                int num5 = (int)PInvoke.SendMessage(travelBtnController.Handle, 0x412, (IntPtr)(0x100 + num), IntPtr.Zero);
                                int num6 = num5 | 2;
                                PInvoke.SendMessage(travelBtnController.Handle, 0x411, (IntPtr)(0x100 + num), (IntPtr)num6);
                            }
                            if((num == 2) && (flag || flag2))
                            {
                                RECT rect;
                                IntPtr hWnd = PInvoke.SendMessage(travelBtnController.Handle, 0x423, IntPtr.Zero, IntPtr.Zero);
                                if(hWnd != IntPtr.Zero)
                                {
                                    PInvoke.SendMessage(hWnd, 0x41c, IntPtr.Zero, IntPtr.Zero);
                                }
                                PInvoke.GetWindowRect(travelBtnController.Handle, out rect);
                                NavigationButtons_DropDownOpening(buttonNavHistoryMenu, new EventArgs());
                                buttonNavHistoryMenu.DropDown.Show(new Point(rect.left - 2, rect.bottom + 1));
                            }
                            break;
                        }
                        PInvoke.ReleaseCapture();
                        for(int i = 0; i < 3; i++)
                        {
                            int num3 = (int)PInvoke.SendMessage(travelBtnController.Handle, 0x412, (IntPtr)(0x100 + i), IntPtr.Zero);
                            int num4 = num3 & -3;
                            PInvoke.SendMessage(travelBtnController.Handle, 0x411, (IntPtr)(0x100 + i), (IntPtr)num4);
                        }
                        if((num == 0) && flag)
                        {
                            NavigateCurrentTab(true);
                        }
                        else if((num == 1) && flag2)
                        {
                            NavigateCurrentTab(false);
                        }
                        break;
                    }
                case WM.LBUTTONDBLCLK:
                    m.Result = IntPtr.Zero;
                    return true;

                case WM.USER + 1:
                    if(((((int)((long)m.LParam)) >> 0x10) & 0xffff) == 1)
                    {
                        return false;
                    }
                    m.Result = (IntPtr)1;
                    return true;

                case WM.MOUSEACTIVATE:
                    if(buttonNavHistoryMenu.DropDown.Visible)
                    {
                        m.Result = (IntPtr)4;
                        buttonNavHistoryMenu.DropDown.Close(ToolStripDropDownCloseReason.AppClicked);
                        return true;
                    }
                    return false;

                case WM.NOTIFY:
                    {
                        NMHDR nmhdr = (NMHDR)Marshal.PtrToStructure(m.LParam, typeof(NMHDR));
                        if(nmhdr.code != -530)
                        {
                            return false;
                        }
                        NMTTDISPINFO nmttdispinfo = (NMTTDISPINFO)Marshal.PtrToStructure(m.LParam, typeof(NMTTDISPINFO));
                        string str;
                        if(nmttdispinfo.hdr.idFrom == ((IntPtr)0x100))
                        {
                            str = MakeTravelBtnTooltipText(true);
                            if(str.Length > 0x4f)
                            {
                                str = "Back";
                            }
                        }
                        else if(nmttdispinfo.hdr.idFrom == ((IntPtr)0x101))
                        {
                            str = MakeTravelBtnTooltipText(false);
                            if(str.Length > 0x4f)
                            {
                                str = "Forward";
                            }
                        }
                        else
                        {
                            return false;
                        }
                        nmttdispinfo.szText = str;
                        Marshal.StructureToPtr(nmttdispinfo, m.LParam, false);
                        m.Result = IntPtr.Zero;
                        return true;
                    }
                default:
                    return false;
            }
            m.Result = IntPtr.Zero;
            return true;
        }


        private bool FolderLinkClicked(IDLWrapper wrapper, Keys modifierKeys, bool middle) {
            QTUtility2.log("QTTabBarClass FolderLinkClicked");
            MouseChord chord = QTUtility.MakeMouseChord(middle ? MouseChord.Middle : MouseChord.Left, modifierKeys);
            BindAction action;
            if(Config.Mouse.LinkActions.TryGetValue(chord, out action)) {
                DoBindAction(action, false, null, wrapper);
                return true;
            }
            else {
                QTUtility2.log("QTTabBarClass FolderLinkClicked δ��ȡ�����õĶ���");
                return false;
            }
        }

        internal bool TryGetSelection(out Address[] adSelectedItems, out string pathFocused, bool fDisplayName) {
            return ShellBrowser.TryGetSelection(out adSelectedItems, out pathFocused, fDisplayName);
        }

       

        private void tsmiBranchRoot_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            QTUtility2.log("QTTabBarClass tsmiBranchRoot_DropDownItemClicked");
            QTabItem tag = (QTabItem)((ToolStripMenuItem)sender).Tag;
            if(tag != null) {
                NavigateBranches(tag, ((QMenuItem)e.ClickedItem).MenuItemArguments.Index);
            }
        }

        public override void UIActivateIO(int fActivate, ref MSG Msg) {
            QTUtility2.log("QTTabBarClass UIActivateIO");
            if(fActivate != 0) {
                tabControl1.Focus();
                tabControl1.FocusNextTab(ModifierKeys == Keys.Shift, true, false);
            }
        }

        [ComUnregisterFunction]
        private static void Unregister(Type t) {
            QTUtility2.log("QTTabBarClass Unregister");
            string name = t.GUID.ToString("B");
            try {
                using(RegistryKey key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Internet Explorer\Toolbar")) {
                    key.DeleteValue(name, false);
                }
            }
            catch {
            }
            try {
                using(RegistryKey key2 = Registry.ClassesRoot.OpenSubKey("CLSID", true)) {
                    try {
                        key2.DeleteSubKeyTree(name);
                    }
                    catch {
                    }
                    try {
                        key2.DeleteSubKeyTree("{D2BF470E-ED1C-487F-A444-2BD8835EB6CE}");
                    }
                    catch {
                    }
                }
            }
            catch {
            }
            
            return;
            // TODO: Make the following code optional in the Uninstaller.
#if false
            try {
                using(RegistryKey key3 = Registry.Users) {
                    try {
                        foreach(string str2 in key3.GetSubKeyNames()) {
                            bool flag = true;
                            try {
                                using(RegistryKey key4 = key3.OpenSubKey(str2 + @"\Software\Quizo", true)) {
                                    if(key4 != null) {
                                        try {
                                            key4.DeleteSubKeyTree("QTTabBar");
                                            string[] subKeyNames = key4.GetSubKeyNames();
                                            flag = (subKeyNames != null) && (subKeyNames.Length > 0);
                                        }
                                        catch {
                                        }
                                    }
                                }
                            }
                            catch {
                            }
                            try {
                                if(!flag) {
                                    using(RegistryKey key5 = key3.OpenSubKey(str2 + @"\Software", true)) {
                                        if(key5 != null) {
                                            key5.DeleteSubKeyTree("Quizo");
                                        }
                                    }
                                }
                            }
                            catch {
                            }
                        }
                    }
                    catch {
                    }
                }
            }
            catch {
            }
#endif
        }

        private void UpOneLevel()
        {
            // ������һ��Ŀ¼
            if(CurrentTab.TabLocked) {
                QTabItem tab = CurrentTab.Clone();
                AddInsertTab(tab);
                tabControl1.SelectTab(tab);
            }
            if(!QTUtility.IsXP) {
                PInvoke.SendMessage(WindowUtils.GetShellTabWindowClass(ExplorerHandle), 0x111, (IntPtr)0xa022, IntPtr.Zero);
            }
            else {
                PInvoke.SendMessage(ExplorerHandle, 0x111, (IntPtr)0xa022, IntPtr.Zero);
            }
        }

        internal static void WaitTimeout(int msec) {
            Thread.Sleep(msec);
        }
        /**
         * ��Ϣ��� by indiff
         */
        protected override void WndProc(ref Message m) {
            try {
                switch(m.Msg) {
                    case WM.APP + 1: // todo: what sends this?
                        NowModalDialogShown = m.WParam != IntPtr.Zero;
                        return;

                    case WM.DROPFILES:  // �϶��ļ�
                        HandleFileDrop(m.WParam);
                        break;

                    case WM.DRAWITEM:
                    case WM.MEASUREITEM:
                    case WM.INITMENUPOPUP:
                        if(m.HWnd == Handle && shellContextMenu.TryHandleMenuMsg(m.Msg, m.WParam, m.LParam)) {
                            return;
                        }
                        break;
                }
                base.WndProc(ref m);
            }
            catch(Exception ex) {
                QTUtility2.MakeErrorLog(ex, String.Format("Message: {0:x4}", m.Msg));
            }
        }

        // todo: This seems like it should go after every new tab creation, no?
        private void RestoreWindow()
        {
            bool fIsIconic = PInvoke.IsIconic(ExplorerHandle);
            InstanceManager.RemoveFromTrayIcon(Handle);
            WindowUtils.BringExplorerToFront(ExplorerHandle);
            if(fIsIconic) {
                foreach(QTabItem item2 in tabControl1.TabPages) {
                    item2.RefreshRectangle();
                }
                tabControl1.Refresh();
            }
        }

        protected override void OnDpiChanged(int oldDpi, int newDpi)
        {
            QTUtility2.log("QTTabBarClass OnDpiChanged");
            int iType = 0;
            if (Config.Tabs.MultipleTabRows)
            {
                iType = Config.Tabs.ActiveTabOnBottomRow ? 1 : 2;
            }
            // this.SetBarRows(this.tabControl.RowCount);
            SetBarRows(tabControl1.SetTabRowType(iType));
        }



        #region ��ǩ���¼���
        public RebarController rebarController;
        protected string CurrentAddress;
        protected QTabItem CurrentTab;
        protected int BandHeight;
        public static int BandHeightSpace = 3;
        protected ShellBrowserEx ShellBrowser;
        
        protected List<QTabItem> lstActivatedTabs = new List<QTabItem>(0x10);
        protected IntPtr ExplorerHandle;
        protected Dictionary<int, ITravelLogEntry> LogEntryDic = new Dictionary<int, ITravelLogEntry>();
        protected AbstractListView listView = new AbstractListView();
        protected ListViewMonitor listViewManager;

        protected List<ToolStripItem> lstPluginMenuItems_Sys;
        protected List<ToolStripItem> lstPluginMenuItems_Tab;
        protected ITravelLogStg TravelLog;

        public bool HideExplorer
        {
            get
            {
                return fHideExplorer;
            }
        }

        protected override bool IsTabSubFolderMenuVisible => throw new NotImplementedException();
        private int navBtnsFlag;
        protected bool NavigatedByCode;
        protected bool NowTabsAddingRemoving;
        protected bool NowInTravelLog;
        protected bool NowModalDialogShown;
        protected bool NowOpenedByGroupOpener;
        protected bool NowTabCloned;
        protected bool NowTabCreated;
        protected bool NowTabDragging;
        protected bool NowTopMost;
        protected bool fNavigatedByTabSelection;
        protected int CurrentTravelLogIndex;
        // TODO: Options is simple so this should be cleaned up a lot
        protected ToolStripClasses optionsStrip;
        protected ToolStripButton buttonOptions;

        protected void AddToHistory(QTabItem closingTab)
        {
            string currentPath = closingTab.CurrentPath;
            if ((Config.Misc.KeepHistory && !string.IsNullOrEmpty(currentPath)) && !IsSearchResultFolder(currentPath))
            {
                if (QTUtility2.IsShellPathButNotFileSystem(currentPath) && (currentPath.IndexOf("???") == -1))
                {
                    currentPath = currentPath + "???" + closingTab.GetLogHash(true, 0);
                }
                StaticReg.ClosedTabHistoryList.Add(currentPath);
            }
        }

        // �رձ�ǩ�� ��������򲻹ر�
        protected bool CloseTab(QTabItem closingTab, bool fCritical, bool fSkipSync = false)
        {
            if (closingTab == null)
            {
                return false;
            }
            if ((!fCritical && closingTab.TabLocked) && QTUtility2.PathExists(closingTab.CurrentPath))
            {
                return false;
            }
            int index = tabControl1.TabPages.IndexOf(closingTab);
            if (index == -1)
            {
                return false;
            }
            lstActivatedTabs.Remove(closingTab);
            AddToHistory(closingTab);
            tabControl1.TabPages.Remove(closingTab);
            closingTab.OnClose();
            if (closingTab != CurrentTab)
            {
                if (!fSkipSync)
                {
                    QTabItem.CheckSubTexts(tabControl1);
                }
                return true;
            }
            CurrentTab = null;
            int tabCount = tabControl1.TabCount;
            if (tabCount == 0) return true;
            QTabItem tabPage = null;
            switch (Config.Tabs.NextAfterClosed)
            {
                case TabPos.Right:
                    tabPage = tabControl1.TabPages[index == tabCount ? index - 1 : index];
                    break;

                case TabPos.Left:
                    tabPage = tabControl1.TabPages[index == 0 ? 0 : index - 1];
                    break;

                case TabPos.Rightmost:
                    tabPage = tabControl1.TabPages[tabCount - 1];
                    break;

                case TabPos.Leftmost:
                    tabPage = tabControl1.TabPages[0];
                    break;

                case TabPos.LastActive:
                    if (lstActivatedTabs.Count > 0)
                    {
                        QTabItem lastTab = lstActivatedTabs[lstActivatedTabs.Count - 1];
                        lstActivatedTabs.RemoveAt(lstActivatedTabs.Count - 1);
                        tabPage = tabControl1.TabPages.Contains(lastTab)
                                ? lastTab
                                : tabControl1.TabPages[0];
                    }
                    else
                    {
                        tabPage = tabControl1.TabPages[0];
                    }
                    break;
            }
            if (tabPage != null)
            {
                tabControl1.SelectTab(tabPage);
            }
            else
            {
                tabControl1.SelectTab(0);
            }
            return true;
        }

        protected void CloseTabs(IEnumerable<QTabItem> tabs, bool fCritical = false)
        {
            tabControl1.SetRedraw(false);
            bool closeCurrent = false;
            foreach (QTabItem tab in tabs)
            {
                if (tab == CurrentTab)
                    closeCurrent = true;
                else
                    CloseTab(tab, fCritical, true);
            }
            if (closeCurrent)
            {
                CloseTab(CurrentTab, fCritical);
            }
            else
            {
                QTabItem.CheckSubTexts(tabControl1);
            }
            if (tabControl1.TabCount > 0)
            {
                tabControl1.SetRedraw(true);
            }
        }

        // TODO: Optional params
        protected bool CloseTab(QTabItem closingTab)
        {
            return ((tabControl1.TabCount > 1) && CloseTab(closingTab, false));
        }

        protected void ShowMessageNavCanceled(string failedPath, bool fModal)
        {
            QTUtility2.log("QTTabBarClass ShowMessageNavCanceled: " + failedPath);
            QTUtility2.MakeErrorLog(null, string.Format("Failed navigation: {0}", failedPath));
            MessageForm.Show(ExplorerHandle,
                string.Format(QTUtility.TextResourcesDic["TabBar_Message"][0], failedPath),
                string.Empty,
                MessageBoxIcon.Asterisk,
                0x2710,
                fModal);
        }

        protected void CancelFailedTabChanging(string newPath)
        {
            if (!CloseTab(tabControl1.SelectedTab, true))
            {
                if (tabControl1.TabCount == 1)
                {
                    WindowUtils.CloseExplorer(ExplorerHandle, 2);
                }
                else
                {
                    ShowMessageNavCanceled(newPath, false);
                    if (CurrentTab == null)
                    {
                        tabControl1.SelectedIndex = 0;
                    }
                }
            }
            else
            {
                StaticReg.ClosedTabHistoryList.Remove(newPath);
                if (tabControl1.TabCount == 0)
                {
                    ShowMessageNavCanceled(newPath, true);
                    WindowUtils.CloseExplorer(ExplorerHandle, 2);
                }
                else
                {
                    if (CurrentTab == null)
                    {
                        tabControl1.SelectedIndex = 0;
                    }
                    else
                    {
                        tabControl1.SelectTab(CurrentTab);
                    }
                    ShowMessageNavCanceled(newPath, false);
                }
            }
        }

        protected bool NavigateToPastSpecialDir(int hash)
        {
            IEnumTravelLogEntry ppenum = null;
            try
            {
                ITravelLogEntry entry2;
                if (TravelLog.EnumEntries(0x31, out ppenum) != 0)
                {
                    goto Label_007C;
                }
            Label_0013:
                do
                {
                    if (ppenum.Next(1, out entry2, 0) != 0)
                    {
                        goto Label_007C;
                    }
                    if (entry2 != LogEntryDic[hash])
                    {
                        goto Label_0057;
                    }
                }
                while (TravelLog.TravelTo(entry2) != 0);
                NowInTravelLog = true;
                CurrentTravelLogIndex++;
                return true;
            Label_0057:
                if (entry2 != null)
                {
                    Marshal.ReleaseComObject(entry2);
                }
                goto Label_0013;
            }
            catch (Exception exception)
            {
                QTUtility2.MakeErrorLog(exception);
            }
            finally
            {
                if (ppenum != null)
                {
                    QTUtility2.log("ReleaseComObject ppenum");
                    Marshal.ReleaseComObject(ppenum);
                }
            }
        Label_007C:
            return false;
        }

        /**
        * TODO config to refresh  when tab control selected index changed
        * ���л���ǩ��ʱ�������Ƿ����ˢ��
        * �����쳣���
        * System.NullReferenceException: δ�������������õ������ʵ����
          �� QTTabBarLib.Interop.IShellBrowser.BrowseObject(IntPtr pidl, SBSP wFlags)
          �� QTTabBarLib.ShellBrowserEx.Navigate(IDLWrapper idlw, SBSP flags)
          �� QTTabBarLib.QTTabBarClass.tabControl1_SelectedIndexChanged(Object sender, EventArgs e)
        */
        protected void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            QTUtility2.log("tabControl1_SelectedIndexChanged" );
            QTabItem selectedTab = tabControl1.SelectedTab;
            string currentPath = selectedTab.CurrentPath;
            if (IsSpecialFolderNeedsToTravel(currentPath) &&
               LogEntryDic.ContainsKey(selectedTab.GetLogHash(true, 0)))
            {
                NavigatedByCode = true;
                CurrentTab = selectedTab;
                while (lstActivatedTabs.Remove(CurrentTab))
                {
                }
                lstActivatedTabs.Add(CurrentTab);
                if (lstActivatedTabs.Count > 15)
                {
                    lstActivatedTabs.RemoveAt(0);
                }
                fNavigatedByTabSelection = NavigateToPastSpecialDir(CurrentTab.GetLogHash(true, 0));
                if (tabControl1.Focused)
                {
                    listView.SetFocus();
                }
            }
            else
            {
                IDLWrapper idlw = null;
                if ((selectedTab.CurrentIDL != null) && (selectedTab.CurrentIDL.Length > 0))
                {
                    idlw = new IDLWrapper(selectedTab.CurrentIDL);
                }
                if ((idlw == null) || !idlw.Available)
                {
                    idlw = new IDLWrapper(selectedTab.CurrentPath);
                }
                using (idlw)
                {
                    if (!idlw.Available)
                    {
                        CancelFailedTabChanging(currentPath);
                        return;
                    }
                    CurrentTab = selectedTab;
                    while (lstActivatedTabs.Remove(CurrentTab))
                    {
                    }
                    lstActivatedTabs.Add(CurrentTab);
                    if (lstActivatedTabs.Count > 15)
                    {
                        lstActivatedTabs.RemoveAt(0);
                    }
                    if (((currentPath != CurrentAddress) ||
                        (QTUtility.IsXP && (currentPath == QTUtility.PATH_SEARCHFOLDER))) ||
                       NowTabCloned)
                    {
                        NavigatedByCode = true;
                        fNavigatedByTabSelection = true;
                        NowTabCloned = false;
                        if (ShellBrowser.Navigate(idlw) != 0)
                        {
                            CancelFailedTabChanging(currentPath);
                            return;
                        }
                    }
                    else
                    {
                        SyncTravelState();
                    }
                }
                if (tabControl1.Focused)
                {
                    listView.SetFocus();
                }
            }
        }

        protected bool IsSpecialFolderNeedsToTravel(string path)
        {
            int index = path.IndexOf("*?*?*");
            if (index != -1)
            {
                path = path.Substring(0, index);
            }
            if (!IsSearchResultFolder(path))
            {
                if (path.PathEquals("::{13E7F612-F261-4391-BEA2-39DF4F3FA311}"))
                {
                    return true;
                }
                if (!path.PathStartsWith(QTUtility.ResMisc[0]) && (!path.EndsWith(QTUtility.ResMisc[0], StringComparison.OrdinalIgnoreCase) || Path.IsPathRooted(path)))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool IsSearchResultFolder(string path)
        {
            return path.PathStartsWith(QTUtility.IsXP ? QTUtility.ResMisc[2] : QTUtility.PATH_SEARCHFOLDER);
        }

        protected void tabControl1_RowCountChanged(object sender, QEventArgs e)
        {
            SetBarRows(e.RowCount);
        }


        protected void SetBarRows(int count)
        {
            // QTUtility2.log("QTTabBarClass SetBarRows");
            // BandHeight = (count * (Config.Skin.TabHeight - 3 )) ;
            // BandHeight = (count * (Config.Skin.TabHeight - 3));
            // BandHeight = (count * (Config.Skin.TabHeight + BandHeightSpace));
            BandHeight = (count * (Config.Skin.TabHeight) + BandHeightSpace);
            // BandHeight = (count * (Config.Skin.TabHeight + 10 )) ;
            // fix bug
            /**
           �쳣�ı�
System.NullReferenceException: δ�������������õ������ʵ����
   �� QTTabBarLib.QTTabBarClass.SetBarRows(Int32 count)
   �� QTTabBarLib.QTabControl.CalculateItemRectangle_MultiRows()
   �� QTTabBarLib.QTabControl.OnPaint_MultipleRow(PaintEventArgs e)
   �� QTTabBarLib.QTabControl.OnPaint(PaintEventArgs e)
   �� System.Windows.Forms.Control.PaintWithErrorHandling(PaintEventArgs e, Int16 layer)
   �� System.Windows.Forms.Control.WmPaint(Message& m)
   �� System.Windows.Forms.Control.WndProc(Message& m)
   �� QTTabBarLib.QTabControl.WndProc(Message& m)
   �� System.Windows.Forms.Control.ControlNativeWindow.WndProc(Message& m)
   �� System.Windows.Forms.NativeWindow.Callback(IntPtr hWnd, Int32 msg, IntPtr wparam, IntPtr lparam)
             **/
            if (null != rebarController)
            {
                rebarController.RefreshHeight();
            }
        }

        protected void tabControl1_Deselecting(object sender, QTabCancelEventArgs e)
        {
            if (e.TabPageIndex != -1)
            {
                SaveSelectedItems(e.TabPage);
            }
        }

        /**
         * ����ѡ����
         */
        protected void SaveSelectedItems(QTabItem tab)
        {
            Address[] addressArray;
            string str;
            if (
                ((tab != null) && !string.IsNullOrEmpty(CurrentAddress)) &&
                ShellBrowser.TryGetSelection(out addressArray, out str, false, ShellBrowser))
            {
                if (addressArray != null && addressArray.Length > 0)
                {
                    QTUtility2.log("SaveSelectedItems addressArray " + addressArray[0].Path);
                }
                tab.SetSelectedItemsAt(CurrentAddress, addressArray, str);
            }
        }

        
        protected void tabControl1_Selecting(object sender, QTabCancelEventArgs e)
        {
            if (NowTabsAddingRemoving)
            {
                QTUtility2.log("tabControl1_Selecting");
                e.Cancel = true;
            }
        }


        // ����+�Ű�ť������±�ǩ�¼�
        private void tabControl1_PlusButtonClicked(object sender, QTabCancelEventArgs e)
        {
            // �±�ǩ��ť qwop
            string clipPath = QTUtility2.GetStringClipboard();
            if (String.IsNullOrEmpty(clipPath))
            {
                // �����ȡ����������·�������Ĭ��
                openDefault();
                return;
            }
            clipPath = clipPath.Trim().Trim(new char[] { ' ', '"' });
            string[] pathArr = { "a:\\", "b:\\", "c:\\", "d:\\", "e:\\", "f:\\", "g:\\", "h:\\", "i:\\" };
            bool blockSelecting = false, fForceNew = true;
            // �����������һ���ļ�·�������Ҵ�����򿪸���Ŀ¼���߸���Ŀ¼
            if (File.Exists(clipPath))
            {
                try
                {
                    QTUtility2.log("tabControl1_PlusButtonClicked file exist " + clipPath);
                    string pathRoot = Path.GetPathRoot(clipPath);
                    DirectoryInfo di = new DirectoryInfo(clipPath);
                    if (Directory.Exists(di.Parent.FullName))
                    {
                        OpenNewTab(di.Parent.FullName, blockSelecting, fForceNew);

                    }
                    else
                    {
                        OpenNewTab(pathRoot, blockSelecting, fForceNew);
                    }

                    string selectMe = Path.GetFileName(clipPath);
                    ShellBrowser.TrySetSelection(new Address[] { new Address(selectMe) }, null, true);
                }
                catch (Exception e1)
                {
                    QTUtility2.MakeErrorLog(e1, "tabControl1_PlusButtonClicked for file:" + clipPath);
                    openDefault();
                }
            }
            else if (Directory.Exists(clipPath)) // ������ֱ����һ��Ŀ¼��򿪱�ǩ
            {
                try
                {
                    QTUtility2.log("tabControl1_PlusButtonClicked Directory exist " + clipPath);
                    OpenNewTab(clipPath, blockSelecting, fForceNew);
                }
                catch (Exception e1)
                {
                    QTUtility2.MakeErrorLog(e1, "tabControl1_PlusButtonClicked for director :" + clipPath);
                    openDefault();
                }
            }
            else
            {

                // ��ָ���̷�Ŀ¼
                /*
                for ( int i = 0; i < pathArr.Length; i++ )
                {
                    if (Directory.Exists(pathArr[i]))
                    {
                        OpenNewTab(pathArr[i], blockSelecting, fForceNew);
                        break;
                    }
                }
                */
                openDefault();
            }
        }

        private void openDefault()
        {
            bool isOpend = false;
            // �����õ�·����ַ
            using (IDLWrapper wrapper = new IDLWrapper(Config.Window.DefaultLocation))
            {
                QTUtility2.log("tabControl1_PlusButtonClicked others default " + Config.Window.DefaultLocation);
                OpenNewTab(wrapper, false, true);
                isOpend = true;
            }

            if (!isOpend)
            {
                string idl = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"; // �ҵĵ��ԣ� Ĭ�ϴ�
                using (IDLWrapper w = new IDLWrapper(idl))
                {
                    QTUtility2.log("tabControl1_PlusButtonClicked �ҵĵ��� ");
                    OpenNewTab(w, false, true);
                }
            }
        }

        protected override int CalcBandHeight(int count)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
