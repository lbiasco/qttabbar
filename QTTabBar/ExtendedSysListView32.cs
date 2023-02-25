//    This file is part of QTTabBar, a shell extension for Microsoft
//    Windows Explorer.
//    Copyright (C) 2007-2021  Quizo, Paul Accisano
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
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using QTTabBarLib.Interop;

namespace QTTabBarLib {
    internal class ExtendedSysListView32 : ExtendedListViewCommon {

        private static SolidBrush sbAlternate;
        private NativeWindowController EditController;
        private List<int> lstColumnFMT;
        private bool fListViewHasFocus;
        private int iListViewItemState;
        private int iHotItem = -1;


        internal ExtendedSysListView32(ShellBrowserEx ShellBrowser, IntPtr hwndShellView, IntPtr hwndListView, IntPtr hwndSubDirTipMessageReflect)
                : base(ShellBrowser, hwndShellView, hwndListView, hwndSubDirTipMessageReflect) {
            SetStyleFlags();
        }

        private int CorrectHotItem(int iItem) {
            if(QTUtility.IsXP && iItem == -1 && ShellBrowser.ViewMode == FVM.DETAILS && ShellBrowser.GetItemCount() > 0) {
                RECT rect = GetItemRect(0, LVIR.LABEL);
                Point mousePosition = Control.MousePosition;
                PInvoke.ScreenToClient(Handle, ref mousePosition);
                int minX = Math.Min(rect.left, rect.right);
                int maxX = Math.Max(rect.left, rect.right);
                if(minX <= mousePosition.X && mousePosition.X <= maxX) {
                    iItem = HitTest(new Point(minX + 2, mousePosition.Y), false);
                }
            }
            return iItem;
        }

        private bool EditController_MessageCaptured(ref Message msg) {
            // QTUtility2.debugMessage(msg);
            if(msg.Msg == 0xb1 /* EM_SETSEL */ && msg.WParam.ToInt32() != -1) {
                msg.LParam = EditController.OptionalHandle;
                EditController.MessageCaptured -= EditController_MessageCaptured;
            }
            return false;
        }

        protected override bool OnShellViewNotify(NMHDR nmhdr, ref Message msg) {
            // Process WM.NOTIFY.  These are all notifications from the 
            // SysListView32 control.  We will not get ANY of these on 
            // Windows 7, which means every single one of them has to 
            // have an alternative somewhere for the ItemsView control,
            // or it's not going to happen.
            switch(nmhdr.code) {
                case -12: // NM_CUSTOMDRAW
                    // This is for drawing alternating row colors.  I doubt
                    // very much we'll find an alternative for this...
                    return HandleCustomDraw(ref msg);

                case LVN.ITEMCHANGED: {
                        QTUtility2.log("LVN.ITEMCHANGED");
                        bool flag = false;
                        NMLISTVIEW nmlistview2 = (NMLISTVIEW)Marshal.PtrToStructure(msg.LParam, typeof(NMLISTVIEW));
                        if(nmlistview2.uChanged == 8 /*LVIF_STATE*/) {
                            uint newSelected = nmlistview2.uNewState & LVIS.SELECTED;
                            uint oldSelected = nmlistview2.uOldState & LVIS.SELECTED;
                            uint newDrophilited = nmlistview2.uNewState & LVIS.DROPHILITED;
                            uint oldDrophilited = nmlistview2.uOldState & LVIS.DROPHILITED;
                            uint newCut = nmlistview2.uNewState & LVIS.CUT;
                            uint oldCut = nmlistview2.uOldState & LVIS.CUT;
                            if(flag) {
                                if (nmlistview2.iItem != -1 && 
                                    ((newSelected != oldSelected) || 
                                     (newDrophilited != oldDrophilited) || 
                                     (newCut != oldCut)) &&
                                    ShellBrowser.ViewMode == FVM.DETAILS)
                                {
                                    QTUtility2.log("LVN.ITEMCHANGED nmlistview2.iItem " + nmlistview2.iItem);
                                    PInvoke.SendMessage(nmlistview2.hdr.hwndFrom, LVM.REDRAWITEMS, (IntPtr)nmlistview2.iItem, (IntPtr)nmlistview2.iItem);
                                }
                            }
                            if(newSelected != oldSelected) {
                                QTUtility2.log("newSelected != oldSelected  OnSelectionChanged " );
                                OnSelectionChanged(ref msg);
                            }
                        }
                        break;
                    }

                case LVN.INSERTITEM:
                case LVN.DELETEITEM:
                    // Handled through undocumented WM_USER+174 message
                    if(Config.Tips.ShowSubDirTips) {
                        HideSubDirTip(1);
                    }
                    ShellViewController.DefWndProc(ref msg);
                    OnItemCountChanged();
                    return true;

                case LVN.BEGINDRAG:
                    // This won't be necessary it seems.  On Windows 7, when you
                    // start to drag, a MOUSELEAVE message is sent, which hides
                    // the SubDirTip anyway.
                    ShellViewController.DefWndProc(ref msg);
                    HideSubDirTip(0xff);
                    return true;

                case LVN.ITEMACTIVATE: {
                    // Handled by catching Double Clicks and Enter keys.  Ugh...
                    NMITEMACTIVATE nmitemactivate = (NMITEMACTIVATE)Marshal.PtrToStructure(msg.LParam, typeof(NMITEMACTIVATE));
                    Keys modKeys =
                        (((nmitemactivate.uKeyFlags & 1) == 1) ? Keys.Alt : Keys.None) |
                        (((nmitemactivate.uKeyFlags & 2) == 2) ? Keys.Control : Keys.None) |
                        (((nmitemactivate.uKeyFlags & 4) == 4) ? Keys.Shift : Keys.None);
                    if(OnSelectionActivated(modKeys)) return true;
                    break;
                }

                case LVN.ODSTATECHANGED:
                    break;

                case LVN.HOTTRACK:
                    // Handled through WM_MOUSEMOVE.
                    if(Config.Tips.ShowTooltipPreviews || Config.Tips.ShowSubDirTips) {
                        NMLISTVIEW nmlistview = (NMLISTVIEW)Marshal.PtrToStructure(msg.LParam, typeof(NMLISTVIEW));
                        int iItem = CorrectHotItem(nmlistview.iItem);
                        if(iHotItem != iItem) {
                            OnHotItemChanged(iItem);
                            iHotItem = iItem;
                        }
                    }
                    break;

                case LVN.KEYDOWN: {
                    // Handled through WM_KEYDOWN.
                    NMLVKEYDOWN nmlvkeydown = (NMLVKEYDOWN)Marshal.PtrToStructure(msg.LParam, typeof(NMLVKEYDOWN));
                    if(OnKeyDown((Keys)nmlvkeydown.wVKey)) {
                        msg.Result = (IntPtr)1;
                        return true;
                    }
                    else {
                        return false;
                    }                        
                }
                    
                case LVN.GETINFOTIP: {
                    // Handled through WM_NOTIFY / TTN_NEEDTEXT
                    NMLVGETINFOTIP nmlvgetinfotip = (NMLVGETINFOTIP)Marshal.PtrToStructure(msg.LParam, typeof(NMLVGETINFOTIP));
                    return OnGetInfoTip(nmlvgetinfotip.iItem, GetHotItem() != nmlvgetinfotip.iItem); // TODO there's got to be a better way.
                }

                case LVN.BEGINLABELEDIT:
                    // This is just for file renaming, which there's no need to
                    // mess with in Windows 7.
                    ShellViewController.DefWndProc(ref msg);
                    break;

                case LVN.ENDLABELEDIT: {
                    // TODO
                    NMLVDISPINFO nmlvdispinfo2 = (NMLVDISPINFO)Marshal.PtrToStructure(msg.LParam, typeof(NMLVDISPINFO));
                    OnEndLabelEdit(nmlvdispinfo2.item);
                    break;
                }
            }
            return false;
        }

        private void SetStyleFlags()
        {
            if (ShellBrowser == null) return;  // qt desktop tool 启用空指针问题 https://www.yuque.com/indiff/lc0r1g/kqgkr0
            if(ShellBrowser.ViewMode != FVM.DETAILS) return;
            uint flags = 0;
            flags &= ~LVS_EX.GRIDLINES;
            flags &= ~LVS_EX.FULLROWSELECT;
            const uint mask = LVS_EX.GRIDLINES | LVS_EX.FULLROWSELECT;
            PInvoke.SendMessage(Handle, LVM.SETEXTENDEDLISTVIEWSTYLE, (IntPtr)mask, (IntPtr)flags);
        }

        protected override IntPtr GetEditControl() {
            return PInvoke.SendMessage(Handle, LVM.GETEDITCONTROL, IntPtr.Zero, IntPtr.Zero);
        }

        protected override Rectangle GetFocusedItemRect() {
            if(HasFocus()) {
                int code = ShellBrowser.ViewMode == FVM.DETAILS ? LVIR.LABEL : LVIR.BOUNDS;
                return GetItemRect(ShellBrowser.GetFocusedIndex(), code).ToRectangle();
            }
            return new Rectangle(0, 0, 0, 0);
        }

        public override int GetHotItem() {
            return CorrectHotItem(base.GetHotItem());
        }

        protected override Point GetSubDirTipPoint(bool fByKey) {
            int iItem = fByKey ? ShellBrowser.GetFocusedIndex() : GetHotItem();
            int x, y;
            Point ret;
            RECT rect;
            switch(ShellBrowser.ViewMode) {
                case FVM.DETAILS:
                    rect = GetItemRect(iItem, LVIR.LABEL);
                    x = rect.right;
                    y = rect.top;
                    y += (rect.bottom - y)/2;
                    ret = new Point(x - 19, y - 7);
                    break;

                case FVM.SMALLICON:
                    rect = GetItemRect(iItem);
                    x = rect.right;
                    y = rect.top;
                    x -= (rect.bottom - y) / 2;
                    y += (rect.bottom - y) / 2;
                    ret = new Point(x - 9, y - 7);
                    break;

                case FVM.CONTENT:
                case FVM.TILE:
                    rect = GetItemRect(iItem, LVIR.ICON);
                    y = rect.bottom;
                    x = rect.right;
                    ret = new Point(x - 16, y - 16);
                    break;

                case FVM.THUMBSTRIP:
                case FVM.THUMBNAIL:
                    rect = GetItemRect(iItem, LVIR.ICON);
                    if(QTUtility.IsXP) rect.right -= 13;
                    y = rect.bottom;
                    x = rect.right;
                    ret = new Point(x - 16, y - 16);
                    break;

                case FVM.ICON:
                    rect = GetItemRect(iItem, LVIR.ICON);
                    if(QTUtility.IsXP) {
                        int num3 = (int)PInvoke.SendMessage(Handle, LVM.GETITEMSPACING, IntPtr.Zero, IntPtr.Zero);
                        Size iconSize = SystemInformation.IconSize;
                        rect.right = ((rect.left + (((num3 & 0xffff) - iconSize.Width) / 2)) + iconSize.Width) + 8;
                        rect.bottom = (rect.top + iconSize.Height) + 6;
                    }
                    y = rect.bottom;
                    x = rect.right;
                    ret = new Point(x - 16, y - 16);
                    break;

                case FVM.LIST:
                    if(QTUtility.IsXP) {
                        rect = GetItemRect(iItem, LVIR.ICON);
                        using(SafePtr pszText = new SafePtr(520)) {
                            LVITEM structure = new LVITEM {
                                pszText = pszText,
                                cchTextMax = 260,
                                iItem = iItem,
                                mask = 1
                            };
                            PInvoke.SendMessage(Handle, LVM.GETITEM, IntPtr.Zero, ref structure);
                            int num4 = (int)PInvoke.SendMessage(Handle, LVM.GETSTRINGWIDTH, IntPtr.Zero, pszText);
                            num4 += 20;
                            rect.right += num4;
                            rect.top += 2;
                            rect.bottom += 2;                            
                        }
                    }
                    else {
                        rect = GetItemRect(iItem, LVIR.LABEL);
                    }                    
                    y = rect.bottom;
                    x = rect.right;
                    ret = new Point(x - 16, y - 16);
                    break;

                default:
                    rect = GetItemRect(iItem);
                    y = rect.bottom;
                    x = rect.right;
                    ret = new Point(x - 16, y - 16);
                    break;

            }
            PInvoke.ClientToScreen(Handle, ref ret);
            return ret;
        }
        // 使用箭头键时候环绕选择文件夹
        protected override bool HandleCursorLoop(Keys key) {
            int focusedIdx = ShellBrowser.GetFocusedIndex();
            int itemCount = ShellBrowser.GetItemCount();
            int selectMe = -1;
            FVM viewMode = ShellBrowser.ViewMode;
            if(viewMode == FVM.TILE && QTUtility.IsXP) {
                viewMode = FVM.ICON;
            }
            switch(viewMode) {
                case FVM.CONTENT:
                case FVM.DETAILS:
                case FVM.TILE:
                    if(key == Keys.Up && focusedIdx == 0) {
                        selectMe = itemCount - 1;
                    }
                    else if(key == Keys.Down && focusedIdx == itemCount - 1) {
                        selectMe = 0;
                    }
                    break;

                case FVM.ICON:
                case FVM.SMALLICON:
                case FVM.THUMBNAIL:
                case FVM.LIST:
                    Keys KeyNextItem, KeyPrevItem, KeyNextPage, KeyPrevPage;
                    IntPtr MsgNextPage, MsgPrevPage;
                    if(viewMode == FVM.LIST) {
                        KeyNextItem = Keys.Down;
                        KeyPrevItem = Keys.Up;
                        KeyNextPage = Keys.Right;
                        KeyPrevPage = Keys.Left;
                        MsgNextPage = (IntPtr)LVNI.TORIGHT;
                        MsgPrevPage = (IntPtr)LVNI.TOLEFT;
                    }
                    else {
                        KeyNextItem = Keys.Right;
                        KeyPrevItem = Keys.Left;
                        KeyNextPage = Keys.Down;
                        KeyPrevPage = Keys.Up;
                        MsgNextPage = (IntPtr)LVNI.BELOW;
                        MsgPrevPage = (IntPtr)LVNI.ABOVE;
                    }

                    int nextPageIdx = (int)PInvoke.SendMessage(Handle, LVM.GETNEXTITEM, (IntPtr)focusedIdx, MsgNextPage);
                    if(nextPageIdx == -1 || nextPageIdx == focusedIdx) {
                        nextPageIdx = (int)PInvoke.SendMessage(Handle, LVM.GETNEXTITEM, (IntPtr)focusedIdx, MsgPrevPage);
                    }
                    else if(QTUtility.IsXP) {
                        int testIdx = (int)PInvoke.SendMessage(Handle, LVM.GETNEXTITEM, (IntPtr)nextPageIdx, MsgPrevPage);
                        if(testIdx != focusedIdx) {
                            nextPageIdx = (int)PInvoke.SendMessage(Handle, LVM.GETNEXTITEM, (IntPtr)focusedIdx, MsgPrevPage);
                        }
                    }
                    if(nextPageIdx == -1 || nextPageIdx == focusedIdx) {
                        if(key == KeyNextItem) {
                            if(focusedIdx == itemCount - 1) {
                                selectMe = 0;
                            }
                            else {
                                RECT thisRect = GetItemRect(focusedIdx);
                                RECT nextRect = GetItemRect(focusedIdx + 1);
                                if(viewMode == FVM.LIST) {
                                    if(nextRect.top < thisRect.top) selectMe = 0;
                                }
                                else if(nextRect.left < thisRect.left) {
                                    selectMe = 0;
                                }
                            }
                        }
                        else if(key == KeyPrevItem && focusedIdx == 0) {
                            selectMe = itemCount - 1;
                        }
                        else if(key == KeyNextPage || key == KeyPrevPage) {
                            if(QTUtility.IsXP) {
                                return true;
                            }
                        }
                    }
                    else {
                        int pageCount = Math.Abs(focusedIdx - nextPageIdx);
                        int page = focusedIdx % pageCount;
                        if(key == KeyNextItem && (page == pageCount - 1 || focusedIdx == itemCount - 1)) {
                            selectMe = focusedIdx - page;
                        }
                        else if(key == KeyPrevItem && page == 0) {
                            selectMe = Math.Min(focusedIdx + pageCount - 1, itemCount - 1);
                        }
                        else if(key == KeyNextPage && focusedIdx + pageCount >= itemCount) {
                            selectMe = page;
                        }
                        else if(key == KeyPrevPage && focusedIdx < pageCount) {
                            int x = itemCount - focusedIdx - 1;
                            selectMe = x - x % pageCount + focusedIdx;
                        }
                    }
                    break;

            }

            if(selectMe >= 0) {
                SetRedraw(false);
                ShellBrowser.SelectItem(selectMe);
                PInvoke.SendMessage(Handle, LVM.REDRAWITEMS, (IntPtr)focusedIdx, (IntPtr)focusedIdx);
                SetRedraw(true);
                return true;
            }
            else {
                return false;
            }
        }

        // 处理自定义绘制
        private bool HandleCustomDraw(ref Message msg) {
            return false;
        }

        private void OnFileRename(IDLWrapper idl) {
            if(!idl.Available || idl.IsFileSystemFolder) return;
            string path = idl.Path;
            if(File.Exists(path)) {
                string extension = Path.GetExtension(path);
                if(!string.IsNullOrEmpty(extension) && extension.PathEquals(".lnk")) {
                    return;
                }
            }
            IntPtr hWnd = GetEditControl();
            if(hWnd == IntPtr.Zero) return;

            using(SafePtr lParam = new SafePtr(520)) {
                if((int)PInvoke.SendMessage(hWnd, WM.GETTEXT, (IntPtr)260, lParam) <= 0) return;
                string str3 = Marshal.PtrToStringUni(lParam);
                if(str3.Length > 2) {
                    int num = str3.LastIndexOf(".");
                    if(num > 0) {
                        // Explorer will send the EM_SETSEL message to select the
                        // entire filename.  We will intercept this message and
                        // change the params to select only the part before the
                        // extension.
                        EditController = new NativeWindowController(hWnd);
                        EditController.OptionalHandle = (IntPtr)num;
                        EditController.MessageCaptured += EditController_MessageCaptured;
                    }
                }
            }
        }

        private RECT GetItemRect(int iItem, int LVIRCode = LVIR.BOUNDS) {
            RECT rect = new RECT {left = LVIRCode};
            PInvoke.SendMessage(Handle, LVM.GETITEMRECT, (IntPtr)iItem, ref rect);
            return rect;
        }

        public override int HitTest(Point pt, bool screenCoords) {
            if(screenCoords) {
                PInvoke.ScreenToClient(ListViewController.Handle, ref pt);
            }
            LVHITTESTINFO structure = new LVHITTESTINFO {pt = pt};
            int num = (int)PInvoke.SendMessage(ListViewController.Handle, LVM.HITTEST, IntPtr.Zero, ref structure);
            return num;
        }

        public override bool HotItemIsSelected() {
            // TODO: I don't think HOTITEM means what you think it does.
            int hot = (int)PInvoke.SendMessage(ListViewController.Handle, LVM.GETHOTITEM, IntPtr.Zero, IntPtr.Zero);
            if(hot == -1) return false;
            int state = (int)PInvoke.SendMessage(ListViewController.Handle, LVM.GETITEMSTATE, (IntPtr)hot, (IntPtr)LVIS.SELECTED);
            return ((state & LVIS.SELECTED) != 0);
        }

        public override bool IsTrackingItemName() {
            if(ShellBrowser.ViewMode == FVM.DETAILS) return true;
            if(ShellBrowser.GetItemCount() == 0) return false;
            RECT rect = PInvoke.ListView_GetItemRect(ListViewController.Handle, 0, 0, 2);
            Point mousePosition = Control.MousePosition;
            PInvoke.MapWindowPoints(IntPtr.Zero, ListViewController.Handle, ref mousePosition, 1);
            return (Math.Min(rect.left, rect.right) <= mousePosition.X) && (mousePosition.X <= Math.Max(rect.left, rect.right));
        }

        protected override bool ListViewController_MessageCaptured(ref Message msg) {
            if(base.ListViewController_MessageCaptured(ref msg)) {
                return true;
            }

            switch(msg.Msg) {
                // Style flags are reset when the view is changed.
                case LVM.SETVIEW:
                    SetStyleFlags();
                    break;

                // On Vista/7, we don't get a LVM.SETVIEW, but we do
                // get this.
                case WM.SETREDRAW:
                    if(msg.WParam != IntPtr.Zero) {
                        SetStyleFlags();
                    }
                    break;

            }
            return false;
        }

        public override bool PointIsBackground(Point pt, bool screenCoords) {
            if(screenCoords) {
                PInvoke.ScreenToClient(ListViewController.Handle, ref pt);
            }
            LVHITTESTINFO structure = new LVHITTESTINFO {pt = pt};
            if(QTUtility.IsXP) {
                return -1 == (int)PInvoke.SendMessage(ListViewController.Handle, LVM.HITTEST, IntPtr.Zero, ref structure);
            }
            else {
                PInvoke.SendMessage(ListViewController.Handle, LVM.HITTEST, (IntPtr)(-1), ref structure);
                return structure.flags == 1 /* LVHT_NOWHERE */;
            }
        }
    }
}
