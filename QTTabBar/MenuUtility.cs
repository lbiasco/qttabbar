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
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using QTTabBarLib.Interop;

namespace QTTabBarLib {
    internal static class MenuUtility {

        [ThreadStatic]
        public static bool InMenuLoop;

        private static Font StartUpTabFont;


        // TODO: this is absent from Quizo's sources.  Figure out why.
        private static void AddChildrenOnOpening(DirectoryMenuItem parentItem) {
            bool fTruncated;
            DirectoryInfo info = new DirectoryInfo(parentItem.Path);
            EventPack eventPack = parentItem.EventPack;
            foreach(DirectoryInfo info2 in info.GetDirectories()
                    .Where(info2 => (info2.Attributes & FileAttributes.Hidden) == 0)) {
                string text = QTUtility2.MakeNameEllipsis(info2.Name, out fTruncated);
                DropDownMenuReorderable reorderable = new DropDownMenuReorderable(null);
                reorderable.MessageParent = eventPack.MessageParentHandle;
                reorderable.ItemRightClicked += eventPack.ItemRightClickEventHandler;
                reorderable.ImageList = QTUtility.ImageListGlobal;
                DirectoryMenuItem item = new DirectoryMenuItem(text);
                item.SetImageReservationKey(info2.FullName, null);
                item.Path = info2.FullName;
                item.EventPack = eventPack;
                item.ModifiedDate = info2.LastWriteTime;
                if(fTruncated) {
                    item.ToolTipText = info2.Name;
                }
                item.DropDown = reorderable;
                item.DoubleClickEnabled = true;
                item.DropDownItems.Add(new ToolStripMenuItem());
                item.DropDownItemClicked += realDirectory_DropDownItemClicked;
                item.DropDownOpening += realDirectory_DropDownOpening;
                item.DoubleClick += eventPack.DirDoubleClickEventHandler;
                parentItem.DropDownItems.Add(item);                
            }
            foreach(FileInfo info3 in info.GetFiles()
                    .Where(info3 => (info3.Attributes & FileAttributes.Hidden) == 0)) {
                string fileNameWithoutExtension;
                string ext = info3.Extension.ToLower();
                switch(ext) {
                    case ".lnk":
                    case ".url":
                        fileNameWithoutExtension = Path.GetFileNameWithoutExtension(info3.Name);
                        break;

                    default:
                        fileNameWithoutExtension = info3.Name;
                        break;
                }
                string str4 = fileNameWithoutExtension;
                QMenuItem item2 = new QMenuItem(QTUtility2.MakeNameEllipsis(fileNameWithoutExtension, out fTruncated), MenuTarget.File, MenuGenre.Application);
                item2.Path = info3.FullName;
                item2.SetImageReservationKey(info3.FullName, ext);
                item2.MouseMove += qmi_File_MouseMove;
                if(fTruncated) {
                    item2.ToolTipText = str4;
                }
                parentItem.DropDownItems.Add(item2);
            }
        }

        public static QMenuItem CreateMenuItem(MenuItemArguments mia) {
            QMenuItem item = new QMenuItem(QTUtility2.MakePathDisplayText(mia.Path, false), mia);
            if(((mia.Genre == MenuGenre.Navigation) && mia.IsBack) && (mia.Index == 0)) {
                item.ImageKey = "current";
            }
            else {
                item.SetImageReservationKey(mia.Path, null);
            }
            item.ToolTipText = QTUtility2.MakePathDisplayText(mia.Path, true);
            return item;
        }

        // todo: check vs quizo's
        public static List<ToolStripItem> CreateRecentFilesItems() {
            List<ToolStripItem> ret = new List<ToolStripItem>();
            List<string> toRemove = new List<string>();
            if(StaticReg.ExecutedPathsList.Count > 0) {
                foreach(string path in StaticReg.ExecutedPathsList.Reverse()) {
                    if(QTUtility2.IsNetworkPath(path) || File.Exists(path)) {
                        QMenuItem item = new QMenuItem(QTUtility2.MakeNameEllipsis(Path.GetFileName(path)), MenuGenre.RecentFile);
                        item.Path = item.ToolTipText = path;
                        item.SetImageReservationKey(path, Path.GetExtension(path));
                        ret.Add(item);
                    }
                    else {
                        toRemove.Add(path);
                    }   
                }
            }
            foreach(string str in toRemove) {
                StaticReg.ExecutedPathsList.Remove(str);
            }
            return ret;
        }

        // todo: check vs quizo's
        public static List<ToolStripItem> CreateUndoClosedItems(ToolStripDropDownItem dropDownItem) {
            List<ToolStripItem> ret = new List<ToolStripItem>();
            string[] reversedLog = StaticReg.ClosedTabHistoryList.Reverse().ToArray();
            if(dropDownItem != null) {
                while(dropDownItem.DropDownItems.Count > 0) {
                    dropDownItem.DropDownItems[0].Dispose();
                }
            }
            if(reversedLog.Length > 0) {
                if(dropDownItem != null) {
                    dropDownItem.Enabled = true;
                }
                foreach(string entry in reversedLog) {
                    if(entry.Length <= 0) continue;
                    if(!QTUtility2.PathExists(entry)) {
                        StaticReg.ClosedTabHistoryList.Remove(entry);
                    }
                    else {
                        QMenuItem item = CreateMenuItem(new MenuItemArguments(entry, MenuTarget.Folder, MenuGenre.History));
                        if(dropDownItem != null) {
                            dropDownItem.DropDownItems.Add(item);
                        }
                        ret.Add(item);
                    }
                }
            }
            else if(dropDownItem != null) {
                dropDownItem.Enabled = false;
            }
            return ret;
        }

        private static void qmi_File_MouseMove(object sender, MouseEventArgs e) {
            QMenuItem item = (QMenuItem)sender;
            if(item.ToolTipText != null || string.IsNullOrEmpty(item.Path)) return;
            string str = item.Path.StartsWith("::") ? item.Text : Path.GetFileName(item.Path);
            string shellInfoTipText = ShellMethods.GetShellInfoTipText(item.Path, false);
            if(shellInfoTipText != null) {
                if(str == null) {
                    str = shellInfoTipText;
                }
                else {
                    str = str + "\r\n" + shellInfoTipText;
                }
            }
            item.ToolTipText = str;
        }

        private static void realDirectory_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e) {
            if (e.ClickedItem is DirectoryMenuItem) { return;}

            try {
                Process.Start(((QMenuItem)e.ClickedItem).Path);
            } catch {
                MessageBox.Show(
                    String.Format(
                        QTUtility.TextResourcesDic["ErrorDialogs"][0],
                        e.ClickedItem.Name
                    ),
                    QTUtility.TextResourcesDic["ErrorDialogs"][1],
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Hand
                );
            }
        }

        private static void realDirectory_DropDownOpening(object sender, EventArgs e) {
            DirectoryMenuItem parentItem = (DirectoryMenuItem)sender;
            if(!parentItem.OnceOpened) {
                parentItem.OnceOpened = true;
                parentItem.DropDown.SuspendLayout();
                parentItem.DropDownItems[0].Dispose();
                AddChildrenOnOpening(parentItem);
                parentItem.DropDown.ResumeLayout();
                if(!QTUtility.IsXP) {
                    parentItem.DropDown.BringToFront();
                }
            }
            else {
                DateTime lastWriteTime = Directory.GetLastWriteTime(parentItem.Path);
                if(parentItem.ModifiedDate != lastWriteTime) {
                    parentItem.DropDown.SuspendLayout();
                    parentItem.ModifiedDate = lastWriteTime;
                    while(parentItem.DropDownItems.Count > 0) {
                        parentItem.DropDownItems[0].Dispose();
                    }
                    AddChildrenOnOpening(parentItem);
                    parentItem.DropDown.ResumeLayout();
                }
            }
        }
    }

    internal sealed class EventPack {
        public EventHandler DirDoubleClickEventHandler;
        public bool FromTaskBar;
        public ItemRightClickedEventHandler ItemRightClickEventHandler;
        public IntPtr MessageParentHandle;

        public EventPack(IntPtr hwnd, ItemRightClickedEventHandler handlerRightClick, EventHandler handlerDirDblClick, bool fFromTaskBar) {
            MessageParentHandle = hwnd;
            ItemRightClickEventHandler = handlerRightClick;
            DirDoubleClickEventHandler = handlerDirDblClick;
            FromTaskBar = fFromTaskBar;
        }
    }
}
