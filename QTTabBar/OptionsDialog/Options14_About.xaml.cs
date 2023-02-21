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
using System.Diagnostics;

namespace QTTabBarLib {
    internal partial class Options14_About : OptionsDialogTab {
        public Options14_About() {
            InitializeComponent();
        }

        public override void InitializeConfig() {
            try {
                // 设置默认的title 和版本
                string str = QTUtility.CurrentVersion.ToString();
                if (QTUtility.BetaRevision.Major > 0)
                {
                    str = str + " Beta " + QTUtility.BetaRevision.Major;
                }
                else if (QTUtility.BetaRevision.Minor > 0)
                {
                    str = str + " Alpha " + QTUtility.BetaRevision.Minor;
                }

                lblVersion.Content = "QTTabBar " + str + " " + QTUtility2.MakeVersionString();
            }
            catch (Exception exception)
            {
                QTUtility2.MakeErrorLog(exception, "Options14_About InitializeConfig");
            }         
        }

        public override void ResetConfig() {    
        }

        public override void CommitConfig() {
        }
    }
}
