# PyConsole project
# Copyright (C) 2007 Michael Graz
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

import wx, wx.stc, pyconsole

ID_ABOUT      = wx.NewId()
ID_EXIT       = wx.NewId()
ID_DEBUG      = wx.NewId()
ID_FS_PWD     = wx.NewId()
ID_FS_DIR     = wx.NewId()
ID_BK_HOME    = wx.NewId()
ID_BK_DOCS    = wx.NewId()
ID_BK_PYCON   = wx.NewId()
ID_NET_YAHOO  = wx.NewId()
ID_NET_GOOGLE = wx.NewId()
ID_NET_LOCAL  = wx.NewId()
ID_CONSOLE    = wx.NewId()

#----------------------------------------------------------------------

class MyFrame(wx.Frame):
    def __init__(self, parent, ID, title):
        wx.Frame.__init__(self, parent, ID, title,
                         wx.DefaultPosition, wx.Size(700, 500))

        self.CreateStatusBar()
        self.SetStatusText("")
        menu_file = wx.Menu()
        menu_file.Append(ID_DEBUG, "De&bug", "debug")
        menu_file.Append(ID_ABOUT, "&About", "More information about this program")
        menu_file.AppendSeparator()
        menu_file.Append(ID_EXIT, "E&xit", "Terminate the program")
        menu_bar = wx.MenuBar()
        menu_bar.Append(menu_file, "&File");

        menu_file_system = wx.Menu ()
        menu_file_system.Append (ID_FS_DIR, "&Dir", "Directory listing")
        menu_file_system.Append (ID_FS_PWD, "&Pwd", "Print current directory")
        menu_bar.Append(menu_file_system, "File&System");

        menu_bookmarks = wx.Menu ()
        menu_bookmarks.Append (ID_BK_HOME, "&Home", "Home directory")
        menu_bookmarks.Append (ID_BK_DOCS, "&Docs", "My Documents")
        menu_bookmarks.Append (ID_BK_PYCON, "&PyConsole", "PyConsole source")
        menu_bar.Append(menu_bookmarks, "&Bookmark");

        menu_network = wx.Menu ()
        # menu_network.Append (ID_NET_LOCAL, "localhost", "Ping localhost")
        menu_network.Append (ID_NET_YAHOO, "www.&yahoo.com", "Ping Yahoo")
        menu_network.Append (ID_NET_GOOGLE, "www.&google.com", "Ping Google")
        menu_bar.Append(menu_network, "&Network");

        self.SetMenuBar(menu_bar)

        wx.EVT_MENU(self, ID_DEBUG, self.debug)
        wx.EVT_MENU(self, ID_ABOUT, self.OnAbout)
        wx.EVT_MENU(self, ID_EXIT,  self.TimeToQuit)
        wx.EVT_MENU(self, ID_FS_PWD,  self.FsPwd)
        wx.EVT_MENU(self, ID_FS_DIR,  self.FsDir)
        wx.EVT_MENU(self, ID_BK_HOME,  self.BkHome)
        wx.EVT_MENU(self, ID_BK_DOCS,  self.BkDocs)
        wx.EVT_MENU(self, ID_BK_PYCON,  self.BkPyCon)
        wx.EVT_MENU(self, ID_NET_YAHOO,  self.NetYahoo)
        wx.EVT_MENU(self, ID_NET_GOOGLE,  self.NetGoogle)
        wx.EVT_MENU(self, ID_NET_LOCAL,  self.NetLocal)

        self.text_ctrl = ConsoleProcessWindow (parent=self)

    def FsPwd (self, event):
        self.text_ctrl.sendline ('pwd')

    def FsDir (self, event):
        self.text_ctrl.sendline ('dir')

    def BkHome (self, event):
        self.text_ctrl.sendline (r'cd C:\Home')

    def BkDocs (self, event):
        self.text_ctrl.sendline (r'cd C:\Documents and Settings\Mike\My Documents')

    def BkPyCon (self, event):
        self.text_ctrl.sendline (r'cd C:\PyConsole')

    def NetYahoo (self, event):
        self.text_ctrl.sendline (r'ping www.yahoo.com')

    def NetGoogle (self, event):
        self.text_ctrl.sendline (r'ping www.google.com')

    def NetLocal (self, event):
        self.text_ctrl.sendline (r'ping localhost')

    def debug (self, event):
        self.text_ctrl.GotoLine (1)

    def OnAbout(self, event):
        dlg = wx.MessageDialog(self, "This sample program shows off\n"
                              "frames, menus, statusbars, and this\n"
                              "message dialog.",
                              "About Me", wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()

    def TimeToQuit(self, event):
        self.Close(True)

#----------------------------------------------------------------------

class ConsoleEvent (wx.PyEvent):
    def __init__(self, x, y, text):
        wx.PyEvent.__init__ (self)
        self.SetEventType (ID_CONSOLE)
        self.x = x
        self.y = y
        self.text = text

#----------------------------------------------------------------------

class ConsoleProcessWindow (wx.stc.StyledTextCtrl):

    def __init__ (self, *args, **kwargs):
        wx.stc.StyledTextCtrl.__init__ (self, *args, **kwargs)
        # self.StyleSetFont (0, wx.SystemSettings.GetFont (wx.SYS_ANSI_FIXED_FONT))
        # self.StyleSetSpec (wx.stc.STC_STYLE_DEFAULT, "size:8,face:Courier")
        # self.StyleSetSpec (wx.stc.STC_STYLE_DEFAULT, "size:20,face:Helvetica")
        # self.StyleSetSpec (wx.stc.STC_STYLE_DEFAULT, "size:10,face:OCRB")
        self.StyleSetSpec (wx.stc.STC_STYLE_DEFAULT, "size:10,face:Lucida Console")
        self.line_count = 0
        self.last_update_pos = 0
        self.last_update_line = 0
        self.prompt_len = 0
        self.lst_history = []
        self.auto_comp = False
        self.tab_key_down = False
        self.Bind(wx.EVT_KEY_UP, self.OnKeyUp)
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        self.Connect(-1, -1, ID_CONSOLE, self.OnConsole)
        self.console_process = pyconsole.ConsoleProcess ('cmd.exe',
                console_update=self.console_update)

    def OnKeyDown(self, event):
        self.auto_comp = self.AutoCompActive()
        self.tab_key_down = not self.auto_comp and not event.Modifiers and event.GetKeyCode() == wx.WXK_TAB
        event.Skip ()

    def OnKeyUp(self, event):
        # if auto-completing just return
        if self.auto_comp:
            return
        # make sure ALT, CTRL, SHIFT not pressed
        if event.Modifiers:
            event.Skip ()
            return
        key = event.GetKeyCode()
        if key == wx.WXK_RETURN:
            self.line_count += 1
            text = self.GetTextRange (self.last_update_pos, self.CurrentPos).rstrip()
            self.send (text + '\n')
            event.Skip ()
            return
        if key == wx.WXK_TAB and self.tab_key_down:
            if not self.AutoCompActive():
                text = self.GetTextRange (self.last_update_pos, self.CurrentPos)
                self.send (text + '\t')
            event.Skip ()
            return
        event.Skip ()
        text, pos = self.GetCurLine()
        if pos - self.prompt_len == 1:
            text = text[self.prompt_len:]
            text_lower = text.lower()
            lst_show = [x for x in self.lst_history if x.lower().startswith (text_lower)]
            if lst_show:
                separator = '\x01'
                self.AutoCompSetSeparator (ord(separator))
                self.AutoCompShow (1, separator.join (lst_show))

    def OnConsole (self, event):
        self.do_console_update (event.x, event.y, event.text)

    def send (self, text):
        if text.endswith ('\n'):
            input = text.strip ()
            if not input in self.lst_history:
                self.lst_history.append (input)
        self.console_process.write (text)

    def sendline (self, text):
        self.send (text + '\n')

    def console_update (self, x, y, text):
        wx.PostEvent (self, ConsoleEvent (x, y, text))

    def do_console_update (self, x, y, text):
        if y > 0 and y < self.line_count:
            return  # TODO this needs to be more sophisticated
        if y == self.line_count:
            self.GotoLine (y)
            self.CurrentPos += x
            self.DelLineRight ()
        else:
            while y > self.line_count:
                self.AppendText('\n')
                self.line_count += 1
                self.GotoLine (self.line_count)
        if y > self.last_update_line:
            self.prompt_len = len(text)
            self.last_update_line = y
        self.AppendText(text)
        self.CurrentPos = self.TextLength
        self.last_update_pos = self.CurrentPos
        self.SelectionStart = self.SelectionEnd

#----------------------------------------------------------------------

class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, -1, "Hello from wxPython")
        frame.Show(True)
        self.SetTopWindow(frame)
        return True

app = MyApp(0)
app.MainLoop()

