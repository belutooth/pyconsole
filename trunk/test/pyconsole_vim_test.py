import os, sys, time, unittest, tempfile, subprocess
import win32con, win32process, win32com.client
sys.path.insert (0, '..')

def get_gvim_exe ():
    import win32api, win32con
    hkey = win32api.RegOpenKey (win32con.HKEY_LOCAL_MACHINE, r'SOFTWARE\Vim\Gvim')
    return win32api.RegQueryValueEx (hkey, 'path')[0]

def get_vim_exe ():
    return os.path.join (os.path.dirname (get_gvim_exe()), 'vim.exe')

def get_this_file ():
    try: fn = __file__
    except: fn = sys.argv[0]
    return os.path.abspath (fn)

def get_this_dir ():
    return os.path.dirname (get_this_file())

def get_this_path (*args):
    return os.path.abspath (os.path.join (get_this_dir(), *args))

# HKEY_CLASSES_ROOT\Vim.Application\CLSID

# Assumptions
# - gvim.exe is installed and registered
# - gvim.exe has clientserver and python support
# - vim.exe is in the same directory as gvim.exe

# basic flow

# construct a vim server name suffixed by timestamp
# start gvim with that servername, start with no plugins
# source pyconsole_vim.vim file from directory above
# wait a few seconds
# send keystrokes to echo server name
# save keystrokes to save buffer as a file
# look at file to make sure that the last line matches
#
# longer test
# create a file of 5000 lines
# send a command to type the file
# compare the output of the file

class VimControl:
    def __init__ (self):
        self.gvim_exe = get_gvim_exe()
        self.vim_exe = get_vim_exe()
        self.server_name = None
        self.gvim_pid = None
        self.wsh = win32com.client.Dispatch ('WScript.Shell')

    def _run_vim (self, *args):
        lst_cmd_line = (self.vim_exe, ) + args
        process = subprocess.Popen (lst_cmd_line, stdout=subprocess.PIPE)
        output = process.communicate ()
        return output [0]

    def _run_gvim (self, *args):
        flags = win32process.NORMAL_PRIORITY_CLASS
        si = win32process.STARTUPINFO()
        # si.wShowWindow = win32con.SW_HIDE
        # si.wShowWindow = win32con.SW_MINIMIZE
        cmd_line = '%s %s' % (self.gvim_exe, ' '.join(args))
        tpl_result = win32process.CreateProcess (None, cmd_line, None, None, 0, flags, None, '.', si)
        self.gvim_pid = tpl_result[2]
        time.sleep (0.5)

    def get_unique_name (self):
        name = time.strftime('%H%M%S')
        if name.upper() in self.get_lst_server():
            raise Exception ('Vim server name: %s already in use' % name)
        return name

    def get_lst_server (self):
        lst_output = self._run_vim ('--serverlist')
        return lst_output.splitlines()

    def run_unique_gvim (self, *args):
        self.server_name = self.get_unique_name ()
        lst_arg = ['--servername', self.server_name] + list(args)
        self._run_gvim (*lst_arg)
        return self.server_name

    def send_keys (self, text):
        if not self.gvim_pid:
            raise Exception ('No pid for gvim')
        self.wsh.AppActivate (self.gvim_pid)
        time.sleep (0.1)
        self.wsh.SendKeys (text)

    def remote_send (self, text):
        self._run_vim ('--servername', self.server_name, '--remote-send', text)

    def remote_expr (self, *args):
        return self._run_vim ('--servername', self.server_name, '--remote-expr', *args)

#----------------------------------------------------------------------

def _get_file_tmp ():
    fd, filename = tempfile.mkstemp ('.tmp', 'pyconsole_vim_test_')
    os.close (fd)
    os.remove (filename)
    return filename

class SimpleTestCase (unittest.TestCase):
    vc = None
    file_tmp = _get_file_tmp()

    def setUp (self):
        if SimpleTestCase.vc is None:
            SimpleTestCase.vc = VimControl()
            self.vc.run_unique_gvim ('--noplugin')
            self.vc.remote_send (':source ../pyconsole_vim.vim\n')
            self.vc.remote_send (':call PyConsole()\n')
            time.sleep (2)
        self.vim_line_count = int(self.vc.remote_expr ("line('$')"))
        # print 'xx self.vim_line_count:', self.vim_line_count
        self.vim_line_value = self.vc.remote_expr ("getbufline('%','$')[0]")[:-2]
        # print 'xx self.vim_line_value: >%s<' % self.vim_line_value

    def get_text (self):
        self.vc.remote_send ('<esc>:w '+self.file_tmp+'<cr>')
        self.vc.remote_send ('A')
        time.sleep (0.2)
        fp = file(self.file_tmp)
        lst_text = fp.read().splitlines()
        fp.close ()
        os.remove (self.file_tmp)
        first_line = lst_text[self.vim_line_count-1]
        # print 'xx first_line (1): >%s<' % first_line
        first_line = first_line[len(self.vim_line_value):]
        # print 'xx first_line (2): >%s<' % first_line
        return [first_line] + lst_text[self.vim_line_count:]

    def xx_test_simple_1 (self):
        self.vc.send_keys ('echo '+self.vc.server_name+'\r')
        time.sleep (0.5)
        lst_text = self.get_text ()
        lst_expected = ['echo '+self.vc.server_name, self.vc.server_name]
        self.assertEqual (lst_expected, lst_text[:2])

    def xx_test_simple_2 (self):
        self.vc.send_keys ('echo abcdefghijklmnopqrstuvwxyz\r')
        time.sleep (0.5)
        lst_text = self.get_text ()
        lst_expected = ['echo abcdefghijklmnopqrstuvwxyz', 'abcdefghijklmnopqrstuvwxyz']
        self.assertEqual (lst_expected, lst_text[:2])

    def test_many_lines (self):
        fd_big_tmp, file_big_tmp = tempfile.mkstemp ('.tmp', 'pyconsole_vim_test_')
        print 'xx file_big_tmp:', file_big_tmp
        fp_big_tmp = file(file_big_tmp, 'w')
        os.close (fd_big_tmp)
        line_count_test = 4000
        for i in xrange(line_count_test):
            print >>fp_big_tmp, '%s abcdefghijklmnopqrstuvwxyz' % i
        fp_big_tmp.close ()
        self.vc.send_keys ('type %s\r' % file_big_tmp)
        # wait for line count to stop updating
        vim_line_count_previous = 0
        for i in xrange(200):
            time.sleep (0.25)
            vim_line_count_current = int(self.vc.remote_expr ("line('$')"))
            print 'xx %d: vim_line_count_current: %s' % (i, vim_line_count_current, )
            if vim_line_count_current == vim_line_count_previous:
                break
            vim_line_count_previous = vim_line_count_current
        os.remove (file_big_tmp)
        lst_text = self.get_text ()
        for i in xrange(1, line_count_test+1):
            self.assertEquals (str(i-1), lst_text[i].split()[0])

#----------------------------------------------------------------------

if __name__ == '__main__':
    unittest.main ()
