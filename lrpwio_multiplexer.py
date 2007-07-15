import time, pyconsole

class SampleConsoleProcess (pyconsole.ConsoleProcess):
    def __init__ (self, cmd_line):
        pyconsole.ConsoleProcess.__init__ (self, cmd_line,
            console_update=self.console_update,
            console_process_end=self.console_process_end)
        self.running = True

    def console_update (self, x, y, text):
        print 'pyconsole:', text.strip()

    def console_process_end (self):
        self.running = False

p1 = SampleConsoleProcess('python lrpwio.py One')
p2 = SampleConsoleProcess('python lrpwio.py Two')
p3 = SampleConsoleProcess('python lrpwio.py Three')

while p1.running or p2.running or p3.running:
    time.sleep (0.1)

