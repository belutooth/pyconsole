import time, Queue, pyconsole

def ConsoleProcessGenerator (cmd_line):
    queue = Queue.Queue()
    class CP (pyconsole.ConsoleProcess):
        def __init__ (self, cmd_line):
            self.running = True
            pyconsole.ConsoleProcess.__init__ (self, cmd_line,
                console_update=self.console_update,
                console_process_end=self.console_process_end)
        def console_update (self, x, y, text):
            queue.put (text.strip())
        def console_process_end (self):
            self.running = False
    process = CP (cmd_line)
    while process.running:
        yield queue.get()

for line in ConsoleProcessGenerator('python lrpwio.py Queue'):
    print 'pyconsole:', line
