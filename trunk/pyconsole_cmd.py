import pyconsole

y_last = 0

def output (x, y, text):
    global y_last
    if y > y_last:
        print '---', text
    y_last = y

cmd = pyconsole.ConsoleProcess ('cmd.exe', console_update=output)

