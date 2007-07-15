import sys, time, random

for i in range(8):
    print ' (%s) long running process with intermittent output %s' % \
        (sys.argv[1], i, )
    time.sleep (random.uniform (0.1, 1.0))
