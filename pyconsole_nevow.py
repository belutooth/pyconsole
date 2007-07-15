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

import os, sys, time
import simplejson
from twisted.python import util
from nevow import rend, athena, loaders, static
import pyconsole

#----------------------------------------------------------------------

class ConsoleProcessNevow (athena.LiveFragment):
    jsClass = u'PyConsole.Nevow'
    docFactory = loaders.xmlfile (util.sibpath (__file__, 'pyconsole_nevow.html'))

    def __init__ (self, *a, **kw):
        super (ConsoleProcessNevow, self).__init__ (*a, **kw)
        self.console_process = pyconsole.ConsoleProcess ('cmd.exe',
                console_update_many=self.console_update_many)
        self.line_count = 0

    def cp_input (self, text):
        self.console_process.write (text + '\n')
    athena.expose (cp_input)

    def console_update_many (self, lst_update):
        lst_html = []
        for msg_type, x, y, text_len, text in lst_update:
            if y < self.line_count:
                continue
            linefeeds = y - self.line_count
            if linefeeds > 0:
                lst_html.append ('<br/>' * linefeeds)
            lst_html.append (to_html(text))
            self.line_count = y
        html = ''.join(lst_html)
        if not html:
            return
        deferred = self.callRemote ('cp_output', 0, 0, unicode(html))
        count = 1000
        while count > 0:
            if hasattr(deferred, 'result'):
                break
            time.sleep (0.1)  # ugly
            count -= 1

def to_html (text):
    lst_result = []
    for c in text:
        if c == ' ': lst_result.append ('&nbsp;')
        elif c == '&': lst_result.append ('&amp;')
        elif c == '<': lst_result.append ('&lt;')
        elif c == '>': lst_result.append ('&gt;')
        elif c == '"': lst_result.append ('&quot;')
        else: lst_result.append (c)
    return ''.join (lst_result)

#----------------------------------------------------------------------

page_html = '''\
<html
    xmlns:nevow="http://nevow.com/ns/nevow/0.1"
    xmlns:athena="http://divmod.org/ns/athena/0.7">
<head>
    <nevow:invisible nevow:render="liveglue"/>
    <link rel="stylesheet" href="css/pyconsole_nevow.css"/>
</head>
<body>
<athena:handler event="onload" handler="onload"/>
<div nevow:render="console_process"/>
</body>
</html>
'''

class ConsoleProcessPage (athena.LivePage):
    addSlash = True
    docFactory = loaders.xmlstr(page_html)
    child_css = static.File (os.path.dirname (__file__))

    def __init__ (self, *a, **kw):
        super (ConsoleProcessPage, self).__init__ (*a, **kw)
        self.jsModules.mapping ['PyConsole'] = util.sibpath (__file__, 'pyconsole_nevow.js')
        self.initialized = False

    def render_console_process (self, ctx, data):
        self.console_process_nevow = ConsoleProcessNevow ()
        self.console_process_nevow.page = self
        return ctx.tag [self.console_process_nevow]

#----------------------------------------------------------------------

index_html = '''<html><body>
<a href="console_process_page">Console Process Page</a><br/>
</body></html>
'''
class Index (rend.Page):
    addSlash = True
    docFactory = loaders.xmlstr(index_html)
    child_css = static.File ('.')

    def child_console_process_page (self, ctx):
        return ConsoleProcessPage ()

#----------------------------------------------------------------------

from twisted.application import service, strports
from twisted.python import log, logfile
from nevow import appserver

log.FileLogObserver.timeFormat = '%H:%M:%S'

application = service.Application ('trt_main')
twisted_service = strports.service ('8080', appserver.NevowSite (Index()))
twisted_service.setServiceParent (application)

