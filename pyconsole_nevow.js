// Copyright (C) 2007 Michael Graz

if (typeof PyConsole == 'undefined') { var PyConsole = {}; }

PyConsole.Nevow = Nevow.Athena.Widget.subclass('PyConsole.Nevow');

PyConsole.Nevow.methods(

    function cp_output(self, x, y, text) {
        var cp_output = document.getElementById('cp_output');
        cp_output.innerHTML += text;
        window.scrollTo (0, 100000000);   // shortcut to scroll to bottom
    }

);

