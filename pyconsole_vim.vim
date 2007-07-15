" PyConsole project
" Copyright (C) 2007 Michael Graz
"
" This library is free software; you can redistribute it and/or
" modify it under the terms of the GNU Lesser General Public
" License as published by the Free Software Foundation; either
" version 2.1 of the License, or (at your option) any later version.
"
" This library is distributed in the hope that it will be useful,
" but WITHOUT ANY WARRANTY; without even the implied warranty of
" MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
" Lesser General Public License for more details.
"
" You should have received a copy of the GNU Lesser General Public
" License along with this library; if not, write to the Free Software
" Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
"
" To run from cmd line:
"   gvim "+source pyconsole_vim.vim"
" or else within vim:
"   :source pyconsole_vim.vim

set swapsync=
set updatetime=200
set nocursorline
let g:console_process_row = -1
let g:console_process_row_last = -2

au CursorHold <buffer> call CheckUpdated()
function CheckUpdated()
    if g:console_process_row > g:console_process_row_last
        let g:console_process_row_last = g:console_process_row
        " since this a timer based command the :startinsert
        " does not work.  So send in the line append command
        call remote_send (v:servername, 'A')
    endif
endfunction

python import pyconsole_vim
python vc = pyconsole_vim.VimConsole('cmd.exe')

imap <buffer> <cr> <esc>:python vc.exec_line()<cr>
imap <buffer> <tab> <esc>:python vc.exec_part()<cr>
