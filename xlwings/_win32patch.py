"""See https://github.com/xlwings/xlwings/issues/1789
Can be removed again if there's a solution for
https://github.com/mhammond/pywin32/issues/1870

This file's content is taken from pywin32 v301,
distributed under the following license:

Unless stated in the specfic source file, this work is
Copyright (c) 1996-2008, Greg Stein and Mark Hammond.
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions
are met:

Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

Redistributions in binary form must reproduce the above copyright
notice, this list of conditions and the following disclaimer in
the documentation and/or other materials provided with the distribution.

Neither names of Greg Stein, Mark Hammond nor the name of contributors may be used
to endorse or promote products derived from this software without
specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ``AS
IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE REGENTS OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

import pythoncom


class CoClassBaseClass:
    def __init__(self, oobj=None):
        if oobj is None:
            oobj = pythoncom.new(self.CLSID)
        self.__dict__["_dispobj_"] = self.default_interface(oobj)

    def __repr__(self):
        return "<win32com.gen_py.%s.%s>" % (__doc__, self.__class__.__name__)

    def __getattr__(self, attr):
        d = self.__dict__["_dispobj_"]
        if d is not None:
            return getattr(d, attr)
        raise AttributeError(attr)

    def __setattr__(self, attr, value):
        if attr in self.__dict__:
            self.__dict__[attr] = value
            return
        try:
            d = self.__dict__["_dispobj_"]
            if d is not None:
                d.__setattr__(attr, value)
                return
        except AttributeError:
            pass
        self.__dict__[attr] = value

    # Special methods don't use __getattr__ etc, so explicitly delegate here.
    # Some wrapped objects might not have them, but that's OK - the attribute
    # error can just bubble up.
    def __call__(self, *args, **kwargs):
        return self.__dict__["_dispobj_"].__call__(*args, **kwargs)

    def __str__(self, *args):
        return self.__dict__["_dispobj_"].__str__(*args)

    def __int__(self, *args):
        return self.__dict__["_dispobj_"].__int__(*args)

    def __iter__(self):
        return self.__dict__["_dispobj_"].__iter__()

    def __len__(self):
        return self.__dict__["_dispobj_"].__len__()

    def __nonzero__(self):
        return self.__dict__["_dispobj_"].__nonzero__()
