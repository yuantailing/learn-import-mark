# -*- coding: utf-8 -*-

from __future__ import absolute_import
from __future__ import division
from __future__ import print_function
from __future__ import unicode_literals

import py2exe
import six

from distutils.core import setup


assert six.PY2

setup(
    name='learn-import-mark'.encode(),
    windows=[{'script': 'src/main.py'}],
    options={
        'py2exe': {
            'includes': [
                'Queue',
                'Tkinter',
                'cookielib',
                'tkFont',
                'tkMessageBox',
                'tkFileDialog',
                'ttk',
                'urllib2',
            ]
        }
    },
)
