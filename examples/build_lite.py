import os
from zipfile import ZipFile

this_dir = os.path.dirname(os.path.abspath(__file__))
if not os.path.exists(os.path.join(this_dir, '_build')):
    os.makedirs(os.path.join(this_dir, '_build'))

for this_dir, dirs, files in os.walk(this_dir):
    if '_build' in dirs:
        dirs.remove('_build')
    for d in dirs:
        with ZipFile(os.path.join('_build', d + '.zip'), 'w') as zf:
            zf.write(os.path.join(this_dir, d, 'LICENSE.txt'), 'LICENSE.txt')
            zf.write(os.path.join(this_dir, d, d + '.py'), d + '.py')
            zf.write(os.path.join(this_dir, d, d + '.xlsm'), d + '.xlsm')
            if d == 'database':
                zf.write(os.path.join(this_dir, d, 'chinook.sqlite'), 'chinook.sqlite')

