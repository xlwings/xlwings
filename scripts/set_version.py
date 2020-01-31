import os

version_file = 'version'
print('GITHUB_REF: ' + os.environ['GITHUB_REF'])
if os.environ['GITHUB_REF'].startswith('refs/tags'):
    version_string = os.environ['GITHUB_REF'][10:]
else:
    version_string = os.environ['GITHUB_SHA'][:7]
with open(version_file, 'w') as f:
    f.write(version_string)
