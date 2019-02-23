# Release process

1. If the REST API docs changed: Run `/scripts/build_rest_api_docs.py` and commit the result.
1. Add release notes to `docs/whatsnew.rst`
2. [Create a new release](https://github.com/ZoomerAnalytics/xlwings/releases/new) with tag and name (no leading `v`): `x.x.x`

   This kicks off the appveyor pipeline that will:
   
   * Build the dlls, the `xlwings.xlam`, the standalone files and `xlwings.bas` (they depend on the version number) and the Python package
   * Upload the Python package to pypi and `xlwings.xlam` to the GitHub release page
   * Trigger a rebuild of https://www.xlwings.org so it is updated with latest version/date

3. readthedocs.org should also trigger a new build automatically (login with GH account)
4. The [conda-forge](https://github.com/conda-forge/xlwings-feedstock) package seems to automatically
   create a PR these days shortly after uploading the package to pypi, so the only tasks left here is to check back after
   a few hours and merge it if all build pipelines built successfully.
