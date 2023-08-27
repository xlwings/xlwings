# Release process

1. Run `python xlwingsjs/build.py --version x.x.x` and commit the result.
2. Add release notes to `docs/whatsnew.rst`
3. [Create a new release](https://github.com/ZoomerAnalytics/xlwings/releases/new) with tag and name (no leading `v`): `x.x.x`

   This kicks off the GitHub actions pipeline that will:
   
   * Build the dlls, the `xlwings.xlam`, the standalone files and `xlwings.bas` (they depend on the version number) and the Python package
   * Upload the Python package to PyPI
   * Trigger a rebuild of https://www.xlwings.org so it is updated with latest version

4. readthedocs.org triggers a new build automatically (login with GH account)
5. The [conda-forge](https://github.com/conda-forge/xlwings-feedstock) package automatically
   creates a PR shortly after uploading the package to PyPI, so the only tasks left here is to check back after
   a few hours and merge it if all pipelines built successfully.
