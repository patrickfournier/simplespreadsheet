from distutils.core import setup
import rst_simplespreadsheet

setup(name='rst_simplespreadsheet',
      version=rst_simplespreadsheet.__version__,
      author=rst_simplespreadsheet.__author__,
      author_email='patrick@whiskyechobravo.com',
      description=('Simple spreadsheet extension for reStructuredText tables.'),
      url='https://github.com/patrickfournier/simplespreadsheet',
      long_description=rst_simplespreadsheet.__doc__,
      license=rst_simplespreadsheet.__license__,
      py_modules=['rst_simplespreadsheet'],
      classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Science/Research",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Topic :: Documentation",
        "Topic :: Software Development :: Documentation",
        "Topic :: Utilities",
        ],
      )
