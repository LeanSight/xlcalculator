import os
import setuptools


def read(*rnames):
    return open(os.path.join(os.path.dirname(__file__), *rnames)).read()


TESTS_REQUIRE = [
    'coverage',
    'flake8',
    'mock',
    'pytest',
    'pytest-cov',
]

# Optional dependencies for Excel file generation
EXCEL_GENERATION_REQUIRE = [
    'xlwings; platform_system=="Windows"',  # Windows-only for Excel integration
]


setuptools.setup(
    name="xlcalculator-numpy2",
    version='0.5.2.post0+numpy1.24-2.x.python312',
    author="Bradley van Ree (Original), LeanSight (numpy 2.0 fork)",
    author_email="brads@bradbase.net",
    description="Converts MS Excel formulas to Python and evaluates them. [numpy 2.0 compatible fork]",
    long_description=(
        "# xlcalculator (numpy 2.0 compatible fork)\n\n"
        "This is an unofficial fork of xlcalculator that adds compatibility with modern NumPy and Python versions.\n\n"
        "## Changes from Original:\n"
        "- ✅ Compatible with NumPy 1.24+ and 2.x (tested with 1.26.4 and 2.3.3)\n"
        "- ✅ Requires Python 3.12+ (validated on 3.12.1)\n"
        "- ✅ Uses NumPy-compatible yearfrac fork\n"
        "- ✅ All tests pass with NumPy 1.26.4 and 2.3.3\n\n"
        "Original repository: https://github.com/bradbase/xlcalculator\n"
        "This fork: https://github.com/LeanSight/xlcalculator\n\n"
        "---\n\n"
        + read('README.rst')
        + '\n\n' +
        read('CHANGES.rst')
        ),
    long_description_content_type='text/markdown',
    url="https://github.com/LeanSight/xlcalculator",
    packages=setuptools.find_packages(),
    license="MIT",
    keywords=['xls',
        'Excel',
        'spreadsheet',
        'workbook',
        'data analysis',
        'analysis'
        'reading excel',
        'excel formula',
        'excel formulas',
        'excel equations',
        'excel equation',
        'formula',
        'formulas',
        'equation',
        'equations',
        'timeseries',
        'time series',
        'research',
        'scenario analysis',
        'scenario',
        'modelling',
        'model',
        'unit testing',
        'testing',
        'audit',
        'calculation',
        'evaluation',
        'data science',
        'openpyxl',
        'numpy2',
        'fork'
    ],
    classifiers=[
        "Development Status :: 4 - Beta",  # Changed from Production/Stable as this is a fork
        "Intended Audience :: Developers",
        "Intended Audience :: Science/Research",
        "Intended Audience :: Information Technology",
        "Intended Audience :: Financial and Insurance Industry",
        "License :: OSI Approved :: MIT License",
        "Natural Language :: English",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.13",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Scientific/Engineering",
        "Topic :: Scientific/Engineering :: Information Analysis",
        "Topic :: Scientific/Engineering :: Mathematics",
        "Topic :: Software Development :: Libraries",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Software Development :: Testing",
        "Topic :: Software Development :: Testing :: Unit",
        "Topic :: Utilities",
    ],
    install_requires=[
        'jsonpickle',
        'numpy>=1.24.0',
        'pandas>=2.3.0',
        'openpyxl',
        'numpy-financial',
        'mock',
        'scipy>=1.14.1'
    ],
    extras_require=dict(
        test=TESTS_REQUIRE,
        build=[
            'pip-tools',
        ],
        excel_functions=[
            'yearfrac @ git+https://github.com/LeanSight/yearfrac.git',
        ],
        excel_generation=EXCEL_GENERATION_REQUIRE,
    ),
    python_requires='>=3.12',  # Validated on 3.12.1, compatible with 3.13+
    tests_require=TESTS_REQUIRE,
    include_package_data=True,
    zip_safe=False,
)