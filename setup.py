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


setuptools.setup(
    name="xlcalculator-numpy2",
    version='0.5.1.numpy2.dev1',
    author="Bradley van Ree (Original), LeanSight (numpy 2.0 fork)",
    author_email="brads@bradbase.net",
    description="Converts MS Excel formulas to Python and evaluates them. [numpy 2.0 compatible fork]",
    long_description=(
        "# xlcalculator (numpy 2.0 compatible fork)\n\n"
        "This is an unofficial fork of xlcalculator that adds compatibility with numpy 2.0+ and Python 3.13.\n\n"
        "## Changes from Original:\n"
        "- ✅ Compatible with numpy 2.0+ (was restricted to <2.0)\n"
        "- ✅ Requires Python 3.13 (validated platform)\n"
        "- ✅ Uses numpy 2.0 compatible yearfrac fork\n"
        "- ✅ All tests pass with numpy 2.3.2\n\n"
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
        'numpy>=2.1.0',
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
            'git+https://github.com/LeanSight/yearfrac.git',  # LeanSight numpy 2.0 compatible fork
        ],
    ),
    python_requires='>=3.13',  # Only validated on Python 3.13
    tests_require=TESTS_REQUIRE,
    include_package_data=True,
    zip_safe=False,
)