from setuptools import setup, find_packages

with open("README.md", "r") as fh:
    long_description = fh.read()
VERSION = '0.0.1'

setup(
    name='FuckExcel',
    version=VERSION,
    description='Just Fuck Excel!',
    long_description=long_description,
    long_description_content_type="text/markdown",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    keywords='Excel',
    author='RhythmLian',
    url="https://github.com/Rhythmicc/qs",
    license='MIT',
    packages=find_packages(),
    include_package_data=True,
    zip_safe=True,
    install_requires=['openpyxl']
)