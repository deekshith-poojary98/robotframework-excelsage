from setuptools import setup, find_packages


with open("README.md", "r") as f:
    long_description = f.read()

setup(
    name="robotframework-excelsage",
    version="1.1.2",
    description="A package for integrating Robot Framework with Excel using openpyxl and pandas",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/deekshith-poojary98/robotframework-excelsage",
    author="Deekshith Poojary",
    maintainer="Deekshith Poojary",
    author_email="deekshithpoojary355@gmail.com",
    license="Apache License 2.0",
    packages=find_packages(),
    python_requires=">=3.8",
    keywords="excel testing testautomation robotframework robotframework-excelsage robotframework-excellibrary robotframework-excellib",
    install_requires=[
        "pyarrow",
        "pandas>=2.2.0",
        "openpyxl>=3.1.0",
        "robotframework>=5.0.1",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: Apache Software License",
        "Operating System :: OS Independent",
        "Framework :: Robot Framework :: Library",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Testing",
    ],
)
