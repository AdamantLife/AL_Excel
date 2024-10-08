import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="AL_Excel",
    version="1.0.3",
    author="AdamantLife",
    author_email="",
    description="A collection of code snippets and high-level interfaces",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/AdamantLife/AL_Excel",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License (GPL)",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.8',
    install_requires=[
        "openpyxl",
        ]
)
