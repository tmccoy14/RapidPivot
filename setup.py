from setuptools import setup, find_packages

setup(
    name="rpcli",
    version="0.1",
    description="rpcli is a tool to help sort and format excel spreadsheets.",
    author="Tucker McCoy",
    author_email="tuckermmccoy@gmail.com",
    keywords="excel automation python",
    packages=find_packages(exclude=["tests"]),
    install_requires=[
        "Click==7.0",
        "colorama==0.4.3",
        "tabulate==0.8.7",
        "xlwt==1.3.0",
    ],
    entry_points={"console_scripts": ["rpcli=src.main:cli"]},
)
