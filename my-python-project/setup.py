from setuptools import setup, find_packages

setup(
    name="empral-system",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pillow",
        "mysql-connector-python",
    ],
)