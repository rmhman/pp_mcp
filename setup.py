from setuptools import setup, find_packages

setup(
    name="powerpoint-mcp",
    version="0.1.0",
    packages=find_packages(),
    install_requires=["python-pptx", "mcp"],
    entry_points={"console_scripts": ["powerpoint-mcp-server = powerpoint_server:main"]},
)
