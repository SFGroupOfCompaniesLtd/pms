from setuptools import setup, find_packages

with open("requirements.txt") as f:
	install_requires = f.read().strip().split("\n")

# get version from __version__ variable in property_management/__init__.py
from property_management import __version__ as version

setup(
	name="property_management",
	version=version,
	description="Property Management system for Real Estate Vertical",
	author="Nihantra C. Patel",
	author_email="contact@solufy.in",
	packages=find_packages(),
	zip_safe=False,
	include_package_data=True,
	install_requires=install_requires
)
