from setuptools import setup

setup(
    name="pycalc",
    version="0.2",
    description="Python wrapper of the Rust calc crate",
    author="Nicolás Hatcher",
    author_email="nicolas.hatcher@equalto.com",
    packages=["pycalc"],
    include_package_data=True,
    zip_safe=False,
)
