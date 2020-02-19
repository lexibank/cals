from setuptools import setup
import json


with open("metadata.json", encoding="utf-8") as fp:
    metadata = json.load(fp)


setup(
    name="lexibank_cals",
    description=metadata["title"],
    license=metadata.get("license", ""),
    url=metadata.get("url", ""),
    py_modules=["lexibank_cals"],
    include_package_data=True,
    zip_safe=False,
    entry_points={"lexibank.dataset": ["cals=lexibank_cals:Dataset"]},
    install_requires=["pylexibank>=2.1.0", "python-docx"],
    extras_require={"test": ["pytest-cldf"]},
)
