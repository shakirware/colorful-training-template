from setuptools import find_packages, setup

setup(
    name="colorful_training_template",
    version="0.1.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "openpyxl==3.1.5",
        "pandas==2.2.3",
        "PyYAML==6.0.2",
    ],
    python_requires=">=3.7",
    entry_points={
        "console_scripts": [
            "colorful_training_template=colorful_training_template.main:main"
        ]
    },
)
