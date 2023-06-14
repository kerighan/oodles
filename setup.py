import setuptools


setuptools.setup(
    name="oodles",
    version="0.0.1",
    author="Maixent Chenebaux",
    author_email="max.chbx@gmail.com",
    description="Easiest library to manipulate Google Slides and Sheets",
    url="https://github.com/kerighan/oodles",
    packages=setuptools.find_packages(),
    include_package_data=True,
    install_requires=["addict==2.4.0", "google-api-python-client", "google-auth"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.5"
)
