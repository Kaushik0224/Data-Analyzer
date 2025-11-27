from setuptools import setup, find_packages

setup(
    name="excel-analyzer-pro",
    version="2.0.0",
    description="Professional Excel Data Analyzer with Modern UI",
    author="Excel Analyzer Team",
    packages=find_packages(),
    python_requires=">=3.8",
    install_requires=[
        "streamlit==1.28.1",
        "pandas==2.1.0",
        "openpyxl==3.11.0",
        "plotly==5.17.0",
        "reportlab==4.0.7",
    ],
)
