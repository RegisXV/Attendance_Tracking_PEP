from setuptools import setup, find_packages

setup(
    name='attendance_tracker',
    version='1.0',
    packages=find_packages(),
    install_requires=[
        'openpyxl',
    ],
    entry_points={
        'console_scripts': [
            'attendance_tracker=main:main',
        ],
    },
)