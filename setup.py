from setuptools import setup, find_packages


def readme():
  with open('README.md', 'r') as f:
    return f.read()


setup(
  name='KOMPAS Tools',
  version='0.1',
  author='Maksim Pinchuk',
  author_email='mikarisar@gmail.com',
  description='This library contains tools for KOMPAS 3D automation',
  long_description=readme(),
  long_description_content_type='text/markdown',
  url='https://github.com/NayadaEngineer/KOMPAS_tools',
  packages=find_packages(),
  install_requires=['pythoncom, win32, datetime, numpy'],
  classifiers=[
    'Programming Language :: Python :: 3.11.2',
    'License :: OSI Approved :: MIT License',
    'Operating System :: Microsoft Windows 10 Pro 22H2'
  ],
  keywords='cad drawing kompas kompas3d automation engineering',
  project_urls={
    'GitHub': 'https://github.com/NayadaEngineer/KOMPAS_tools'
  },
  python_requires='>=3.6'
)
