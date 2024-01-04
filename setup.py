from setuptools import setup, find_packages


def readme():
  with open('README.md', 'r', encoding="utf-8") as f:
    return f.read()


setup(
  name='KOMPAS_tools',
  version='0.2',
  author='Mikarisar',
  author_email='mikarisar@gmail.com',
  description='This library contains tools for KOMPAS 3D automation',
  long_description=readme(),
  long_description_content_type='text/markdown',
  url='https://github.com/Mikarisar/KOMPAS_tools',
  packages=find_packages(),
  install_requires=['pywin32', 'numpy'],
  classifiers=[
    'Programming Language :: Python :: 3.11',
    'License :: OSI Approved :: MIT License',
    'Operating System :: OS Independent'
  ],
  keywords='cad drawing kompas kompas3d automation engineering',
  project_urls={
    'GitHub': 'https://github.com/Mikarisar/KOMPAS_tools'
  },
  python_requires='>=3.6'
)
