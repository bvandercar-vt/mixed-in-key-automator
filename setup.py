from setuptools import setup

setup(
   name='mixed_in_key_automator',
   version='1.2.0',
   description='Windows automator for Mixed In Key. Known works for MIK versions 10 and 11',
   author='Blake Vandercar',
   author_email='TODO',
   packages=["mik_automator"],   
   package_dir={"mik_automator": "src"},
   install_requires=['pywinauto']
)