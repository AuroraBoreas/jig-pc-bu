import setuptools

setuptools.setup(
    name='jigpc-auto-pkg',
    version='0.0.4',
    author='ZL',
    author_email='liang.zhang@sony.com',
    description='',
    long_description='',
    license='MIT',
    packages=setuptools.find_namespace_packages(include=['JigPC_pkgs.*']),
)