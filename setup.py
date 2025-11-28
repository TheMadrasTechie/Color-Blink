from setuptools import setup, find_packages

with open("README.md", "r") as f:
    description = f.read()


setup(
name='color_blink',
version='1.0.1',
packages=find_packages(),
install_requires=[
    'opencv-python-headless==4.12.0.88',
    'argparse==1.4.0',
    'xlwings==0.33.16',
    'pillow==11.3.0'
],
entry_points={
"console_scripts": [
"color_blink-reduce-image = color_blink:reduce_image_size_argparse",
"color_blink-image-to-excel = color_blink:image_to_excel_argparse",
"color_blink-blink-generate = color_blink:blink_excel_cell_argparse",
"color_blink-test = color_blink:work",
],},
long_description=description,
long_description_content_type="text/markdown",
)