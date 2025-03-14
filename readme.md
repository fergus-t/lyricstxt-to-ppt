# lyricstxt-to-ppt

A Python tool to automate the task of reading a txt file line by line, and create a white on black text powerpoint presentation, for use as a teleprompter for singers during performance.

## Usage

` python main.py input_file_name.txt `

The program will automatically output a .pptx file with the same file name as the input .txt file. 

Optionally, you may specify the font size, and whether or not to preserve blank lines in the output powerpoint file. Blank lines in the .txt file will show up as a blank slide in powerpoint. Multiple consecutive blank lines will be merged into a single blank line.

The command below takes the file input_file_name.txt, and outputs with a font size of 80, and keeping blank lines in the output powerpoint slide.

` python main.py input_file_name.txt --font_size 80 --preserve_newline`



## Installation

Tested on Python 3.10. 

Ensure that python and pip is installed. 

Then, run `pip install -r requirements.txt` in the project directory.
