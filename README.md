# Presentation Grading Script
## Abstract
This script was my attempt at making grading presentation scores much easier. At the end of the day, this task is repetitive by hand and it is completely comprised of data. We're (probably) computer scientists and we're (probably) better than this. So, I wrote this script to help both me and you save some time in the long wrong and make this process much less time and soul consuming.

There is, however, one caveat. While I am quite confident in this script, I wrote it in a day. I do plan on continuing development throughout the other presentations. All that being said, I do not believe that I can guarantee 100% correctness. This lack of confidence mostly stems from the development process, human errors, and my subsequent error-checking, were realized and developed in realtime. That is to say, there was little planning for what kind of errors could pop up. When I first started this, I had no expectation that students would modify the standard sections of the spreadsheet; this turned out to be very, very common, unfortunately. I did compare several handpicked, random spreadsheets to hand grade and test against my script and came up with the same results, but the most I will say is that this script should give you *mostly* correct scores *most* of the time. If it doesn't, that one's on you pal, fix it! Contribute to open source!

## Prerequisites
1. The primary prequisite to run this script is Python3, refer to the internet for instructions on how to install.

2. The only other prerequisite is a Python package called "openpyxl".
My prefered method of installing Python packages is: `python3 -m pip install openpyxl`.
I do this so that I can guarantee the package is installed under the same Python installation I will be using to run the script.

You can also just install it by typing `pip install openpyxl` or `pip3 install openpyxl` but if it doesn't work after you do this, revert to the previous example.

## Running
I'd give installation a dedicated section, but it is as simple as cloning this git repository (or downloading this project from the "Files" section on canvas).

After you "struggle" through this installation procress, you should have an empty directory entitled "submissions" in the root directory. This directory is where you will put, you guessed it, your students' submissions. You should download their submissions using the "Download Submissions" button on the assignment page on Canvas as this will prefix their Excel files with the name; this script uses the file name to check for self-grading errors.

This script specifically selects Excel documents with a ".xlsx" file extension, every other file in this directory will be ignored. This should only benefit you, by ignoring other files accidentally left in this directory, and never hurt you, since every student will be starting from the given rubric spreadsheet. You should still check to make sure nobody exported it into a different file format and account for that.

This script will work automatically and tell you what you should re-check when it finds something it isn't sure of.

`python3 grade.py` will start the process. A trimmed down, more digestible version of this output will be automatically saved in the root directory along with a csv for a more concise viewing experience. These, however, do not contain the extra, verbose, information that is printed to STDOUT. If you would like to save these logs, you should perform a redirect in your shell of choice, most likely: `python3 grade.py > full.log`.

While running, this will print out its steps along the way. This includes telling you what sheets contain negative numbers and are, likely, invalid. It also prints out type mismatches, usually when the user adds an extra cell near the grading portion, and how many perfect grades a student hands out, which you should be suspicious of. The script will also automatically try to assign students, who someone has either put in the wrong group or mispelled their name despite it being pre-typed, to their rightful place. Meaning that, if a person is registered as two people, it will make one last round to attempt to fix it. It even prints out a handy diagram telling you who was successfully reassigned (just in case it wasn't as successful as it thinks) and who was left unassigned.
