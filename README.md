# Headshot generator for UQ courses.

Generates a list of PDF files containing the headshots and names of a course, one PDF per tutorial/workshop

This takes an enrolment spreadsheet downloaded from Allocate+ and a zip file containing headshots,
and for each tutorial/workshop/studio, it creates a PDF file with the headshots and names of students
 enrolled. 

Steps to use:
1. Download an xls of your class from Allocate+ (https://timetable.my.uq.edu.au/odd/admin). Select only the tutorials/workshops/studios/whatever, and then select 'Class List'.
2. Download the headshots for your couse from my.uq.edu.au -> 'My Courses' -> Select course -> 'In Person', and 'Student headshots'. Download the zip file
3. Run this script using:  python3 headshots.py course_classlist.xls headshots_zip_file.zip

## Example

Run the example class list using:

> python3 headshots.py Sample_classlist.xls motown_legends.zip
