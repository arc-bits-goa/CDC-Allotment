# CDC-Registration-20-21

This Project was completed as a part of the AUGSD Software team of BIRLA INSTITUTE OF TECHNOLOGY & SCIENCE PILANI, K.K. BIRLA GOA CAMPUS

Python scripts have been used for allotting Course sections to Students based on their preferences and their allotted Priority (PR) Numbers. 

This project was necessary as there were frequent clashes with the sections of CDCs (Compulsory Disciplinary Courses) with Lab sections and other courses' sections. 

Students of CS, EEE, ECE and ENI were provided with several options. Their preferences were gathered and the resultant excel sheets' template is in the folder "Responses".

The various class*.xlsx files have the mapping of a particular option to the class numbers, class times and class codes.

combination.xlsx contains the maximum capacity of each option.

GOA PR.xlsx contains the allotted Priority numbers of each student. 

Student_ID_correctness_check.py checks the correctness of the ID number entered by the student during preference-filling.

capacity_check.py allows to check the capacity of the options such that each student who has to complete a CDC gets a section allotted.

Lab_allotment.py, Lecture_allotment.py and Lecture_EMT_ED_allotment.py contain the main code which allots each student with the most suitable option. The factors considered were the student's priority number, the student's preferences and the maximum number of students who can be allotted to each "option". 
On running the .py files, the allotments are generated as excel files.