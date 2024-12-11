# Instructify
#### Disclaimer: This language is for a school project and is not yet a complete language. Might be added to at a later date.

This is Instructify, a language meant for use by educators to grade tests, create lesson templates, and grade distributions (and possibly more in the future). 

#### How to Run (as of 12/11/2024)
- Import necessary libraries to python using pip install
   - python pptx
   - python docx
   - python textX
- Go into the src folder in the repository and download it
- In the folder are: 
   - example Instructify programs docTemplate, gradeDistribution, lessonTemplate, and gradeTest
   
   - instructify.py (the interpreter)

   - instructify.tx(textX file)

   - example images for the lessonTemplate program
      - northeastus.jpg
      - pacificus.jpg
      - southernus.jpg
      - usmap.jpg

- open the src folder in your IDE of choice
- To run a particular program change the index inside the brackets for programs[] on line 459, instructify.py
- to make your own programs simply create a .instruct file inside the same directory of the other files and start coding
- if you want to place images into a slide, follow the example on lessonTemplate, and make sure it is in the same directory as the other files
- if you want to use the gradeTest template, make sure to replace the example email address with your own

#### Output:
- For lessonTemplate.instruct, a new pptx file will appear in the src directory
- For docTemplate.instruct, a new docx file will appear in the directory
- For gradeTest.instruct, a text file will appear in the directory and an email sent to a recipient if specified in the code
- For gradeDistribution.instruct, jpg files will appear in the directory
