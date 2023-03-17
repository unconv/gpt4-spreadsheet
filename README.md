# Excel Math Test by GPT-4

The code in this repository was created by the GPT-4 model of ChatGPT based on a prompt to create an Excel Math Test for 5th graders.

Watch the video where I made it: https://www.youtube.com/watch?v=EvA-rhWwd9k

## The Prompt

ChatGPT was given the following prompt:

```
I want to create a math test for 5th grade students in Excel. The test will have 10 questions. Every question should have a cell for points. Every question will get 0-5 points. In the end, the total points should be calculated. Also, the maximum possible points should be calculated. Then from these numbers, a grade should be calculated (A, B, C, D or F). I want you to create this math test including the questions and the formulas for calculating the points and grades. Also add styling to the spreadsheet so that the question is in bold and the answer is not. The cell for the points of each question should have a yellow background. There should also be a place in the top left corner for the name of the student.
```

And then told to create a PHP script that creates the described spreadsheet (in `mathtest.php` and `mathtest2.php`).

The `csv-spreadsheet.csv` was created with the following prompt added to the above prompt:

```
I know you can't create Spreadsheets directly, so let's create our own syntax for styled spreadsheets with formulas that can easily be converted into an XLS file programmatically. It could be similar to XML on CSV but have some sort of syntax for styling and adding formulas.

Please create this spreadsheet format and create the spreadsheet I described with it.
```

It was then asked to create a PHP script that will be able to parse the new CSV format and convert it into an XLS file (`csv-parser.php`)
