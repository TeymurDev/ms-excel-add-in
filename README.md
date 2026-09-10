# It took me 3 hours and 34 minutes to do the task.

## I will explain you step by step how I did this task.

I tried to do this task with ChatGPT and Cursor but as you know that their models are not strong enough I decided to use Claude:

1. I asked Claude to implement this task below:
2. 
Develop an Excel add-in that:

Uses the Office.js framework

Uses the shared runtime

Uses Typescript

Adds a new custom function, FACTORIALROW, to Microsoft Excel.

The function should be in the TESTVELIXO namespace.

The function accepts a single numeric value N and returns a row of calculated factorials:

[0!, 1!, 2!, 3!, ..., N!]

Important: It should not be a comma-separated list of values inside one cell, but a spill range of N cells.

2. I received an Error regarding custom functions and fixed it. the issue was I forget to put code below on the top of functions.ts.
/// <reference types="office-js" />
/// <reference types="office-runtime" />
3. I asked allow localhost loopback for microsoft edge  webview. Shall I do it? Anser was yes.
4. I receive this result when paste =TESTVELIXO.FACTORIALROW(5) to excel worksheet
   <img width="134" height="46" alt="image" src="https://github.com/user-attachments/assets/5e9c7cc7-4618-4335-911a-dea61fe4870b" />

5. <img width="349" height="143" alt="image" src="https://github.com/user-attachments/assets/fba606e5-7c9a-4b3c-974e-0fd6dab47856" />
6. <img width="938" height="204" alt="image" src="https://github.com/user-attachments/assets/e0a95333-0e28-43b3-9e5d-b0726a6be300" />
7. <img width="1013" height="135" alt="image" src="https://github.com/user-attachments/assets/dcd638d1-da38-43de-a75b-d495647c2227" />
8. <img width="763" height="334" alt="image" src="https://github.com/user-attachments/assets/1c05bfc5-c771-4299-a042-4bc9042377f6" />
9. I had an issue that function generated result 1
    <img width="203" height="82" alt="image" src="https://github.com/user-attachments/assets/1c5c00b9-e105-49e1-83d6-400749031c04" />

10. the first thing that you have mentioned is ok

   "result": {
                "dimensionality": "matrix",
                "type": "number"
            }
but dynamic-array spilling still shows 1

when I type inside cell ={1,2,3,4,5}

11. Finally I had an issue that my version of ms office doesn't support dynamic-array spilling. But ms 365 supports it. I tried to run ms 365 online but microsoft for no reason blocked my account creation 😔

    Task is ready and you can check it. Looking forward to hearing from you. Thank you!





