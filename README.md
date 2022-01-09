# **stock-analysis with VBA**

## **Overview of Project**
**_Background:_** Doing some more research to expand the dataset to include the entire stock market over the last few years. 

### **Results**
**_Results Before:_** Before refactoring the code the resulting times to execute were 1.015625 seconds for 2017 data and 1.066406 seconds for 2018 data

![Old 2017](https://user-images.githubusercontent.com/96216509/148696891-83e443b1-84a7-4481-acb7-bf5c07f7c1cf.png)
![old 2018](https://user-images.githubusercontent.com/96216509/148696901-19012310-8979-4b42-8cc5-ba8451d6deef.png)


**_Results After:_** After refactoring the code the resulting times to execute were 0.25 seconds for 2017 data and 0.2148438 seconds for 2018 data

![VBA_Challenge_2017](https://user-images.githubusercontent.com/96216509/148696938-8f87accb-d8cf-47e7-8bfe-08e16af67450.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/96216509/148696956-b2d20d9a-aea8-46e1-9c06-39cacfa6dd27.png)


**_Conclusion:_** Refactoring the code successfully made the VBA script run faster.

## Summary

### What are the advantages or disadvantages of refactoring code?   

#### Some advantages:
- Can make the code more efficient—by taking fewer steps
- Can use less memory
- Can improve the logic of the code to make it easier for future users to read

#### Some disadvantages:
- Can be risky when the application is big by introducing bugs
- Can be risky when existing code doesn't have proper test cases to find introduced bugs
- Can be risky when developers do not understand the original code
- Can takes extra time and budget that might not be available

### How do these pros and cons apply to refactoring the original VBA script?

Refactoring the original VBA script resulted in a more efficient program that resulted in faster run times. In this exercise since the data set was not too large the performance increase was minimal and when wieghed against the extra time if took to refactor the original code not of much value. Hours of time were spent to save a fraction of a second in exectution time. If the data set were much larger and executed many times only then might the refactoring have been worth the time to do.  
