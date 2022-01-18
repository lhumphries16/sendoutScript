<div id="top"></div>
<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Don't forget to give the project a star!
*** Thanks again! Now go create something AMAZING! :D
-->

<!-- ABOUT THE PROJECT -->
## About The Project

Started this during the holidays 2021 becuase my significant other thought of a way that she could be more efficient in her role at work and wondered if I could try and make it for her. Sounded like a fun project. The role requires her to fill out "sendout request" PDF templates and to enter the same information into a spreadsheet row on a shared file.

So far the tool can:
* Read form fields from the PDF template input
* Generate form inputs and labels dynamically for each field on the GUI
* Generate and store output PDF in /Filled_Requests
* Copy excel row of data using pandas.

Right now there are a lot of hardcoded aspects of the code and I would like to get that down in the future. I am going to continue to look into ways to further automate her workflow. Currently having some problems with cross compatability between Mac and PC that I will also be fixing soon.


### Built With


* [pandas](https://pandas.pydata.org/)
* [pdfrw](https://github.com/pmaupin/pdfrw)
* [tkinter](https://docs.python.org/3/library/tkinter.html)
