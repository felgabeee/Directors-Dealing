
# Directors' dealings
The whole purpose of this Python project is to create a Graphical User Interface (GUI) able to extract and store (in a a excel file) the relevant information from the Director's dealings PDF files.
All files are publicly available and can be found here: https://bdif.amf-france.org/fr?typesInformation=DD

Note: directors' dealings refer to transactions with securities or related derivative financial instruments carried out by individuals who perform some executive function within the company.

# Overview/Start guide
The final UX output should looks like that:


![alt text](https://github.com/felgabeee/Directors_Dealing/blob/main/images/AMF_UX.PNG?raw=true)

The main class DD has three main instances that work in the following order:

* The instance get_DD() will download every directors' dealings pdf file from the AMF/ESMA website in a given period of time (start_date up to end_date) and store these files in a directory named "FR_DD" (it will be created by default if you are running the code for the first time).

* The instance extract_DD() will loop through each pdf file previously downloaded and extract the relevent information using [regex](https://fr.wikipedia.org/wiki/Expression_r%C3%A9guli%C3%A8re).
The outup of that class instance is an excel file: Directos_Dealing _extract.xlsx (again, it will be created by default if you are running the code for the first time).

* Lastly, the UX() instance will create the GUI application. You can specify a search word (for instance the name of a publicly tradded company), a start and end date and a desired path for the excel file.

# Common issues and solutions

* The whole class was built using Selenium to interact with the website HTML elements => if the website changes so do the HTML elements and the code will return an error. In that case simply change Xpath within the code.

* The path of the webdriver should be located in you current working directory or specificed with an r in front of the path like in this example: r"C:\Users\ernes\OneDrive\Bureau\chromedriver.exe"

* If the access to the excel fie is denied then you probably have it opened => close it.

* If the language you are using on your computer is different from FR or ENG then you will need to change the __init__ instance accordingly.


