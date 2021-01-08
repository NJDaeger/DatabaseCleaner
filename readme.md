To get this running, you will need to have maven installed. To install maven, follow the instructions on the link below.
https://maven.apache.org/install.html

To run this application, after you install maven fully, I recommend loading the project inside of your IDE. Most modern IDEs support running a maven project. I recommend IntelliJ, but others will work as well.

To actually clean the database, put the file CMIC_DATA.accdb inside of the same directory where the src folder is (dont put it inside of the src folder). Running the program after doing that will generate 5 different excel spreadsheets.

PC_Database is the list of all PCs that have a PC number at central.
Hardware_Database is the list of all hardware at central. (monitors, TVs, cameras, etc)
Model_Database is a map of hardware types to hardware.
Room_Database is a map of all currently saved rooms at all central locations
Phone_Database is a list of all phones out currently at central