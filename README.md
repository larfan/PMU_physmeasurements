# PMU_physmeasurements
This repository contains python code, which automatically transfers data from 3 different measuring device's files into separate .xlsx files. Afterwards one final script, transfers all the averages of the day to a collective .xlsx file, that links to the other three files of the day.

# Filepaths
The filepaths most of the time need to be modified at the start of the script. While writing into the files it doesn't really matter where they are located. However the only important step is in the end with the averageall.py file, because here the excel formula is specifically referering to a directory on a usb stick:<br/>
    &nbsp;filesdirectory='=MITTELWERT(\'F:\\allaverage\\{}'.format(specfolder)<br/><br/>
As you can see the mounted usb stick needs to be F:. This is the case, because I am running the scripts on a linux machine and then transferring all the files onto the stick. Then I open the gesmessungenhauptblatt.xlsx on a windows machine, when it's still on the stick and copy the values to the final file. 
