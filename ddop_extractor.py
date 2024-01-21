# import libraries
from xml.dom import minidom
import csv
from datetime import datetime
import os.path
import calendar
import subprocess
import pandas as pd
from inspect import getsourcefile
from os.path import abspath
from tkinter import filedialog
from tkinter import messagebox


def main():
    # this function converts little endian hex to big endian hex
    def little_endian_to_big_endian(argument):
        argument = bytearray.fromhex(argument)
        argument.reverse()
        argument = ''.join(format(n, "02x") for n in argument).upper()
        return argument

    # a function to set the name of the outputted xml file
    # format: {Manufacturer}_{Type}_{Name}_{3-char-month}_{year}.xml
    def set_name_of_file(b, d):
        implement_type = ""
        implement_manufacturer = ""
        implement_name = b
        current_month = calendar.month_abbr[datetime.now().month]  # gets abbreviated month name
        current_year = datetime.now().year
        d = little_endian_to_big_endian(d)

        # dictionary to map vehicle_system_id to vehicle_system_description
        mapping_dict = {0: "Non-specific System", 1: "Tractor", 2: "Tillage", 3: "Secondary Tillage",
                        4: "Planters_Seeders", 5: "Fertilizers",
                        6: "Sprayers", 7: "Harvesters", 8: "Root Harvesters", 9: "Forge", 10: "Irrigation",
                        11: "Transport_Trailer", 12: "Farm Yard Operations", 13: "Powered Auxiliary Devices",
                        14: "Special Crops", 15: "Earth Work",
                        16: "Skidder", 17: "Sensor_Systems", 19: "Timber_Harvesters", 20: "Forwarders",
                        21: "Timber Loaders", 22: "Timber Processing Machine", 23: "Mulchers", 24: "Utility Vehicles",
                        25: "Slurry_Manure Applicators",
                        26: "Feeders_Mixers", 27: "Weeders-Non-chemical weed control", 28: "Turf and Lawn Care Mowers"}

        # get the type
        byte_7 = int(d[12:14], 16) >> 1
        for key in mapping_dict:
            if byte_7 == key:
                implement_type = mapping_dict.get(key).upper()

        # get the manufacturer
        byte_4 = int(d[6:8] + d[4], 16) >> 1
        with open(os.path.join(abspath(getsourcefile(lambda: 0)).strip("ddop_extractor.py"), "Manufacturer IDs.csv"),
                  'r') as csv_file:
            big_list = csv.reader(csv_file)
            for sub_list in big_list:
                if sub_list[0] == str(byte_4):
                    implement_manufacturer = sub_list[1].split()[0].upper()

        # return full name
        return f'{implement_manufacturer}_{implement_type}_{implement_name.upper()}_{current_month}_{current_year}'

    def extract_and_write(parsed_xml_file, dvc_list, directory):
        counter = 0

        # variables to hold parent element attributes
        VersionMajor = ""
        VersionMinor = ""
        ManagementSoftwareManufacturer = ""
        ManagementSoftwareVersion = ""
        TaskControllerManufacturer = ""
        TaskControllerVersion = ""
        DataTransferOrigin = ""
        P094_XML_VERSION = ""
        P094_ADDITIONAL = ""

        # get the parent element ("ISO11783_TASKDATA") and its attributes:
        parent_element_list = parsed_xml_file.getElementsByTagName("ISO11783_TaskData")
        for x in parent_element_list:
            VersionMajor = x.getAttribute("VersionMajor")
            VersionMinor = x.getAttribute("VersionMinor")
            ManagementSoftwareManufacturer = x.getAttribute("ManagementSoftwareManufacturer")
            ManagementSoftwareVersion = x.getAttribute("ManagementSoftwareVersion")
            TaskControllerManufacturer = x.getAttribute("TaskControllerManufacturer")
            TaskControllerVersion = x.getAttribute("TaskControllerVersion")
            DataTransferOrigin = x.getAttribute("DataTransferOrigin")
            P094_XML_VERSION = x.getAttribute("P094_XML_VERSION")
            P094_ADDITIONAL = x.getAttribute("P094_ADDITIONAL")

        parent_element = f'\n<ISO11783_TaskData VersionMajor="{VersionMajor}" VersionMinor="{VersionMinor}"' \
                         f' ManagementSoftwareManufacturer="{ManagementSoftwareManufacturer}"' \
                         f' ManagementSoftwareVersion="{ManagementSoftwareVersion}"' \
                         f' TaskControllerManufacturer="{TaskControllerManufacturer}"' \
                         f' TaskControllerVersion="{TaskControllerVersion}"' \
                         f' DataTransferOrigin="{DataTransferOrigin}"' \
                         f' P094_XML_VERSION="{P094_XML_VERSION}"' \
                         f' P094_ADDITIONAL="{P094_ADDITIONAL}">'

        """
         filter DVC objects to the ones we want. Then for each
         DVC object, extract attributes B and D. 
         Scan the directory and look for all outputted
         CSV files (TLG#.xml). For each file, we want to extract
         initial values from it, add them to each DVC, write 
         those DVCs in a new XML file, and properly name the file. 
        """
        # make sure to grab DVCs with "A" attribute value between 1 and 999999 only
        filtered_dvc_list = [dvc_object for dvc_object in dvc_list if
                             1 <= int(dvc_object.getAttribute("A")[4:]) <= 999999]

        for DVC in filtered_dvc_list:
            counter += 1
            # get attributes "B" and "D" from each DVC object
            name = DVC.getAttribute("B")
            attribute_d = DVC.getAttribute("D")

            for root, dirs, files in os.walk(directory):
                for file_name in files:
                    name_of_file, extension = os.path.splitext(file_name)
                    if extension == ".CSV":
                        df = pd.read_csv(
                            os.path.join(directory, name_of_file + ".CSV"))  # read the CSV file into a dataframe
                        for colName in df.columns:
                            if not str(colName).startswith("DVC"):
                                df = df.drop(colName, axis=1)  # remove unwanted columns
                        for column_name in df.columns:
                            if not 1 <= int(str(column_name).split('/')[0][4:]) <= 999999:
                                df = df.drop(column_name, axis=1)  # remove columns with DVC out of range
                        df = df.apply(lambda l: pd.Series(l.dropna().values))  # remove empty cells from each column
                        df = df.head(4)

                        if len(df.index) == 4:
                            for c in df.columns:
                                dvc_x = str(c).split('/')[0]  # gets DVC-###
                                if DVC.getAttribute("A") == dvc_x:
                                    det_x = str(c).split('/')[1]  # gets DET-###
                                    ddi = str(c).split('/')[2][4:]  # gets DDI
                                    for det in DVC.getElementsByTagName("DET"):
                                        if det.getAttribute("A") == det_x:
                                            for dor in det.getElementsByTagName("DOR"):
                                                for dpd in DVC.getElementsByTagName("DPD"):
                                                    if dpd.getAttribute("A") == dor.getAttribute("A") and \
                                                            dpd.getAttribute("B") == ddi:
                                                        dor.setAttribute("P094_dpd_value", str(df.loc[3].at[c]))
                                        for dor in det.getElementsByTagName("DOR"):
                                            if not dor.hasAttribute("P094_dpd_value"):
                                                dor.setAttribute("P094_dpd_value", "0")
                        else:
                            for det in DVC.getElementsByTagName("DET"):
                                for dor in det.getElementsByTagName("DOR"):
                                    dor.setAttribute("P094_dpd_value", '0')

            # set up name of file and where to save it
            save_file_here = os.path.join(directory,
                                          set_name_of_file(name, attribute_d) + ".XML")
            if os.path.isfile(save_file_here):
                save_file_here = os.path.join(directory, set_name_of_file(name, attribute_d) + f'({counter})' + ".XML")

            with open(save_file_here, 'w') as my_file2:
                my_file2.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
                my_file2.write(parent_element)
                my_file2.write("\n\t" + DVC.toxml())
                my_file2.write("\n</ISO11783_TaskData>")

    # filedialog will open to prompt the user to select the
    # TASKDATA.XML file. path of file will be saved in passed_path variable.
    passed_path = filedialog.askopenfilename(initialdir='/C', title='Select a file',
                                             filetypes=[("All files", "*.*")])
    while True:

        try:
            parsed_xml_file = minidom.parse(passed_path)

        except FileNotFoundError:
            messagebox.showinfo(title='warning', message='no file was selected. Application is exiting....')

        else:
            DVC_list = parsed_xml_file.getElementsByTagName("DVC")
            file_directory = os.path.dirname(passed_path)  # get the directory where TASKDATA.XML lives
            subprocess.run("CnhTc2Csv_GNSSDateTimeFixes.exe")
            extract_and_write(parsed_xml_file, DVC_list, file_directory)
            messagebox.showinfo(title='done', message='tool finished writing the files.')
        break


if __name__ == '__main__':
    main()
