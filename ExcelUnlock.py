import os
import tempfile
import glob
import lxml.etree as ET
from zipfile import ZipFile
import shutil
import pyperclip

class Excel_Protection_Remover():

    def __init__(self):
        pass

    def unlock_excel_file(self,file_path):
        path_analysis = self.__anayse_path(file_path)
        td_extracted_components = self.__create_temporary_directory()
        temp_folder = td_extracted_components.name
        td_target = self.__create_temporary_directory()
        temp_folder_target = td_target.name
        self.__unzip_excel_file(file_path,temp_folder)
        self.__loop_through_components(temp_folder)
        self.__create_target_file(temp_folder, temp_folder_target, path_analysis)

    def __create_temporary_directory(self):
        temp_directory = tempfile.TemporaryDirectory()
        print(temp_directory.name)
        return temp_directory

    def __unzip_excel_file(self,file_path, temporary_location):
        with ZipFile(file_path, 'r') as zipObj:
            zipObj.extractall(temporary_location)

    def __remove_protection_tags(self,root, child, tree, file):
        root.remove(child)
        tree.write(file, xml_declaration=True,encoding='utf-8', method="xml")

    def __loop_through_components(self,temp_folder):
        files = glob.glob( temp_folder + '\\**\\*.*' ,recursive=True)
        for file in files:
            if os.path.splitext(file)[1] == '.xml':
                tree =  ET.parse(file)   
                root = tree.getroot()
                children = root.getchildren()

                for child in children:
                    if "sheetProtection" in child.tag or 'workbookProtection' in child.tag:
                        self.__remove_protection_tags(root, child, tree, file)

    def __create_target_file(self,temp_folder,temp_folder_target, path_analysis):
        target_file_name = path_analysis["file_name"] + "_UNLOCKED"
        working_file_path =  temp_folder_target + "\\" +  target_file_name
        shutil.make_archive(base_name=working_file_path,format="zip", root_dir=temp_folder)
        os.rename(working_file_path  + ".zip", working_file_path + path_analysis["file_extension"])
        shutil.copyfile(working_file_path + path_analysis["file_extension"],path_analysis["directory"] + "\\" + target_file_name + path_analysis["file_extension"])
        print("done")

    def __anayse_path(self,file_path):
        path_analysis={}
        path_analysis["directory"] = os.path.dirname(file_path)
        path_analysis["file_name"] = os.path.splitext(os.path.basename(file_path))[0]
        path_analysis["file_extension"] = os.path.splitext(os.path.basename(file_path))[1]
        return path_analysis

file_path =  pyperclip.paste()
exc = Excel_Protection_Remover()
exc.unlock_excel_file(file_path)

print("Done!!!") 