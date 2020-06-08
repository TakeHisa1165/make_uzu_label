import csv
import os

class Write_csv:
    def __init__(self):
        pass

    def write_csv(self, path_dict):
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        with open('path.csv', 'w', newline='') as csvfile:
            fieldnames = ["dir_path"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerow(path_dict)