import win32com.client
import csv

class MSProject(object):
    def __init__(self):
        self.mpp = win32com.client.Dispatch("MSProject.Application")
        self.Project = None
        self.Tasks = None

    def create_file(self, project_file):
        self.mpp.FileNew()
        self.Project = self.mpp.ActiveProject
        self.Tasks = self.Project.Tasks

    def save_file(self, project_file):
        self.mpp.FileSaveAs(project_file)

    def close_file(self):
        self.mpp.FileCloseAll()

    def convert_csv_to_project(self, csv_file, project_file):
        # read csv_file: level, name, duration
        tasks=[]
        with open(csv_file, 'r', encoding='utf-8') as f:
            f.readline()
            for row in csv.reader(f):
                tasks.append(row)

        self.create_file(project_file)
        for task in tasks:
            #print task
            level, name, duration = int(task[0]), task[1], int(task[2])
            new_task = self.Tasks.Add(name)
            new_task.OutlineLevel = level
            if duration != 0:
                new_task.Manual = True
                new_task.Estimated = False
            else:
                new_task.Manual = False
            new_task.Duration = duration*480 # one day 8 working hours = 480s
            #new_task.Start = None
            #new_task.Finish = None
        self.save_file(project_file)
        self.close_file()
        return