import numpy as np
import pandas as pd
from datetime import datetime, timedelta
#import scipy
#from excel_parser import read_excel_tasks
#change
def setup_cpm(f):
    line = list()  # contains a single line
    tasks = dict()  # contains all the tasks
    number = -1
    for line in f:  # slide the file line by line
        singleElement = line  # split a line in subparts
        number += 1
        for i in range(len(singleElement)):  # creating the single task element
            tasks['task' + str(singleElement[0])] = dict()
            tasks['task' + str(singleElement[0])]['id'] = singleElement[0]
            tasks['task' + str(singleElement[0])]['name'] = singleElement[1]
            tasks['task' + str(singleElement[0])]['duration'] = singleElement[2]
            #if (singleElement[3] != "\n"):
            tasks['task' + str(singleElement[0])]['dependencies'] = singleElement[3]
            #else:
            #    tasks['task' + str(singleElement[0])]['dependencies'] = ['-1']            
            tasks['task' + str(singleElement[0])]['ES'] = 0
            if singleElement[4] != '':
               tasks['task' + str(singleElement[0])]['ES'] = singleElement[4]
            tasks['task' + str(singleElement[0])]['EF'] = 0
            tasks['task' + str(singleElement[0])]['LS'] = 0
            tasks['task' + str(singleElement[0])]['LF'] = 0
            tasks['task' + str(singleElement[0])]['float'] = 0
            tasks['task' + str(singleElement[0])]['isCritical'] = False

    return tasks

# =============================================================================
# FORWARD PASS
# =============================================================================
def forward_pass(tasks):
    for taskFW in tasks:  # slides all the tasks
        if (tasks[taskFW]['ES'] != 0):  # checks if it's an independent task
            #tasks[taskFW]['ES'] = 1
            tasks[taskFW]['EF'] = tasks[taskFW]['ES'] + timedelta(days=int(tasks[taskFW]['duration']))
        else:  # not the first task
            for k in tasks.keys():
                for dipendenza in tasks[k]['dependencies']:  # slides all the dependency in a single task
                    # print('task ' + taskFW + ' k '+ k + ' dipendenza ' +dipendenza)
                    if (len(tasks[k]['dependencies']) == 1):  # if the task k has only one dependency
                        tasks[k]['ES'] = tasks['task' + str(dipendenza)]['EF'] + timedelta(days=1)
                        tasks[k]['EF'] = tasks[k]['ES'] + timedelta(days=int(tasks[k]['duration'] - 1))
                    else: # if the task k has more dependency
                        if (tasks[k]['ES'] == 0 or tasks['task' + str(dipendenza)]['EF'] > tasks[k]['ES']):
                            tasks[k]['ES'] = tasks['task' + str(dipendenza)]['EF'] + timedelta(days=1)
                            tasks[k]['EF'] = tasks[k]['ES'] + timedelta(days=int(tasks[k]['duration'] - 1))

    aList = list()  # list of task keys
    for element in tasks.keys():
        aList.append(element)

    bList = list()  # reversed list of task keys
    while len(aList) > 0:
        bList.append(aList.pop())

    return aList,bList, tasks

# =============================================================================
# BACKWARD PASS
# =============================================================================
def back_pass(aList, bList, tasks):
  for taskBW in bList:
    if (bList.index(taskBW) == 0 or tasks[taskBW]['LF'] == 0):  # check if it's the last task (so no more task)
      tasks[taskBW]['LF'] = tasks[taskBW]['EF']
      tasks[taskBW]['LS'] = tasks[taskBW]['ES']

    for dipendenza in tasks[taskBW]['dependencies']:  # slides all the dependency in a single task
      if tasks['task' + str(dipendenza)]['dependencies'] == []:
        tasks['task' + str(dipendenza)]['LF'] = tasks['task' + str(dipendenza)]['EF']
        tasks['task' + str(dipendenza)]['LS'] = tasks['task' + str(dipendenza)]['ES']
        continue
      #if (dipendenza != []):  # check if it's NOT the last task
      if (tasks['task' + str(dipendenza)]['LF'] == 0):  # check if the the dependency is already analyzed
        # print('ID dipendenza: '+str(tasks['task'+dipendenza]['id']) + ' taskBW: '+str(tasks[taskBW]['id']))
        tasks['task' + str(dipendenza)]['LF'] = tasks[taskBW]['LS'] - timedelta(days=1)
        tasks['task' + str(dipendenza)]['LS'] = tasks['task' + str(dipendenza)]['LF'] - timedelta(days=int(tasks['task' + str(dipendenza)]['duration'] - 1))
        tasks['task' + str(dipendenza)]['float'] = tasks['task' + str(dipendenza)]['LF'] == tasks['task' + str(dipendenza)]['EF']
        # print('IF1 dip LS: '+str(tasks['task'+dipendenza]['LS']) +' dip LF: '+str(tasks['task'+dipendenza]['LF']) + ' taskBW: '+str(tasks[taskBW]['id'])+' taskBW ES '+ str(tasks[taskBW]['ES']))
      if (tasks['task' + str(dipendenza)]['LF'] > tasks[taskBW]['LS']):  # put the minimun value of LF for the dependencies of a task
        tasks['task' + str(dipendenza)]['LF'] = tasks[taskBW]['LS'] - timedelta(days=1)
        tasks['task' + str(dipendenza)]['LS'] = tasks['task' + str(dipendenza)]['LF'] - timedelta(days=int(tasks['task' + str(dipendenza)]['duration'] - 1))
        tasks['task' + str(dipendenza)]['float'] = tasks['task' + str(dipendenza)]['LF'] == tasks['task' + str(dipendenza)]['EF']
        # print('IF2 dip LS: '+str(tasks['task'+dipendenza]['LS']) +' dip LF: '+str(tasks['task'+dipendenza]['LF']) + ' taskBW: '+str(tasks[taskBW]['id']))
  
  
  return aList, bList, tasks
# =============================================================================
# PRINTING
# =============================================================================
def printing(tasks):
#('task id, task name, duration, ES, EF, LS, LF, float, isCritical')
  fin = ""
  for i, task in enumerate(tasks):
      if (tasks[task]['float'] == True):
          tasks[task]['isCritical'] = True

      fin += str(i+1) + ". Task id: " + str(tasks[task]['id']) + " with name " + str(tasks[task]['name']) + \
             " has a duration of " + str(tasks[task]['duration']) + \
             ". This task has an Early Start time of " + str(tasks[task]['ES']) + \
              ", an Early Finish time of " + str(tasks[task]['EF']) + \
              ", a Late Start Time of " + str(tasks[task]['LS']) + \
              ", a Late Finish Time of " + str(tasks[task]['LF']) + \
              ". This task's presence in the critical path is " + str(tasks[task]['isCritical']) + ". \n"

  return fin


def printcritical(tasks):
#('task id, task name, duration, ES, EF, LS, LF, float, isCritical')
  fin = "Critical Path tasks: "
  for i, task in enumerate(tasks):
      if (tasks[task]['float'] == 0):
          tasks[task]['isCritical'] = True

          fin += "#" + str(i+1) + ". Task id: " + str(tasks[task]['id']) + " with name: " + str(tasks[task]['name']) + ". "
          
              #  " has a duration of " + str(tasks[task]['duration']) + \
              #  ". This task has an Early Start time of " + str(tasks[task]['ES']) + \
              #    ", an Early Finish time of " + str(tasks[task]['EF']) + \
              #    ", a Late Start Time of " + str(tasks[task]['LS']) + \
              #    ", a Late Finish Time of " + str(tasks[task]['LF']) + \
              #    "."

  return fin


def cpmcalc(f):
    tasks = setup_cpm(f)
    aList, bList, tasks = forward_pass(tasks)
    aList, bList, tasks = back_pass(aList, bList, tasks)
    #return printcritical(tasks)
    return tasks



'''
project_tasks = read_excel_tasks("propel_beta\\project.xlsx")
critical_path = cpmcalc(project_tasks)
'''